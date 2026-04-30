import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, date, timedelta
import re
import math
import io
import pulp

st.set_page_config(page_title="Kế Hoạch SX – QDP", layout="wide")
st.title("🏭 Kế Hoạch Sản Xuất Nhà Máy QDP")

# ─── Sidebar ──────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Tham số nhà máy")
cap_monthly = st.sidebar.number_input(
    "Năng suất tối đa (m²/tháng)", value=10000, step=500, min_value=1000
)
horizon = st.sidebar.selectbox("Tầm nhìn kế hoạch (tháng)", [12, 18, 24], index=1)
show_weighted = st.sidebar.checkbox(
    "Tải trọng có xác suất",
    value=True,
    help="Tích: tính m² × xác suất (kỳ vọng). Bỏ tích: 100% khối lượng (xấu nhất).",
)

# ─── Helpers ──────────────────────────────────────────────────────────────────
def next_month(d):
    if d.month == 12:
        return datetime(d.year + 1, 1, 1)
    return datetime(d.year, d.month + 1, 1)

def months_range(start, end):
    """Danh sách datetime đầu tháng từ start đến end (inclusive)."""
    months = []
    cur = datetime(start.year, start.month, 1)
    end_m = datetime(end.year, end.month, 1)
    while cur <= end_m:
        months.append(cur)
        cur = next_month(cur)
    return months

def extract_prob(note):
    note = str(note).lower()
    if "đã ký hợp đồng" in note:
        return 100
    m = re.search(r"(\d+)\s*%", note)
    return int(m.group(1)) if m else 50

def status_label(prob):
    if prob == 100:
        return "✅ Đã ký HĐ"
    if prob >= 90:
        return "🟠 Khả năng cao (90%)"
    return "🟡 Đang xét (50%)"

def parse_date(val):
    """Chuyển Excel serial number hoặc datetime thành datetime(năm, tháng, 1)."""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        try:
            d = datetime(1899, 12, 30) + timedelta(days=int(val))
            return datetime(d.year, d.month, 1)
        except Exception:
            return None
    try:
        ts = pd.to_datetime(val)
        return datetime(ts.year, ts.month, 1)
    except Exception:
        return None

def detect_end_col(df):
    """Tự động tìm cột chứa ngày hoàn thành: thử col 4 trước, fallback sang col 5.
    Xử lý trường hợp file có cột trống (merged cell) xen giữa các cột ngày."""
    for _, row in df.iterrows():
        if len(row) < 5:
            continue
        try:
            float(row.iloc[2])  # chỉ xét hàng có m² hợp lệ
        except (ValueError, TypeError):
            continue
        if parse_date(row.iloc[4]) is not None:
            return 4  # cấu trúc 5 cột gọn
        if len(row) > 5 and parse_date(row.iloc[5]) is not None:
            return 5  # cấu trúc có cột trống giữa 2 ngày
    return 4

def parse_projects(df):
    # Đọc theo vị trí cột (0-indexed):
    #   col 1 = Tên dự án
    #   col 2 = Khối lượng (m²)
    #   col 3 = Ngày bắt đầu (serial hoặc datetime)
    #   col END = Ngày hoàn thành  (auto-detect: 4 hoặc 5)
    #   col END+1 = Ghi chú (xác suất)
    end_col  = detect_end_col(df)
    note_col = end_col + 1

    projects = []
    skip = {"nan", "", "TÊN DỰ ÁN", "GHI CHÚ:"}
    for _, row in df.iterrows():
        if len(row) <= end_col:
            continue
        name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if name in skip:
            continue
        try:
            m2 = float(row.iloc[2])
        except (ValueError, TypeError):
            continue
        start = parse_date(row.iloc[3])
        end   = parse_date(row.iloc[end_col])
        note  = str(row.iloc[note_col]).strip() if len(row) > note_col and pd.notna(row.iloc[note_col]) else ""
        if start is None or end is None:
            continue
        prob = extract_prob(note)
        projects.append({
            "name": name, "m2": m2,
            "start": start, "end": end,
            "prob": prob, "status": status_label(prob),
        })
    return projects

def capacity_status(pct):
    if pct >= 100: return "🔴 QUÁ TẢI"
    if pct >= 90:  return "🟠 Căng tiến độ"
    if pct >= 70:  return "🟡 Cần chú ý"
    return "🟢 Bình thường"

def row_style(row):
    pct = row["% Sử dụng"]
    if pct >= 100: return ["background-color:#ff4d4d;color:white"]  * len(row)
    if pct >= 90:  return ["background-color:#ff8c00;color:white"]  * len(row)
    if pct >= 70:  return ["background-color:#ffd700"]              * len(row)
    return [""] * len(row)

# ─── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📊 Bảng Tổng Hợp Tải Trọng", "🔧 Tối Ưu Kế Hoạch SX (LP)"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Planning Dashboard
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    uploaded = st.file_uploader(
        "📂 Tải lên file Excel (danh sách dự án)", type=["xlsx", "xls"], key="plan_file"
    )

    if uploaded is None:
        st.info("⬅️ Vui lòng tải lên file Excel danh sách dự án để bắt đầu.")
    else:
        xl = pd.ExcelFile(uploaded)
        sheet = st.selectbox("Chọn sheet dữ liệu", xl.sheet_names, key="plan_sheet")
        try:
            df_raw = xl.parse(sheet, header=None)
        except Exception as e:
            st.error(f"Lỗi đọc sheet: {e}")
            df_raw = None

        if df_raw is not None:
            projects = parse_projects(df_raw)

            if not projects:
                st.error("Không parse được dự án nào. Kiểm tra lại file.")
            else:
                # ── Session state: khởi tạo khi file thay đổi ────────────
                file_id = f"{uploaded.name}_{uploaded.size}_{sheet}"
                if st.session_state.get("t1_file_id") != file_id:
                    orig_rows = [{
                        "name":  p["name"],
                        "m2":    p["m2"],
                        "start": p["start"].date(),
                        "end":   p["end"].date(),
                        "prob":  p["prob"],
                    } for p in projects]
                    st.session_state.t1_file_id    = file_id
                    st.session_state.t1_orig       = orig_rows
                    st.session_state.t1_rows       = [dict(r) for r in orig_rows]
                    st.session_state.t1_editor_ver = st.session_state.get("t1_editor_ver", 0) + 1
                    st.session_state.t1_reset_ask  = False

                rows     = st.session_state.t1_rows
                orig_map = {r["name"]: r for r in st.session_state.t1_orig}

                # ── Toolbar ───────────────────────────────────────────────
                st.subheader("⚙️ Quản lý danh sách dự án")
                tb1, tb2, tb3 = st.columns([1, 1, 4])

                with tb1:
                    exp_rows = [{
                        "Tên dự án":       r["name"],
                        "Khối lượng (m²)": r["m2"],
                        "Bắt đầu":         r["start"].strftime("%m/%Y"),
                        "Kết thúc":        r["end"].strftime("%m/%Y"),
                        "XS%":             r["prob"],
                    } for r in rows]
                    buf = io.BytesIO()
                    pd.DataFrame(exp_rows).to_excel(buf, index=False, engine="openpyxl")
                    st.download_button(
                        "💾 Xuất Excel",
                        data=buf.getvalue(),
                        file_name="du_an_updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="t1_export",
                    )

                with tb2:
                    if not st.session_state.get("t1_reset_ask"):
                        if st.button("↩ Reset tất cả", key="t1_reset_btn"):
                            st.session_state.t1_reset_ask = True
                            st.rerun()
                    else:
                        st.warning("Xác nhận reset?")
                        rc1, rc2 = st.columns(2)
                        with rc1:
                            if st.button("✅ Xác nhận", key="t1_reset_yes"):
                                st.session_state.t1_rows = [
                                    dict(r) for r in st.session_state.t1_orig
                                ]
                                st.session_state.t1_editor_ver += 1
                                st.session_state.t1_reset_ask = False
                                st.rerun()
                        with rc2:
                            if st.button("❌ Huỷ", key="t1_reset_no"):
                                st.session_state.t1_reset_ask = False
                                st.rerun()

                with tb3:
                    n_new = sum(1 for r in rows if r["name"] not in orig_map)
                    n_mod = sum(
                        1 for r in rows
                        if r["name"] in orig_map and r != orig_map[r["name"]]
                    )
                    parts = []
                    if n_new: parts.append(f"🆕 {n_new} dự án mới")
                    if n_mod: parts.append(f"✏️ {n_mod} đã chỉnh sửa")
                    st.caption(
                        ("  |  ".join(parts)) if parts
                        else "Dữ liệu đồng bộ với file Excel gốc."
                    )

                # ── Hàm tính trạng thái dòng ──────────────────────────────
                def _row_tag(r):
                    if r["name"] not in orig_map:
                        return "🆕"
                    if r != orig_map[r["name"]]:
                        return "✏️"
                    return ""

                # ── Bảng editable đầy đủ ─────────────────────────────────
                edit_df = pd.DataFrame([{
                    "Tên dự án":       r["name"],
                    "Khối lượng (m²)": r["m2"],
                    "Bắt đầu":         r["start"],
                    "Kết thúc":        r["end"],
                    "XS%":             r["prob"],
                    "Trạng thái":      _row_tag(r),
                } for r in rows])

                edited = st.data_editor(
                    edit_df,
                    column_config={
                        "Tên dự án":       st.column_config.TextColumn(
                            "Tên dự án", required=True
                        ),
                        "Khối lượng (m²)": st.column_config.NumberColumn(
                            "Khối lượng (m²)", min_value=1, format="%,.0f"
                        ),
                        "Bắt đầu":         st.column_config.DateColumn(
                            "Bắt đầu SX", format="MM/YYYY",
                            help="Chọn ngày bất kỳ trong tháng bắt đầu"
                        ),
                        "Kết thúc":        st.column_config.DateColumn(
                            "Kết thúc SX", format="MM/YYYY",
                            help="Chọn ngày bất kỳ trong tháng kết thúc"
                        ),
                        "XS%":             st.column_config.NumberColumn(
                            "XS%", min_value=0, max_value=100, step=5,
                            help="0=loại | 50=đang xét | 90=khả năng cao | 100=đã ký HĐ",
                        ),
                        "Trạng thái":      st.column_config.TextColumn(
                            "✏️", disabled=True, width="small"
                        ),
                    },
                    num_rows="dynamic",
                    use_container_width=True,
                    hide_index=True,
                    key=f"t1_editor_{st.session_state.get('t1_editor_ver', 0)}",
                )

                # ── Validate + sync về session state ─────────────────────
                def _to_date(val):
                    if pd.isna(val):
                        return None
                    if isinstance(val, (datetime,)):
                        return date(val.year, val.month, 1)
                    if isinstance(val, pd.Timestamp):
                        return date(val.year, val.month, 1)
                    if isinstance(val, date):
                        return date(val.year, val.month, 1)
                    return None

                errs     = []
                new_rows = []
                for _, row in edited.iterrows():
                    nm = str(row.get("Tên dự án", "")).strip() if pd.notna(row.get("Tên dự án")) else ""
                    if not nm:
                        continue
                    try:
                        m2v = round(float(row["Khối lượng (m²)"]), 4)  # round để tránh float drift
                        if m2v < 1:
                            raise ValueError
                    except (TypeError, ValueError):
                        errs.append(f"**{nm}**: Khối lượng phải ≥ 1 m²")
                        continue
                    sd = _to_date(row["Bắt đầu"])
                    ed = _to_date(row["Kết thúc"])
                    if sd is None:
                        errs.append(f"**{nm}**: Ngày bắt đầu không hợp lệ")
                        continue
                    if ed is None:
                        errs.append(f"**{nm}**: Ngày kết thúc không hợp lệ")
                        continue
                    if ed < sd:
                        errs.append(
                            f"**{nm}**: Kết thúc ({ed.strftime('%m/%Y')}) "
                            f"phải sau Bắt đầu ({sd.strftime('%m/%Y')})"
                        )
                        continue
                    try:
                        prv = int(row["XS%"])           # lưu dạng int, không phải float hay string
                        prv = max(0, min(100, prv))
                    except (TypeError, ValueError):
                        prv = 100
                    new_rows.append({"name": nm, "m2": m2v, "start": sd, "end": ed, "prob": prv})

                if errs:
                    for e in errs:
                        st.error(f"⚠️ {e}")

                # ── Detect deletions ──────────────────────────────────────
                edited_names = {
                    str(row.get("Tên dự án", "")).strip()
                    for _, row in edited.iterrows()
                    if pd.notna(row.get("Tên dự án"))
                    and str(row.get("Tên dự án", "")).strip()
                }
                prev_names = {r["name"] for r in st.session_state.t1_rows}
                del_names  = [nm for nm in prev_names if nm not in edited_names]

                if st.session_state.get("t1_delete_confirm"):
                    # ── Confirm dialog ────────────────────────────────────
                    pending = st.session_state.t1_pending_deletes
                    label   = (
                        f"dự án **{pending[0]}**" if len(pending) == 1
                        else f"**{len(pending)}** dự án: " + ", ".join(f"**{n}**" for n in pending)
                    )
                    st.warning(f"Bạn có chắc muốn xóa {label}?")
                    dc1, dc2, _ = st.columns([1, 1, 4])
                    with dc1:
                        if st.button("✅ Xác nhận xóa", key="t1_del_yes"):
                            st.session_state.t1_rows           = st.session_state.t1_pending_rows
                            st.session_state.t1_editor_ver    += 1
                            st.session_state.t1_delete_confirm = False
                            st.session_state.pop("t1_pending_deletes", None)
                            st.session_state.pop("t1_pending_rows", None)
                            st.rerun()
                    with dc2:
                        if st.button("❌ Huỷ xóa", key="t1_del_no"):
                            st.session_state.t1_editor_ver    += 1  # force reload → row restored
                            st.session_state.t1_delete_confirm = False
                            st.session_state.pop("t1_pending_deletes", None)
                            st.session_state.pop("t1_pending_rows", None)
                            st.rerun()

                elif del_names:
                    # New deletion — hold for confirmation before syncing
                    st.session_state.t1_pending_deletes = del_names
                    st.session_state.t1_pending_rows    = new_rows
                    st.session_state.t1_delete_confirm  = True
                    st.rerun()

                elif new_rows and new_rows != st.session_state.t1_rows:
                    st.session_state.t1_rows = new_rows
                    st.rerun()

                # ── Xây projects từ session state (nguồn dữ liệu duy nhất) ─
                projects = [{
                    "name":   r["name"],
                    "m2":     r["m2"],
                    "start":  datetime(r["start"].year, r["start"].month, 1),
                    "end":    datetime(r["end"].year, r["end"].month, 1),
                    "prob":   r["prob"],
                    "status": status_label(r["prob"]),
                } for r in st.session_state.t1_rows]

                if not projects:
                    st.warning("Bảng dự án đang trống.")
                    st.stop()

                # ── Xem thay đổi với màu highlight ───────────────────────
                with st.expander("🎨 Xem bảng thay đổi (màu highlight)", expanded=False):
                    st.markdown(
                        "🟡 **Vàng** = đã chỉnh sửa so với file gốc &nbsp;|&nbsp; "
                        "🟢 **Xanh lá** = dòng mới thêm"
                    )

                    def _highlight_rows(df_row):
                        nm = df_row["Tên dự án"]
                        if nm not in orig_map:
                            return ["background-color:#d4edda; color:#155724"] * len(df_row)
                        orig = orig_map[nm]
                        changed = (
                            df_row["Khối lượng (m²)"] != orig["m2"]
                            or df_row["Bắt đầu"] != orig["start"].strftime("%m/%Y")
                            or df_row["Kết thúc"] != orig["end"].strftime("%m/%Y")
                            or df_row["XS%"] != orig["prob"]
                        )
                        if changed:
                            return ["background-color:#fff3cd; color:#856404"] * len(df_row)
                        return [""] * len(df_row)

                    display_df = pd.DataFrame([{
                        "Tên dự án":       r["name"],
                        "Khối lượng (m²)": r["m2"],
                        "Bắt đầu":         r["start"].strftime("%m/%Y"),
                        "Kết thúc":        r["end"].strftime("%m/%Y"),
                        "XS%":             r["prob"],
                        "Trạng thái":      _row_tag(r),
                    } for r in st.session_state.t1_rows])

                    st.dataframe(
                        display_df.style.apply(_highlight_rows, axis=1),
                        use_container_width=True,
                        hide_index=True,
                    )

                st.divider()

                # ── Horizon months ──────────────────────────────────────────
                today = datetime(datetime.today().year, datetime.today().month, 1)
                horizon_months = []
                cur = today
                for _ in range(horizon):
                    horizon_months.append(cur)
                    cur = next_month(cur)
                month_labels = [m.strftime("%m/%Y") for m in horizon_months]

                # ── Load matrix: load_data[month_label][project_name] = m² ──
                # Xây đồng thời 2 ma trận: kỳ vọng (×XS%) và xấu nhất (100%)
                load_data_w    = {ml: {} for ml in month_labels}  # kỳ vọng
                load_data_full = {ml: {} for ml in month_labels}  # xấu nhất
                for p in projects:
                    proj_months = months_range(p["start"], p["end"])
                    n = len(proj_months)
                    if n == 0:
                        continue
                    for pm in proj_months:
                        ml = pm.strftime("%m/%Y")
                        if ml in load_data_w:
                            load_data_w[ml][p["name"]]    = round(p["m2"] * (p["prob"] / 100) / n, 1)
                            load_data_full[ml][p["name"]] = round(p["m2"] / n, 1)

                # Ma trận dùng cho biểu đồ/bảng tổng hợp theo lựa chọn sidebar
                load_data = load_data_w if show_weighted else load_data_full

                # ── Monthly summary ──────────────────────────────────────────
                summary_rows = []
                for ml in month_labels:
                    total = sum(load_data[ml].values())
                    pct   = round(total / cap_monthly * 100, 1)
                    summary_rows.append({
                        "Tháng":              ml,
                        "Tổng tải (m²)":      round(total),
                        "Năng suất max (m²)": cap_monthly,
                        "% Sử dụng":          pct,
                        "Trạng thái":         capacity_status(pct),
                        "Dư / Thiếu (m²)":    round(cap_monthly - total),
                    })
                df_summary = pd.DataFrame(summary_rows)

                # ── KPI cards ───────────────────────────────────────────────
                n_overload = int((df_summary["% Sử dụng"] >= 100).sum())
                n_tight    = int(((df_summary["% Sử dụng"] >= 90) & (df_summary["% Sử dụng"] < 100)).sum())
                n_warn     = int(((df_summary["% Sử dụng"] >= 70) & (df_summary["% Sử dụng"] < 90)).sum())
                n_ok       = int((df_summary["% Sử dụng"] < 70).sum())

                k1, k2, k3, k4 = st.columns(4)
                k1.metric("🔴 Tháng quá tải",         f"{n_overload} tháng")
                k2.metric("🟠 Căng tiến độ (90–100%)", f"{n_tight} tháng")
                k3.metric("🟡 Cần chú ý (70–90%)",    f"{n_warn} tháng")
                k4.metric("🟢 Bình thường (<70%)",     f"{n_ok} tháng")

                st.divider()

                # ── Danh sách dự án ─────────────────────────────────────────
                with st.expander(f"📋 Danh sách {len(projects)} dự án", expanded=True):
                    df_projects = pd.DataFrame([{
                        "Dự án":          p["name"],
                        "Khối lượng (m²)": f"{p['m2']:,.0f}",
                        "Bắt đầu SX":     p["start"].strftime("%m/%Y"),
                        "Kết thúc SX":    p["end"].strftime("%m/%Y"),
                        "Số tháng SX":    len(months_range(p["start"], p["end"])),
                        "m²/tháng (tb)":  f"{p['m2'] / max(1, len(months_range(p['start'], p['end']))):.0f}",
                        "Trạng thái":     p["status"],
                        "Xác suất":       f"{p['prob']}%",
                    } for p in projects])
                    st.dataframe(df_projects, use_container_width=True, hide_index=True)

                # ── Stacked bar chart ────────────────────────────────────────
                horizon_label = f"{month_labels[0]} – {month_labels[-1]}"
                st.subheader(f"📊 Tải trọng nhà máy: {horizon_label} ({horizon} tháng)")

                COLORS = (
                    px.colors.qualitative.Plotly
                    + px.colors.qualitative.Set2
                    + px.colors.qualitative.Set3
                )

                y_max = max(df_summary["Tổng tải (m²)"].max() * 1.15, cap_monthly * 1.15)

                fig = go.Figure()
                fig.add_hrect(y0=0,                y1=cap_monthly * 0.7,
                              fillcolor="rgba(0,180,0,0.06)",   line_width=0)
                fig.add_hrect(y0=cap_monthly * 0.7, y1=cap_monthly * 0.9,
                              fillcolor="rgba(255,200,0,0.12)", line_width=0)
                fig.add_hrect(y0=cap_monthly * 0.9, y1=y_max,
                              fillcolor="rgba(255,60,0,0.09)",  line_width=0)

                for i, p in enumerate(projects):
                    y_vals = [load_data[ml].get(p["name"], 0) for ml in month_labels]
                    if any(v > 0 for v in y_vals):
                        fig.add_trace(go.Bar(
                            name=p["name"],
                            x=month_labels,
                            y=y_vals,
                            marker_color=COLORS[i % len(COLORS)],
                            hovertemplate="<b>%{x}</b><br>" + p["name"] + ": %{y:,.0f} m²<extra></extra>",
                        ))

                fig.add_hline(
                    y=cap_monthly, line_dash="dash", line_color="red", line_width=2,
                    annotation_text=f"  Tối đa: {cap_monthly:,} m²",
                    annotation_font_color="red", annotation_position="right",
                )
                fig.add_hline(
                    y=cap_monthly * 0.9, line_dash="dot", line_color="darkorange", line_width=1.5,
                    annotation_text=f"  Cảnh báo 90%: {int(cap_monthly * 0.9):,} m²",
                    annotation_font_color="darkorange", annotation_position="right",
                )

                fig.update_layout(
                    barmode="stack",
                    height=560,
                    legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
                    xaxis_title="Tháng",
                    yaxis_title="Khối lượng sản xuất (m²)",
                    yaxis=dict(range=[0, y_max]),
                    margin=dict(t=20, r=180, b=60, l=70),
                )
                st.plotly_chart(fig, use_container_width=True)

                # ── Gantt Chart ──────────────────────────────────────────────
                st.subheader("📅 Gantt — Tiến độ sản xuất các dự án")

                # ── Nhóm màu theo XS% ────────────────────────────────────────
                def _xs_group(prob):
                    if prob == 100:
                        return "100% – Đã ký HĐ"
                    elif prob >= 90:
                        return "90% – Khả năng cao"
                    elif prob >= 50:
                        return "50% – Đang xét"
                    else:
                        return "<50% – Xem xét"

                gantt_color_map = {
                    "100% – Đã ký HĐ":    "#1f77b4",   # xanh dương
                    "90% – Khả năng cao": "#2ca02c",   # xanh lá
                    "50% – Đang xét":     "#ff7f0e",   # cam
                    "<50% – Xem xét":     "#aaaaaa",   # xám
                }

                gantt_rows = []
                for p in projects:
                    # Kéo dài thanh đến hết tháng kết thúc
                    end_display = datetime(p["end"].year, p["end"].month, 1)
                    if end_display.month == 12:
                        end_display = datetime(end_display.year + 1, 1, 1)
                    else:
                        end_display = datetime(end_display.year, end_display.month + 1, 1)
                    n_months = len(months_range(p["start"], p["end"]))
                    gantt_rows.append({
                        "Dự án":        p["name"],
                        "Bắt đầu":      p["start"],
                        "Kết thúc":     end_display,
                        "Nhóm XS":      _xs_group(p["prob"]),
                        "m²":           p["m2"],
                        "XS%":          p["prob"],
                        "Bắt đầu SX":   p["start"].strftime("%m/%Y"),
                        "Kết thúc SX":  datetime(p["end"].year, p["end"].month, 1).strftime("%m/%Y"),
                    })

                df_gantt = pd.DataFrame(gantt_rows).sort_values("Bắt đầu")

                fig_gantt1 = px.timeline(
                    df_gantt,
                    x_start="Bắt đầu",
                    x_end="Kết thúc",
                    y="Dự án",
                    color="Nhóm XS",
                    color_discrete_map=gantt_color_map,
                    custom_data=["m²", "XS%", "Bắt đầu SX", "Kết thúc SX"],
                    labels={"Dự án": ""},
                    category_orders={"Nhóm XS": [
                        "100% – Đã ký HĐ",
                        "90% – Khả năng cao",
                        "50% – Đang xét",
                        "<50% – Xem xét",
                    ]},
                )

                # Hover tooltip tuỳ chỉnh
                fig_gantt1.update_traces(
                    hovertemplate=(
                        "<b>%{y}</b><br>"
                        "Khối lượng: %{customdata[0]:,.0f} m²<br>"
                        "Xác suất: %{customdata[1]}%<br>"
                        "Bắt đầu: %{customdata[2]}<br>"
                        "Kết thúc: %{customdata[3]}"
                        "<extra></extra>"
                    )
                )

                # Đường dọc đỏ = tháng hiện tại
                # px.timeline dùng ms epoch nội bộ → phải truyền timestamp*1000
                fig_gantt1.add_vline(
                    x=today.timestamp() * 1000,
                    line_color="red",
                    line_width=2,
                    line_dash="dash",
                    annotation_text="  Tháng hiện tại",
                    annotation_font_color="red",
                    annotation_position="top right",
                    annotation_font_size=11,
                )

                # Đồng bộ trục X với biểu đồ cột (cùng khoảng horizon)
                x_range_start = horizon_months[0].strftime("%Y-%m-%d")
                x_range_end   = next_month(horizon_months[-1]).strftime("%Y-%m-%d")

                fig_gantt1.update_yaxes(autorange="reversed", tickfont=dict(size=10))
                fig_gantt1.update_xaxes(
                    range=[x_range_start, x_range_end],
                    dtick="M1",
                    tickformat="%m/%Y",
                    tickangle=45,
                    tickfont=dict(size=9),
                )
                fig_gantt1.update_layout(
                    height=max(400, len(projects) * 30 + 150),
                    margin=dict(l=10, r=20, t=40, b=60),
                    legend=dict(
                        title="Xác suất",
                        orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0,
                    ),
                    plot_bgcolor="white",
                )
                st.plotly_chart(fig_gantt1, use_container_width=True)
                st.caption(
                    "🔵 Đã ký HĐ (100%) · 🟢 Khả năng cao (90%) · "
                    "🟠 Đang xét (50%) · ⬜ Xem xét (<50%) · "
                    "🔴 Đường đỏ = tháng hiện tại"
                )

                st.divider()

                # ── Bảng tổng hợp theo tháng ────────────────────────────────
                st.subheader("📋 Bảng tổng hợp tải trọng theo tháng")
                st.dataframe(
                    df_summary.style.apply(row_style, axis=1),
                    use_container_width=True, hide_index=True,
                )

                # ── Chi tiết dự án × tháng ──────────────────────────────────
                with st.expander("🗂️ Chi tiết tải trọng từng dự án × từng tháng"):
                    detail_rows = []
                    for p in projects:
                        row = {
                            "Dự án":      p["name"],
                            "Trạng thái": p["status"],
                            "Tổng m²":    f"{p['m2']:,.0f}",
                        }
                        for ml in month_labels:
                            v = load_data[ml].get(p["name"], 0)
                            row[ml] = f"{v:,.0f}" if v > 0 else ""
                        detail_rows.append(row)
                    # Hàng tổng
                    total_row = {
                        "Dự án": "📊 TỔNG",
                        "Trạng thái": "",
                        "Tổng m²": f"{sum(p['m2'] for p in projects):,.0f}",
                    }
                    for ml in month_labels:
                        t = sum(load_data[ml].values())
                        total_row[ml] = f"{t:,.0f}" if t > 0 else ""
                    detail_rows.append(total_row)
                    st.dataframe(
                        pd.DataFrame(detail_rows),
                        use_container_width=True, hide_index=True,
                    )

                # ── What-If Simulator ────────────────────────────────────────
                st.subheader("🔍 Kiểm tra dự án mới (What-If Simulator)")
                with st.expander(
                    "➕ Nhập thông tin dự án mới để kiểm tra năng lực nhà máy", expanded=True
                ):
                    sim_scenario = st.radio(
                        "Kịch bản tải nền (dự án hiện có):",
                        [
                            "Kỳ vọng — dùng xác suất thực tế (mặc định)",
                            "Xấu nhất — 100% tất cả dự án",
                        ],
                        index=0,
                        horizontal=True,
                        key="sim_scenario",
                    )
                    sim_load = (
                        load_data_w if sim_scenario.startswith("Kỳ") else load_data_full
                    )
                    sim_label = "Tải nền kỳ vọng (m²)" if sim_scenario.startswith("Kỳ") else "Tải nền xấu nhất (m²)"

                    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
                    with c1:
                        new_name = st.text_input("Tên dự án", value="Dự án mới")
                    with c2:
                        new_m2 = st.number_input("Khối lượng (m²)", value=5000, step=500, min_value=0)
                    with c3:
                        new_start_d = st.date_input("Tháng bắt đầu SX", value=date.today())
                    with c4:
                        em = date.today().month + 3
                        ey = date.today().year + (em - 1) // 12
                        em = (em - 1) % 12 + 1
                        new_end_d = st.date_input("Tháng kết thúc SX", value=date(ey, em, 1))

                    if st.button("▶ Kiểm tra tác động", type="primary"):
                        new_start_dt = datetime(new_start_d.year, new_start_d.month, 1)
                        new_end_dt   = datetime(new_end_d.year,   new_end_d.month,   1)

                        if new_end_dt < new_start_dt:
                            st.error("Tháng kết thúc phải sau tháng bắt đầu.")
                        elif new_m2 <= 0:
                            st.error("Khối lượng phải lớn hơn 0.")
                        else:
                            new_proj_months = months_range(new_start_dt, new_end_dt)
                            new_load_pm     = new_m2 / len(new_proj_months)

                            pre_overload = []   # tháng đã quá tải TRƯỚC khi thêm dự án mới
                            new_overload = []   # tháng chỉ quá tải SAU khi thêm
                            tight        = []   # tháng căng 90-100% sau khi thêm
                            sim_rows     = []

                            for ml, m_dt in zip(month_labels, horizon_months):
                                existing = sum(sim_load[ml].values())
                                added    = round(new_load_pm, 1) if m_dt in new_proj_months else 0
                                total    = existing + added
                                pct      = round(total / cap_monthly * 100, 1)

                                if existing > cap_monthly:
                                    pre_overload.append(ml)
                                    status = "⚠️ Đã quá tải từ trước"
                                elif total > cap_monthly:
                                    new_overload.append((ml, round(total - cap_monthly)))
                                    status = "🔴 Quá tải do dự án mới"
                                elif pct >= 90:
                                    tight.append((ml, pct))
                                    status = "🟠 Căng tiến độ"
                                elif pct >= 70:
                                    status = "🟡 Cần chú ý"
                                else:
                                    status = "🟢 Bình thường"

                                sim_rows.append({
                                    "Tháng":            ml,
                                    sim_label:          round(existing),
                                    f"{new_name} (m²)": round(added),
                                    "Tổng (m²)":        round(total),
                                    "% Sử dụng":        pct,
                                    "Trạng thái":       status,
                                })

                            # ── Thông báo ──────────────────────────────────
                            if pre_overload:
                                st.warning(
                                    "⚠️ **Các tháng sau đã quá tải từ trước** "
                                    f"(không liên quan dự án mới): "
                                    f"{', '.join(pre_overload)}"
                                )

                            if new_overload:
                                st.error(
                                    f"🔴 **KHÔNG NÊN NHẬN** — Dự án này gây quá tải thêm "
                                    f"**{len(new_overload)}** tháng:"
                                )
                                for ml, excess in new_overload:
                                    st.write(
                                        f"  • Tháng {ml}: vượt **{excess:,} m²** "
                                        f"so với năng suất tối đa"
                                    )
                            else:
                                if tight:
                                    st.warning(
                                        f"⚡ **CÂN NHẮC KỸ** — Thêm dự án sẽ căng tiến độ "
                                        f"(90–100%) tại **{len(tight)}** tháng."
                                    )
                                st.success(
                                    "✅ **CÓ THỂ NHẬN** — Dự án này không tạo thêm tháng quá tải mới."
                                )

                            df_sim = pd.DataFrame(sim_rows)
                            st.dataframe(
                                df_sim.style.apply(row_style, axis=1),
                                use_container_width=True, hide_index=True,
                            )

                # ── Download ────────────────────────────────────────────────
                csv = df_summary.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "📥 Tải bảng tổng hợp (CSV)",
                    data=csv, file_name="ke_hoach_sx.csv", mime="text/csv",
                )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Lịch sản xuất tối ưu theo dự án (LP)
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Lịch sản xuất tối ưu theo dự án (LP)")
    st.caption(
        "LP tìm lịch sản xuất m²/tháng cho từng dự án trong khung thời gian của nó, "
        "san đều tải trọng tháng, không vượt năng suất nhà máy."
    )

    uploaded_lp = st.file_uploader(
        "📂 Tải lên file Excel (danh sách dự án)", type=["xlsx", "xls"], key="lp_file"
    )

    if uploaded_lp is None:
        st.info("⬅️ Tải lên file Excel để chạy phân bổ LP.")
    else:
        xl_lp = pd.ExcelFile(uploaded_lp)
        sheet_lp = st.selectbox("Chọn sheet dữ liệu", xl_lp.sheet_names, key="lp_sheet")
        try:
            df_lp_raw = xl_lp.parse(sheet_lp, header=None)
        except Exception as e:
            st.error(f"Lỗi đọc sheet: {e}")
            df_lp_raw = None

        if df_lp_raw is not None:
            lp_cap = st.number_input("Năng suất tối đa (m²/tháng)", value=10000, step=500, key="lp_cap")
            lp_weighted = st.checkbox(
                "Dùng trọng số xác suất (LP)",
                value=False,
                help="Tích: LP phân bổ m² × xác suất (kỳ vọng). Bỏ tích: LP phân bổ 100% m² (worst case – mọi dự án đều ký).",
                key="lp_weighted",
            )

            end_col_lp  = detect_end_col(df_lp_raw)
            note_col_lp = end_col_lp + 1
            skip_lp     = {"nan", "", "TÊN DỰ ÁN", "GHI CHÚ:"}

            contracts = []
            for _, row in df_lp_raw.iterrows():
                if len(row) <= end_col_lp:
                    continue
                cname = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                if cname in skip_lp:
                    continue
                try:
                    m2_q = float(row.iloc[2])
                except (ValueError, TypeError):
                    continue
                start = parse_date(row.iloc[3])
                end   = parse_date(row.iloc[end_col_lp])
                if start is None or end is None:
                    continue
                note = (
                    str(row.iloc[note_col_lp]).strip()
                    if len(row) > note_col_lp and pd.notna(row.iloc[note_col_lp])
                    else ""
                )
                prob   = extract_prob(note)
                eff_m2 = m2_q * (prob / 100) if lp_weighted else m2_q
                contracts.append({
                    "name": cname, "m2": m2_q, "eff_m2": eff_m2,
                    "prob": prob, "start": start, "end": end,
                })

            if not contracts:
                st.error("Không có dự án nào được parse.")
            else:
                # ── Bảng chỉnh sửa xác suất trước khi chạy LP ───────────────
                st.markdown("**Xác suất dự án** — chỉnh sửa cột XS% trực tiếp, sau đó nhấn **Chạy LP**:")
                edit_df = pd.DataFrame([{
                    "STT":           i + 1,
                    "Dự án":         p["name"],
                    "m² danh nghĩa": p["m2"],
                    "XS%":           p["prob"],
                    "Bắt đầu":       p["start"].strftime("%m/%Y"),
                    "Kết thúc":      p["end"].strftime("%m/%Y"),
                } for i, p in enumerate(contracts)])

                edited_df = st.data_editor(
                    edit_df,
                    column_config={
                        "STT": st.column_config.NumberColumn(disabled=True, width="small"),
                        "Dự án": st.column_config.TextColumn(disabled=True),
                        "m² danh nghĩa": st.column_config.NumberColumn(
                            format="%,.0f", disabled=True
                        ),
                        "XS%": st.column_config.NumberColumn(
                            "XS% (chỉnh sửa)",
                            min_value=0, max_value=100, step=5,
                            help="0 = loại khỏi lịch | 50 = đang xét | 90 = khả năng cao | 100 = đã ký HĐ",
                        ),
                        "Bắt đầu": st.column_config.TextColumn(disabled=True),
                        "Kết thúc": st.column_config.TextColumn(disabled=True),
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="lp_edit_prob",
                )

                run_lp = st.button("▶ Chạy LP", type="primary", key="run_lp_btn")

                if run_lp:
                    # Cập nhật prob và eff_m2 từ bảng đã chỉnh sửa
                    for i, p in enumerate(contracts):
                        p["prob"]   = int(edited_df.iloc[i]["XS%"])
                        p["eff_m2"] = p["m2"] * (p["prob"] / 100) if lp_weighted else p["m2"]

                    # Chỉ đưa vào LP các dự án có eff_m2 > 0
                    active = [p for p in contracts if p["eff_m2"] > 0]
                    if not active:
                        st.error("Không có dự án nào có khối lượng > 0 để lên lịch.")
                        st.stop()

                    all_d   = [p["start"] for p in active] + [p["end"] for p in active]
                    lp_mons = months_range(min(all_d), max(all_d))
                    midx    = {m: i for i, m in enumerate(lp_mons)}
                    N       = len(lp_mons)
                    month_labels_lp = [m.strftime("%m/%Y") for m in lp_mons]

                    proj_m_idxs = {}
                    for p_idx, p in enumerate(active):
                        proj_m_idxs[p_idx] = [
                            midx[m] for m in months_range(p["start"], p["end"]) if m in midx
                        ]

                    x = {}
                    for p_idx in range(len(active)):
                        for m_idx in proj_m_idxs[p_idx]:
                            x[(p_idx, m_idx)] = pulp.LpVariable(f"x_{p_idx}_{m_idx}", lowBound=0)

                    z = pulp.LpVariable("z", lowBound=0)
                    lp_prob = pulp.LpProblem("Schedule", pulp.LpMinimize)
                    lp_prob += z

                    for m_idx in range(N):
                        monthly = pulp.lpSum(
                            x[(p_idx, m_idx)]
                            for p_idx in range(len(active))
                            if (p_idx, m_idx) in x
                        )
                        lp_prob += monthly <= lp_cap
                        lp_prob += monthly <= z

                    for p_idx, p in enumerate(active):
                        lp_prob += (
                            pulp.lpSum(x[(p_idx, m_idx)] for m_idx in proj_m_idxs[p_idx])
                            == p["eff_m2"]
                        )

                    with st.spinner("Đang giải LP..."):
                        lp_prob.solve(pulp.PULP_CBC_CMD(msg=False))

                    if lp_prob.status != 1:
                        st.error(
                            f"⚠️ Không tìm được lời giải khả thi với năng suất "
                            f"**{lp_cap:,} m²/tháng**. Xem phân tích bên dưới."
                        )

                        # ── Naive load: chuẩn đoán tháng nào bị dồn nhiều ───
                        naive = {ml: {} for ml in month_labels_lp}
                        for p in active:
                            p_months = months_range(p["start"], p["end"])
                            n_pm = len(p_months)
                            if n_pm == 0:
                                continue
                            for pm in p_months:
                                ml = pm.strftime("%m/%Y")
                                if ml in naive:
                                    naive[ml][p["name"]] = p["eff_m2"] / n_pm

                        naive_monthly = {ml: sum(naive[ml].values()) for ml in month_labels_lp}
                        overloaded_naive = [
                            (ml, naive_monthly[ml])
                            for ml in month_labels_lp
                            if naive_monthly[ml] > lp_cap
                        ]

                        # ── LP chẩn đoán: capacity = vô hạn → tìm peak sau khi dàn đều ──
                        with st.spinner("Đang tính năng suất tối thiểu..."):
                            dx = {}
                            for p_idx in range(len(active)):
                                for m_idx in proj_m_idxs[p_idx]:
                                    dx[(p_idx, m_idx)] = pulp.LpVariable(
                                        f"dx_{p_idx}_{m_idx}", lowBound=0
                                    )
                            dz = pulp.LpVariable("dz", lowBound=0)
                            diag_lp = pulp.LpProblem("Diag", pulp.LpMinimize)
                            diag_lp += dz
                            for dm_idx in range(N):
                                diag_lp += (
                                    pulp.lpSum(
                                        dx[(p_idx, dm_idx)]
                                        for p_idx in range(len(active))
                                        if (p_idx, dm_idx) in dx
                                    ) <= dz
                                )
                            for p_idx, p in enumerate(active):
                                diag_lp += (
                                    pulp.lpSum(
                                        dx[(p_idx, m_idx)] for m_idx in proj_m_idxs[p_idx]
                                    ) == p["eff_m2"]
                                )
                            diag_lp.solve(pulp.PULP_CBC_CMD(msg=False))

                        if diag_lp.status == 1:
                            min_cap_need = math.ceil(dz.varValue)
                            # Tải từng tháng theo LP chẩn đoán
                            diag_alloc = {p["name"]: {} for p in active}
                            for p_idx, p in enumerate(active):
                                for m_idx in proj_m_idxs[p_idx]:
                                    val = dx[(p_idx, m_idx)].varValue or 0
                                    if val > 0.05:
                                        diag_alloc[p["name"]][month_labels_lp[m_idx]] = val
                            diag_monthly = {
                                ml: sum(diag_alloc[p["name"]].get(ml, 0) for p in active)
                                for ml in month_labels_lp
                            }
                            peak_ml   = max(diag_monthly, key=diag_monthly.get)
                            peak_load = diag_monthly[peak_ml]
                        else:
                            # Fallback: dự án deadline quá hẹp
                            min_cap_need = math.ceil(
                                max(
                                    p["eff_m2"] / max(1, len(months_range(p["start"], p["end"])))
                                    for p in active
                                )
                            )
                            peak_ml      = max(naive_monthly, key=naive_monthly.get)
                            peak_load    = float(min_cap_need)
                            diag_alloc   = {p["name"]: {} for p in active}
                            diag_monthly = naive_monthly

                        # ── 1. Nguyên nhân ───────────────────────────────────
                        st.markdown("#### 1. Nguyên nhân — Các tháng bị dồn quá nhiều")
                        if overloaded_naive:
                            diag_rows = []
                            for ml, total in sorted(overloaded_naive, key=lambda kv: -kv[1]):
                                top3 = sorted(naive[ml].items(), key=lambda kv: -kv[1])[:3]
                                top3_str = " | ".join(
                                    f"{nm} ({v:,.0f} m²)" for nm, v in top3
                                )
                                diag_rows.append({
                                    "Tháng":                  ml,
                                    "Tải thực cần (m²)":      round(total),
                                    "Năng suất hiện tại (m²)": lp_cap,
                                    "Vượt (m²)":              round(total - lp_cap),
                                    "Dự án chiếm nhiều nhất": top3_str,
                                })
                            st.dataframe(
                                pd.DataFrame(diag_rows),
                                use_container_width=True, hide_index=True,
                            )
                        else:
                            st.write(
                                "Phân bổ đều không vượt năng suất — nguyên nhân do "
                                "deadline quá ngắn so với khối lượng:"
                            )
                            for p in active:
                                n_m = len(months_range(p["start"], p["end"]))
                                if p["eff_m2"] / lp_cap > n_m:
                                    st.write(
                                        f"  • **{p['name']}**: cần tối thiểu "
                                        f"{p['eff_m2']/lp_cap:.1f} tháng, chỉ có {n_m} tháng"
                                    )

                        # ── 2. Gợi ý giải pháp ───────────────────────────────
                        st.markdown("#### 2. Gợi ý giải pháp")
                        shortage = min_cap_need - lp_cap
                        st.info(
                            f"**Tăng năng suất:** Cần tối thiểu **{min_cap_need:,} m²/tháng** "
                            f"để LP có lời giải — tăng thêm **{shortage:,} m²/tháng** "
                            f"so với hiện tại ({lp_cap:,} m²/tháng).\n\n"
                            f"*(Tháng căng nhất sau khi LP dàn đều tối ưu: "
                            f"**{peak_ml}** cần **{round(peak_load):,} m²**)*"
                        )
                        peak_proj = {
                            p["name"]: diag_alloc[p["name"]].get(peak_ml, 0) for p in active
                        }
                        top5 = [(nm, v) for nm, v in
                                sorted(peak_proj.items(), key=lambda kv: -kv[1]) if v > 0.05][:5]
                        if top5:
                            proj_list = "  \n".join(
                                f"  • **{nm}**: {v:,.0f} m²" for nm, v in top5
                            )
                            st.warning(
                                f"**Dời hoặc rút ngắn dự án:** Tháng **{peak_ml}** "
                                f"bị ảnh hưởng nhiều nhất bởi:\n{proj_list}\n\n"
                                f"Hãy cân nhắc dời deadline hoặc giảm XS% "
                                f"các dự án trên trong bảng chỉnh sửa ở trên."
                            )
                    else:
                        # ── Trích kết quả ────────────────────────────────────
                        result = {}
                        for p_idx, p in enumerate(active):
                            result[p["name"]] = {}
                            for m_idx in proj_m_idxs[p_idx]:
                                val = x[(p_idx, m_idx)].varValue or 0
                                if val > 0.05:
                                    result[p["name"]][month_labels_lp[m_idx]] = round(val, 1)

                        monthly_totals = {
                            ml: sum(result[p["name"]].get(ml, 0) for p in active)
                            for ml in month_labels_lp
                        }
                        peak = round(z.varValue or 0, 1)

                        # ── KPI ──────────────────────────────────────────────
                        mode_label = "Kỳ vọng (có xác suất)" if lp_weighted else "Worst case (100% m²)"
                        total_eff  = sum(p["eff_m2"] for p in active)
                        total_nom  = sum(p["m2"]     for p in active)

                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric("Tải trọng đỉnh (LP tối ưu)", f"{peak:,.0f} m²",
                                  f"{peak/lp_cap*100:.1f}% năng suất")
                        n_months_active = sum(1 for t in monthly_totals.values() if t > 0)
                        k2.metric("Tháng có sản xuất", f"{n_months_active}/{N}")
                        k3.metric("Tổng m² lên lịch", f"{total_eff:,.0f}",
                                  f"Danh nghĩa: {total_nom:,.0f}" if lp_weighted else None)
                        k4.metric("Chế độ LP", mode_label)

                        st.divider()

                        # ── Biểu đồ stacked bar ───────────────────────────────
                        st.subheader("📊 Lịch sản xuất theo dự án × tháng")
                        COLORS_LP = (
                            px.colors.qualitative.Plotly
                            + px.colors.qualitative.Set2
                            + px.colors.qualitative.Set3
                        )
                        y_max_lp = max(max(monthly_totals.values(), default=0) * 1.15, lp_cap * 1.15)

                        fig_lp = go.Figure()
                        fig_lp.add_hrect(y0=0, y1=lp_cap * 0.7,
                                         fillcolor="rgba(0,180,0,0.06)", line_width=0)
                        fig_lp.add_hrect(y0=lp_cap * 0.7, y1=lp_cap * 0.9,
                                         fillcolor="rgba(255,200,0,0.12)", line_width=0)
                        fig_lp.add_hrect(y0=lp_cap * 0.9, y1=y_max_lp,
                                         fillcolor="rgba(255,60,0,0.09)", line_width=0)

                        for i, p in enumerate(active):
                            y_vals = [result[p["name"]].get(ml, 0) for ml in month_labels_lp]
                            if any(v > 0 for v in y_vals):
                                seq = i + 1
                                text_vals = [
                                    str(seq) if v >= lp_cap * 0.04 else ""
                                    for v in y_vals
                                ]
                                fig_lp.add_trace(go.Bar(
                                    name=f"{seq}. {p['name']}",
                                    x=month_labels_lp, y=y_vals,
                                    marker_color=COLORS_LP[i % len(COLORS_LP)],
                                    text=text_vals,
                                    textposition="inside",
                                    textfont=dict(size=10, color="white"),
                                    hovertemplate=(
                                        "<b>%{x}</b><br>"
                                        f"#{seq} {p['name']}: " + "%{y:,.0f} m²<extra></extra>"
                                    ),
                                ))

                        fig_lp.add_hline(
                            y=lp_cap, line_dash="dash", line_color="red", line_width=2,
                            annotation_text=f"  Max: {lp_cap:,} m²",
                            annotation_font_color="red", annotation_position="right",
                        )
                        fig_lp.update_layout(
                            barmode="stack", height=540,
                            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
                            xaxis_title="Tháng", yaxis_title="m²",
                            yaxis=dict(range=[0, y_max_lp]),
                            margin=dict(t=20, r=160, b=60, l=70),
                        )
                        st.plotly_chart(fig_lp, use_container_width=True)

                        # ── Gantt Chart ───────────────────────────────────────
                        st.subheader("📅 Gantt — Lịch sản xuất tối ưu")

                        # Sắp xếp dự án theo ngày bắt đầu (sớm nhất ở trên)
                        sorted_gantt = sorted(active, key=lambda p: p["start"])

                        max_prod_val = max(
                            (result.get(p["name"], {}).get(ml, 0)
                             for p in active for ml in month_labels_lp),
                            default=1,
                        ) or 1

                        # Build heatmap matrix: rows=projects, cols=months
                        z_matrix     = []
                        cdata_matrix = []
                        for p in sorted_gantt:
                            cumulative = 0.0
                            row_z, row_c = [], []
                            for ml in month_labels_lp:
                                val = result.get(p["name"], {}).get(ml, 0)
                                if val > 0.05:
                                    cumulative += val
                                cum_pct = min(100.0, cumulative / max(p["eff_m2"], 0.01) * 100)
                                row_z.append(round(val, 1) if val > 0.05 else None)
                                row_c.append([
                                    round(val, 1),
                                    round(cumulative, 1),
                                    round(p["eff_m2"], 1),
                                    round(cum_pct, 1),
                                ])
                            z_matrix.append(row_z)
                            cdata_matrix.append(row_c)

                        # Subplot: heatmap Gantt (trên) + stacked bar capacity (dưới)
                        gantt_h = max(520, len(sorted_gantt) * 28 + 320)
                        fig_gantt = make_subplots(
                            rows=2, cols=1,
                            shared_xaxes=True,
                            row_heights=[0.65, 0.35],
                            vertical_spacing=0.05,
                            subplot_titles=[
                                "Lịch sản xuất từng dự án (màu đậm = SX nhiều hơn)",
                                "Tổng tải theo tháng (m²)",
                            ],
                        )

                        # ── Panel trên: heatmap Gantt ─────────────────────────
                        fig_gantt.add_trace(
                            go.Heatmap(
                                z=z_matrix,
                                x=month_labels_lp,
                                y=[p["name"] for p in sorted_gantt],
                                customdata=cdata_matrix,
                                hovertemplate=(
                                    "<b>%{y}</b><br>"
                                    "Tháng: %{x}<br>"
                                    "SX tháng này: %{customdata[0]:,.0f} m²<br>"
                                    "Lũy kế: %{customdata[1]:,.0f} / "
                                    "%{customdata[2]:,.0f} m² (%{customdata[3]:.1f}%)"
                                    "<extra></extra>"
                                ),
                                colorscale=[
                                    [0.00, "rgb(247,252,253)"],
                                    [0.25, "rgb(204,236,230)"],
                                    [0.50, "rgb(102,194,164)"],
                                    [0.75, "rgb(35,139,69)"],
                                    [1.00, "rgb(0,68,27)"],
                                ],
                                showscale=True,
                                colorbar=dict(
                                    title=dict(text="m²/tháng", side="right"),
                                    thickness=12,
                                    lenmode="fraction", len=0.60,
                                    yanchor="top", y=1.0,
                                    x=1.02,
                                ),
                                zmin=0,
                                zmax=max_prod_val,
                                xgap=1,
                                ygap=1,
                            ),
                            row=1, col=1,
                        )

                        # ── Panel dưới: stacked bar capacity ─────────────────
                        for i, p in enumerate(active):
                            y_vals_g = [result[p["name"]].get(ml, 0) for ml in month_labels_lp]
                            if any(v > 0 for v in y_vals_g):
                                fig_gantt.add_trace(
                                    go.Bar(
                                        x=month_labels_lp,
                                        y=y_vals_g,
                                        name=p["name"],
                                        marker_color=COLORS_LP[i % len(COLORS_LP)],
                                        hovertemplate=(
                                            f"{p['name']}: %{{y:,.0f}} m²"
                                            "<extra></extra>"
                                        ),
                                        showlegend=False,
                                    ),
                                    row=2, col=1,
                                )

                        # Đường năng suất tối đa
                        fig_gantt.add_hline(
                            y=lp_cap, line_dash="dash", line_color="red", line_width=2,
                            annotation_text=f"Max: {lp_cap:,} m²",
                            annotation_font_color="red",
                            row=2, col=1,
                        )
                        # Ngưỡng 80%
                        fig_gantt.add_hline(
                            y=lp_cap * 0.8, line_dash="dot", line_color="darkorange",
                            line_width=1,
                            annotation_text="80%",
                            annotation_font_color="darkorange",
                            row=2, col=1,
                        )
                        # Vùng highlight > 80%
                        cap_upper = max(
                            max(monthly_totals.values(), default=0) * 1.15,
                            lp_cap * 1.15
                        )
                        fig_gantt.add_hrect(
                            y0=lp_cap * 0.8, y1=cap_upper,
                            fillcolor="rgba(255,60,0,0.07)",
                            line_width=0,
                            row=2, col=1,
                        )

                        fig_gantt.update_layout(
                            height=gantt_h,
                            barmode="stack",
                            showlegend=False,
                            margin=dict(l=200, r=80, t=60, b=60),
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                        )
                        # Y trục trên: đảo ngược (sớm nhất ở trên cùng)
                        fig_gantt.update_yaxes(
                            autorange="reversed",
                            tickfont=dict(size=10),
                            row=1, col=1,
                        )
                        # X axis formatting (chỉ hiển thị ở panel dưới vì shared)
                        fig_gantt.update_xaxes(
                            tickangle=45,
                            tickfont=dict(size=9),
                            row=2, col=1,
                        )

                        st.plotly_chart(fig_gantt, use_container_width=True)
                        st.caption(
                            "💡 Dùng toolbar (góc trên phải) để zoom/pan — hai panel đồng bộ trục X. "
                            "Nút 📷 để tải PNG."
                        )

                        # ── Bảng dự án × tháng ───────────────────────────────
                        st.subheader("📋 Bảng phân bổ sản xuất (m²/tháng)")
                        detail_rows = []
                        for i, p in enumerate(active):
                            row_d = {
                                "STT":           i + 1,
                                "Dự án":         p["name"],
                                "XS%":           f"{p['prob']}%",
                                "m² danh nghĩa": f"{p['m2']:,.0f}",
                                "m² lên lịch":   f"{p['eff_m2']:,.0f}",
                                "Bắt đầu":       p["start"].strftime("%m/%Y"),
                                "Kết thúc":      p["end"].strftime("%m/%Y"),
                            }
                            for ml in month_labels_lp:
                                v = result[p["name"]].get(ml, 0)
                                row_d[ml] = f"{v:,.0f}" if v > 0 else ""
                            detail_rows.append(row_d)

                        total_row = {
                            "STT": "", "Dự án": "TONG / % NANG SUAT",
                            "XS%": "",
                            "m² danh nghĩa": f"{total_nom:,.0f}",
                            "m² lên lịch":   f"{total_eff:,.0f}",
                            "Bắt đầu": "", "Kết thúc": "",
                        }
                        for ml in month_labels_lp:
                            t = monthly_totals[ml]
                            pct = t / lp_cap * 100
                            total_row[ml] = f"{t:,.0f} ({pct:.0f}%)" if t > 0 else ""
                        detail_rows.append(total_row)

                        st.dataframe(
                            pd.DataFrame(detail_rows),
                            use_container_width=True, hide_index=True,
                        )

                        # ── Download ─────────────────────────────────────────
                        csv_lp = pd.DataFrame(detail_rows).to_csv(index=False).encode("utf-8-sig")
                        st.download_button(
                            "📥 Tải lịch sản xuất (CSV)",
                            data=csv_lp, file_name="lich_sx_toi_uu.csv", mime="text/csv",
                        )
