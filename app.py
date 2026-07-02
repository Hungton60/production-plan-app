import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, date, timedelta
import re
import io

# Google Sheets integration
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except ImportError:
    GSHEETS_AVAILABLE = False

# ══════════════════════════════════════════════════════════════════════════════
# GOOGLE SHEETS FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_resource(ttl=300)  # Cache 5 minutes
def get_gsheet_client():
    """Initialize Google Sheets client from Streamlit secrets."""
    if not GSHEETS_AVAILABLE:
        return None
    try:
        credentials = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets.readonly",
                "https://www.googleapis.com/auth/drive.readonly",
            ],
        )
        return gspread.authorize(credentials)
    except Exception as e:
        st.error(f"❌ Lỗi kết nối Google Sheets: {e}")
        return None

def read_gsheet_as_excel(sheet_id):
    """Read Google Sheet and return as pandas ExcelFile-like object."""
    try:
        client = get_gsheet_client()
        if client is None:
            return None
        
        spreadsheet = client.open_by_key(sheet_id)
        
        # Create a mock ExcelFile object
        class MockExcelFile:
            def __init__(self, sheets_dict):
                self.sheet_names = list(sheets_dict.keys())
                self._sheets = sheets_dict
            
            def parse(self, sheet_name, header=None):
                if sheet_name in self._sheets:
                    return self._sheets[sheet_name]
                raise ValueError(f"Sheet '{sheet_name}' not found")
        
        # Read all worksheets
        sheets_dict = {}
        for worksheet in spreadsheet.worksheets():
            df = pd.DataFrame(worksheet.get_all_values())
            sheets_dict[worksheet.title] = df
        
        return MockExcelFile(sheets_dict)
        
    except Exception as e:
        st.error(f"❌ Lỗi đọc Google Sheet: {e}")
        return None

# ─── SX Input helpers (PHẢI định nghĩa trước st.set_page_config vì sidebar gọi sớm) ───
def _safe_formula(val):
    if val is None: return None
    if isinstance(val, float) and pd.isna(val): return None
    try: return float(val)
    except (ValueError, TypeError):
        s = str(val).strip()
        if s.startswith("="):
            try: return float(eval(s[1:]))
            except Exception: pass
    return None

def parse_sx_input(xl_sx):
    """Parse SX input format with 2 sheets:
    - máy móc TB: direct m²/hour/machine productivity
    - nhân lực: m²/hour/person (direct) or "tính theo năng suất máy"
    """
    out = {"may_moc": [], "nhan_luc": [], "stage_caps": [],
           "cap_monthly": None, "cap_weekly": None,
           "bottleneck": None, "bottlenecks": []}

    # ── Sheet 1: Máy móc TB ───────────────────────────────────────────────
    sn_mm = next((s for s in xl_sx.sheet_names if "máy" in s.lower() or "may" in s.lower()), None)
    machines_by_code = {}

    if sn_mm:
        df = xl_sx.parse(sn_mm, header=None)
        for _, row in df.iloc[2:].iterrows():
            if len(row) < 8: continue
            ma = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if not ma or ma == "nan": continue
            try:
                ten          = str(row.iloc[1]).strip()
                sl_may       = int(float(str(row.iloc[2])))
                sl_nguoi     = float(str(row.iloc[3]))   # ← FLOAT (e.g. 5.5, 1.5)
                ns_m2_gio    = float(str(row.iloc[4]))
                ca           = int(float(str(row.iloc[5])))
                gio_ca       = int(float(str(row.iloc[6])))
                hs           = float(str(row.iloc[7])) / 100
                cap_thang    = sl_may * ns_m2_gio * ca * gio_ca * 26 * hs
                nv_total     = sl_may * sl_nguoi  # ← có thể là số thập phân

                md = {
                    "ma": ma, "ten": ten, "sl": sl_may, "sl_nguoi": sl_nguoi,
                    "nang_suat_m2_gio": ns_m2_gio, "ca": ca, "gio_ca": gio_ca,
                    "hs_pct": round(hs * 100), "cap_m2_thang": round(cap_thang),
                    "nv_total": nv_total,
                }
                out["may_moc"].append(md)
                machines_by_code[ma] = md
            except Exception:
                continue

    # ── Sheet 2: Nhân lực ─────────────────────────────────────────────────
    sn_nl = next((s for s in xl_sx.sheet_names if "nhân" in s.lower() or "nhan" in s.lower()), None)
    if sn_nl:
        df = xl_sx.parse(sn_nl, header=None)
        for _, row in df.iloc[2:].iterrows():
            if len(row) < 7: continue
            to = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if not to or to == "nan": continue
            try:
                cong_doan    = str(row.iloc[1]).strip()
                nv_input     = float(str(row.iloc[2]))
                ns_val       = str(row.iloc[3]).strip().lower()
                ca           = int(float(str(row.iloc[4])))
                gio_ca       = int(float(str(row.iloc[5])))
                hs           = float(str(row.iloc[6])) / 100

                # Match với máy: dùng từ khóa ưu tiên "dập" > "ghép" > "cnc" > "phay" > "cắt"
                matched_code = None
                if "tính theo" in ns_val or ns_val == "nan" or ns_val == "":
                    for kw in ["dập", "ghép", "cnc", "phay", "cắt"]:
                        if kw in to.lower():
                            for ma, mm in machines_by_code.items():
                                if kw in mm["ten"].lower():
                                    if kw == "cắt" and "dập" in mm["ten"].lower():
                                        continue
                                    matched_code = ma
                                    break
                        if matched_code:
                            break

                if matched_code:
                    machine      = machines_by_code[matched_code]
                    cap_thang    = machine["cap_m2_thang"]
                    nv           = machine["nv_total"]
                    ns_m2_gio    = None
                else:
                    try:
                        ns_m2_gio = float(ns_val)
                    except ValueError:
                        continue   # skip row nếu không parse được
                    cap_thang = nv_input * ns_m2_gio * ca * gio_ca * 26 * hs
                    nv        = nv_input

                out["nhan_luc"].append({
                    "to": to, "cong_doan": cong_doan, "nv": nv,
                    "nang_suat_m2_gio": ns_m2_gio, "ca": ca, "gio_ca": gio_ca,
                    "hs_pct": round(hs * 100), "cap_m2_thang": round(cap_thang),
                    "machine_code": matched_code,
                })
            except Exception:
                continue

    # ── Stage caps & bottleneck ───────────────────────────────────────────
    stage_caps = []
    for nl in out["nhan_luc"]:
        stage_caps.append({
            "cong_doan":    nl["cong_doan"],
            "to":           nl["to"],
            "nv":           nl["nv"],
            "cap_m2_thang": nl["cap_m2_thang"],
            "cap_m2_tuan":  round(nl["cap_m2_thang"] / 4.33),
            "machine_code": nl.get("machine_code"),
        })

    if stage_caps:
        _min = min(s["cap_m2_thang"] for s in stage_caps)
        _bns = [s["cong_doan"] for s in stage_caps if s["cap_m2_thang"] == _min]
        out["cap_monthly"]  = _min
        out["cap_weekly"]   = round(_min / 4.33)
        out["bottleneck"]   = _bns[0]
        out["bottlenecks"]  = _bns
        out["stage_caps"]   = stage_caps

    return out

st.set_page_config(page_title="Kế Hoạch SX – QDP v5.6.26", layout="wide", page_icon="🏭")

# Đảm bảo mobile không thu nhỏ font / scale trang
st.markdown(
    '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">',
    unsafe_allow_html=True,
)

# ─── Custom CSS (chỉ style các class tự định nghĩa, không ghi đè Streamlit) ───
st.markdown("""
<style>
.dash-title {
    background: linear-gradient(135deg, #6c5ce7, #a29bfe);
    color: white; padding: 18px 28px; border-radius: 16px;
    margin-bottom: 12px; display: flex; align-items: center; gap: 14px;
    box-shadow: 0 4px 20px rgba(108,92,231,0.25);
}
.dash-title h1 { margin: 0; font-size: 1.5rem; font-weight: 700; color: white; }
.dash-title p  { margin: 0; font-size: 0.88rem; opacity: 0.88; color: white; }
.kpi-row { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; }
.kpi-card {
    flex: 1; min-width: 155px;
    background: #ffffff;
    border-radius: 14px;
    padding: 18px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.10);
    border-top: 4px solid #6c5ce7;
}
.kpi-card.red    { border-top-color: #e17055; }
.kpi-card.orange { border-top-color: #fd9644; }
.kpi-card.yellow { border-top-color: #e1b000; }
.kpi-card.green  { border-top-color: #00b894; }
.kpi-label {
    font-size: 0.72rem; color: #555; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 6px;
}
.kpi-value { font-size: 1.9rem; font-weight: 800; line-height: 1.15; color: #222; }
.kpi-value.red    { color: #e17055; }
.kpi-value.orange { color: #fd9644; }
.kpi-value.yellow { color: #c49000; }
.kpi-value.green  { color: #00b894; }
.kpi-sub { font-size: 0.71rem; color: #777; margin-top: 3px; }
@media print {
  /* Ẩn chrome UI */
  [data-testid="stSidebar"]         { display: none !important; }
  [data-testid="stHeader"]          { display: none !important; }
  [data-testid="stToolbar"]         { display: none !important; }
  [data-testid="stDecoration"]      { display: none !important; }
  [data-testid="stStatusWidget"]    { display: none !important; }
  .stTabs [data-baseweb="tab-list"] { display: none !important; }
  /* Thu nhỏ toàn bộ để vừa 1 trang A3 ngang */
  [data-testid="stMainBlockContainer"],
  [data-testid="stMain"] .block-container { zoom: 0.62 !important; }
  /* KPI cards compact hơn */
  .kpi-card  { padding: 8px 12px !important; }
  .kpi-value { font-size: 1.3rem !important; }
  .kpi-label { font-size: 0.60rem !important; }
  .kpi-sub   { font-size: 0.58rem !important; }
  .kpi-row   { margin-bottom: 8px !important; gap: 8px !important; }
  /* Banner compact */
  .dash-title    { padding: 8px 16px !important; margin-bottom: 6px !important; }
  .dash-title h1 { font-size: 1.1rem !important; }
  .dash-title p  { font-size: 0.75rem !important; }
  /* Tránh cắt biểu đồ */
  .element-container  { page-break-inside: avoid !important; }
  .stPlotlyChart      { page-break-inside: avoid !important; }
  /* Đặt trang A3 ngang */
  @page { size: A3 landscape; margin: 8mm; }
}

/* ── Mobile responsive ────────────────────────────────────────────────── */
/* Cảnh báo xoay màn hình – chỉ hiện ở điện thoại đứng */
.rotate-hint {
    display: none;
    background: #6c5ce7;
    color: white;
    text-align: center;
    padding: 10px 16px;
    border-radius: 10px;
    font-size: 0.9rem;
    margin-bottom: 12px;
    animation: pulse-hint 2s infinite;
}
@keyframes pulse-hint {
    0%, 100% { opacity: 1; }
    50%       { opacity: 0.75; }
}
@media screen and (max-width: 768px) and (orientation: portrait) {
    .rotate-hint { display: block !important; }
    .kpi-row { flex-direction: column; gap: 8px; }
    .kpi-card { min-width: 100% !important; }
    .dash-title h1 { font-size: 1.1rem !important; }
    .dash-title p  { font-size: 0.78rem !important; }
}
@media screen and (max-width: 1024px) and (orientation: landscape) {
    .kpi-card { min-width: 120px; padding: 12px 14px; }
    .kpi-value { font-size: 1.5rem; }
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="rotate-hint">
  📱↩️ Xoay điện thoại nằm ngang để xem đầy đủ dashboard!
</div>
<div class="dash-title">
  <div style="font-size:2.2rem">🏭</div>
  <div>
    <h1>Kế Hoạch Sản Xuất – Nhà Máy QDP <span style="font-size:0.5em;color:#aaaaaa;font-weight:400">v5.6.26</span></h1>
    <p>Dashboard quản lý tải trọng &amp; tiến độ sản xuất nhôm kính</p>
  </div>
</div>
""", unsafe_allow_html=True)


# ─── Sidebar ──────────────────────────────────────────────────────────────────

st.sidebar.markdown(
    "<div style='text-align:right;color:#aaa;font-size:0.75rem;margin-bottom:8px'>📦 version 5.6.26</div>",
    unsafe_allow_html=True
)

# ── Tải file Năng Lực SX (đặt TRƯỚC các input để lấy giá trị default) ─────────
st.sidebar.header("🔧 Năng Lực Sản Xuất")
uploaded_sx = st.sidebar.file_uploader(
    "📂 Tải file SX input.xlsx (máy móc · nhân lực)",
    type=["xlsx", "xls"], key="sx_file",
    help="File gồm 2 sheet: 'máy móc TB', 'nhân lực'"
)

sx_data = None
_sx_cap_monthly = 10000   # default
_sx_cap_weekly  = 2500    # default

if uploaded_sx is not None:
    _sx_fid = f"{uploaded_sx.name}_{uploaded_sx.size}"
    # Cache kết quả parse vào session state để tránh parse lại mỗi lần rerun
    if st.session_state.get("sx_file_id") != _sx_fid:
        try:
            _xl_sx = pd.ExcelFile(uploaded_sx)
            _parsed = parse_sx_input(_xl_sx)
            st.session_state["sx_file_id"]   = _sx_fid
            st.session_state["sx_data"]      = _parsed
        except Exception as _e:
            st.sidebar.error(f"Lỗi đọc file SX: {_e}")
            st.session_state["sx_data"] = None

    sx_data = st.session_state.get("sx_data")
    if sx_data and sx_data["cap_monthly"]:
        _sx_cap_monthly = sx_data["cap_monthly"]
        _sx_cap_weekly  = sx_data["cap_weekly"] or round(_sx_cap_monthly / 4.33)
        _bns_sidebar = sx_data.get("bottlenecks") or [sx_data["bottleneck"]]
        st.sidebar.success(
            f"✅ **{_sx_cap_monthly:,} m²/tháng**  \n"
            f"Nút thắt: *{' · '.join(_bns_sidebar)}*"
        )
    elif sx_data is not None:
        st.sidebar.warning("⚠️ Không tính được năng suất từ file SX.")

st.sidebar.divider()
st.sidebar.header("⚙️ Tham số nhà máy")

# Nếu đã có file SX → dùng giá trị tính được làm default; vẫn cho chỉnh tay
_use_sx_cap = sx_data is not None and sx_data.get("cap_monthly") is not None
cap_monthly = st.sidebar.number_input(
    "Năng suất tối đa (m²/tháng)",
    value=_sx_cap_monthly, step=500, min_value=100,
    help="Tự động điền từ file SX input nếu đã tải." if _use_sx_cap else "",
)
if _use_sx_cap:
    st.sidebar.caption("↑ Tự tính từ file SX input (có thể điều chỉnh thủ công)")

horizon = st.sidebar.selectbox("Tầm nhìn kế hoạch (tháng)", [12, 18, 24, 36, 48, 60], index=1)
show_weighted = st.sidebar.checkbox(
    "Tải trọng có xác suất",
    value=True,
    help="Tích: tính m² × xác suất (kỳ vọng). Bỏ tích: 100% khối lượng (xấu nhất).",
)

st.sidebar.divider()
st.sidebar.header("📅 Thông số kế hoạch tuần")
cap_weekly = st.sidebar.number_input(
    "Năng suất tối đa (m²/tuần)",
    value=_sx_cap_weekly, step=100, min_value=100,
    help="Tự động điền từ file SX input nếu đã tải." if _use_sx_cap else "",
)
if _use_sx_cap:
    st.sidebar.caption("↑ Tự tính từ file SX input (có thể điều chỉnh thủ công)")

max_indirect = st.sidebar.number_input(
    "KS gián tiếp tối đa (người)", value=40, step=1, min_value=1
)
avail_xuong = st.sidebar.number_input(
    "Diện tích xưởng thực tế (m²)", value=20000, step=500, min_value=100
)
avail_kho_tp = st.sidebar.number_input(
    "Diện tích kho TP thực tế (m²)", value=2000, step=200, min_value=100
)
st.sidebar.divider()
st.sidebar.subheader("📐 Hệ số mặt bằng")
coeff_xuong = st.sidebar.number_input(
    "Hệ số nhà xưởng (m²/m² SP)", value=8.0, step=0.1, min_value=0.1, format="%.2f",
    help="Mặt bằng xưởng yêu cầu = m² SX/tuần × hệ số này"
)
coeff_kho_tp = st.sidebar.number_input(
    "Hệ số kho thành phẩm (m²/m² SP)", value=0.5, step=0.1, min_value=0.1, format="%.2f",
    help="Diện tích kho TP yêu cầu = m² SX/tuần × hệ số này"
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
    """Chuyển Excel serial number hoặc datetime thành datetime, giữ nguyên ngày."""
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        try:
            return datetime(1899, 12, 30) + timedelta(days=int(val))
        except Exception:
            return None
    try:
        ts = pd.to_datetime(val)
        return ts.to_pydatetime().replace(tzinfo=None)
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
    #   col END+1 = XS hệ số (số 0–1) HOẶC Ghi chú (text có %)
    end_col  = detect_end_col(df)
    xs_col   = end_col + 1   # cột XS hoặc ghi chú

    projects = []
    skipped  = []   # ← lưu lại các dòng có tên nhưng thiếu dữ liệu + lý do
    skip = {"nan", "", "TÊN DỰ ÁN", "GHI CHÚ:", "XS (hệ số)"}

    # ── Tìm dòng tiêu đề cột (ô "TÊN DỰ ÁN") làm MỐC ─────────────────────────
    # Mọi dòng NẰM TRÊN mốc này (ví dụ dòng tiêu đề lớn "KẾ HOẠCH SẢN XUẤT...")
    # luôn được bỏ qua âm thầm, không cảnh báo — vì đó chắc chắn không phải
    # dữ liệu dự án. Mọi dòng NẰM DƯỚI mốc này, hễ có tên dự án mà thiếu bất
    # kỳ cột dữ liệu nào (Khối lượng / Ngày bắt đầu / Ngày hoàn thành) đều bị
    # đưa vào danh sách cảnh báo.
    header_idx = None
    for row_idx, row in df.iterrows():
        _cell = str(row.iloc[1]).strip().upper() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        if _cell == "TÊN DỰ ÁN":
            header_idx = row_idx
            break

    for row_idx, row in df.iterrows():
        if header_idx is not None and row_idx <= header_idx:
            continue   # dòng tiêu đề lớn hoặc chính dòng tên cột → bỏ qua âm thầm
        if len(row) <= end_col:
            continue

        name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if name in skip or not name:
            continue   # dòng hoàn toàn không có tên dự án → bỏ qua âm thầm

        _m2_raw    = row.iloc[2] if len(row) > 2 else None
        _start_raw = row.iloc[3] if len(row) > 3 else None
        _end_raw   = row.iloc[end_col] if len(row) > end_col else None

        # ── Kiểm tra TỪNG cột, gom hết các cột bị thiếu vào 1 cảnh báo ───────
        missing = []

        m2 = None
        try:
            m2 = float(_m2_raw)
            if pd.isna(m2):
                raise ValueError
        except (ValueError, TypeError):
            missing.append("Khối lượng (m²)")

        start = parse_date(_start_raw)
        if start is None:
            missing.append("Ngày bắt đầu thực hiện")

        end = parse_date(_end_raw)
        if end is None:
            missing.append("Ngày hoàn thành")

        if missing:
            skipped.append({"name": name, "reason": "Thiếu " + ", ".join(missing)})
            continue

        # Đọc xác suất: ưu tiên cột XS số (0–1 hoặc 1–100), fallback text
        prob = 100
        if len(row) > xs_col and pd.notna(row.iloc[xs_col]):
            _xs = row.iloc[xs_col]
            try:
                _xsf = float(_xs)
                if 0 < _xsf <= 1:        # hệ số kiểu 0.5 / 0.9 / 1.0
                    prob = round(_xsf * 100)
                elif 1 < _xsf <= 100:    # % kiểu 50 / 90 / 100
                    prob = round(_xsf)
                else:
                    prob = extract_prob(str(_xs))
            except (ValueError, TypeError):
                prob = extract_prob(str(_xs))

        projects.append({
            "name": name, "m2": m2,
            "start": start, "end": end,
            "prob": prob, "status": status_label(prob),
        })
    return projects, skipped

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

# ─── Weekly helpers ───────────────────────────────────────────────────────────
def build_weekly_production(projects, show_weighted, weekly_custom=None):
    """Tính m²/tuần cho từng dự án. Ưu tiên dùng weekly_custom nếu có."""
    if weekly_custom is None:
        weekly_custom = {}
    if not projects:
        return {}, [], []
    min_start = pd.Timestamp(min(p["start"] for p in projects))
    max_end   = pd.Timestamp(max(p["end"]   for p in projects)) + pd.DateOffset(months=1)
    week_dates = pd.date_range(start=min_start, end=max_end, freq="W-MON")
    week_labels = [f"{d.isocalendar()[0]}-W{d.isocalendar()[1]:02d}" for d in week_dates]
    load = {wl: {} for wl in week_labels}
    for p in projects:
        p_weekly = weekly_custom.get(p["name"], {})
        factor = (p["prob"] / 100) if show_weighted else 1.0
        if p_weekly:
            for wl, v in p_weekly.items():
                if wl in load:
                    load[wl][p["name"]] = round(v * factor, 1)
        else:
            p_start = pd.Timestamp(p["start"])
            p_end   = pd.Timestamp(p["end"]) + pd.DateOffset(months=1) - pd.Timedelta(days=1)
            proj_weeks = [
                (d, f"{d.isocalendar()[0]}-W{d.isocalendar()[1]:02d}")
                for d in week_dates
                if p_start <= d + pd.Timedelta(days=6) and d <= p_end
            ]
            n = len(proj_weeks)
            if n == 0:
                continue
            m2w = p["m2"] * factor / n
            for _, wl in proj_weeks:
                if wl in load:
                    load[wl][p["name"]] = round(m2w, 1)
    return load, week_labels, list(week_dates)

def _is_numeric(v):
    try:
        float(v)
        return True
    except (ValueError, TypeError):
        return False

def _safe_float(v):
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0

def parse_tech_input(xl):
    """Đọc sheet 'Tech input' → {iso_week_label: tổng_kỹ_sư}."""
    if "Tech input" not in xl.sheet_names:
        return {}
    df = xl.parse("Tech input", header=None)
    if len(df) < 2:
        return {}
    row0, row1 = df.iloc[0], df.iloc[1]
    col_to_week = {}
    current_month = None
    for ci in range(16, len(df.columns)):
        v0 = row0.iloc[ci] if ci < len(row0) else None
        if pd.notna(v0):
            try:
                current_month = pd.to_datetime(v0)
            except Exception:
                pass
        v1 = row1.iloc[ci] if ci < len(row1) else None
        if current_month is not None and pd.notna(v1) and str(v1).strip().lower().startswith("week"):
            try:
                wn = int(str(v1).strip().split()[-1])
                d  = current_month + pd.Timedelta(weeks=wn - 1)
                iso = d.isocalendar()
                col_to_week[ci] = f"{iso[0]}-W{iso[1]:02d}"
            except Exception:
                pass
    # Tìm hàng tổng (hàng cuối có giá trị số > 0 trong các cột tuần)
    # Tránh cộng dồn các hàng lẻ lẫn hàng tổng (sẽ bị double-count)
    weekly = {}
    last_data_row = None
    for ri in range(2, len(df)):
        row = df.iloc[ri]
        has_val = any(
            True for ci in col_to_week
            if ci < len(row) and pd.notna(row.iloc[ci]) and str(row.iloc[ci]).strip() not in ("", "nan")
            and _is_numeric(row.iloc[ci])
        )
        if has_val:
            last_data_row = ri

    if last_data_row is not None:
        # Kiểm tra xem hàng cuối có phải hàng tổng không
        # (tổng các hàng bên trên ≈ giá trị hàng cuối → đây là hàng SUM)
        row_sum = df.iloc[last_data_row]
        above_sum = {}
        for ri in range(2, last_data_row):
            r = df.iloc[ri]
            for ci in col_to_week:
                try:
                    v = float(r.iloc[ci])
                    if not pd.isna(v) and v > 0:
                        above_sum[ci] = above_sum.get(ci, 0) + v
                except (ValueError, TypeError):
                    pass
        # Nếu hàng cuối ≈ tổng các hàng trên → dùng hàng cuối làm nguồn
        sample_cols = list(col_to_week.keys())[:5]
        is_summary = all(
            abs(_safe_float(row_sum.iloc[ci]) - above_sum.get(ci, 0)) < 1
            for ci in sample_cols if ci < len(row_sum)
        )
        if is_summary:
            for ci, wl in col_to_week.items():
                try:
                    v = float(row_sum.iloc[ci])
                    if not pd.isna(v) and v > 0:
                        weekly[wl] = v
                except (ValueError, TypeError):
                    pass
        else:
            # Không có hàng tổng → cộng tất cả hàng lẻ
            for ri in range(2, len(df)):
                row = df.iloc[ri]
                for ci, wl in col_to_week.items():
                    try:
                        v = float(row.iloc[ci])
                        if not pd.isna(v) and v > 0:
                            weekly[wl] = weekly.get(wl, 0) + v
                    except (ValueError, TypeError):
                        pass
    return weekly

def parse_mat_bang(xl_main):
    """Đọc sheet 'Mặt bằng yêu cầu' từ file chính → DataFrame hoặc None."""
    if xl_main is None or "Mặt bằng yêu cầu" not in xl_main.sheet_names:
        return None
    try:
        return xl_main.parse("Mặt bằng yêu cầu", header=0)
    except Exception as e:
        st.error(f"Lỗi đọc sheet Mặt bằng yêu cầu: {e}")
        return None

# ─── Tính weekly data trước khi render tabs (dashboard cần dùng) ──────────────
if st.session_state.get("t1_rows"):
    _rows_pre = st.session_state["t1_rows"]
    _projs_pre = [{"name": r["name"], "m2": r["m2"], "prob": r["prob"],
                   "start": datetime(r["start"].year, r["start"].month, 1),
                   "end":   datetime(r["end"].year,   r["end"].month,   1)}
                  for r in _rows_pre]
    _wload_pre, _wla_pre, _wda_pre = build_weekly_production(
        _projs_pre, show_weighted,
        weekly_custom=st.session_state.get("t1_weekly_custom", {})
    )
    _today_pre  = datetime(datetime.today().year, datetime.today().month, 1)
    _cutoff_pre = pd.Timestamp(_today_pre) + pd.DateOffset(months=horizon)
    _pairs_pre  = [(wl, wd) for wl, wd in zip(_wla_pre, _wda_pre) if wd <= _cutoff_pre] \
                  or list(zip(_wla_pre, _wda_pre))
    st.session_state["t1_w_labels"]    = [p[0] for p in _pairs_pre]
    st.session_state["t1_weekly_load"] = _wload_pre

tab_dash, tab_plan, tab_nanluc, tab_nvcnl, tab_thucte = st.tabs([
    "🏠 Dashboard Tổng Quan",
    "📊 Kế Hoạch Sản Xuất",
    "🔧 Năng Lực SX",
    "📈 Nhu Cầu vs Năng Lực",
    "📋 Tiến Độ Thực Tế",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with tab_dash:
    rows_dash = st.session_state.get("t1_rows", [])
    if not rows_dash:
        st.info("⬅️ Vui lòng vào tab **📊 Kế Hoạch Sản Xuất**, tải file Excel để Dashboard tự động cập nhật.")
    else:
        from datetime import datetime as _dt

        def _nx(d):
            return _dt(d.year+1,1,1) if d.month==12 else _dt(d.year, d.month+1, 1)

        def _mrange(s, e):
            ms, cur, em = [], _dt(s.year,s.month,1), _dt(e.year,e.month,1)
            while cur <= em: ms.append(cur); cur = _nx(cur)
            return ms

        td = _dt(_dt.today().year, _dt.today().month, 1)
        projs = [{"name": r["name"], "m2": r["m2"], "prob": r["prob"],
                  "start": _dt(r["start"].year, r["start"].month, 1),
                  "end":   _dt(r["end"].year,   r["end"].month,   1)}
                 for r in rows_dash]

        total_m2    = sum(p["m2"] for p in projs)
        n_proj      = len(projs)
        signed      = sum(1 for p in projs if p["prob"] == 100)
        in_progress = sum(1 for p in projs if p["start"] <= td <= p["end"])

        # KPI row 1
        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi-card" style="border-top-color:#6c5ce7">
            <div class="kpi-label">Tổng dự án</div>
            <div class="kpi-value" style="color:#6c5ce7">{n_proj}</div>
            <div class="kpi-sub">dự án đang theo dõi</div>
          </div>
          <div class="kpi-card green">
            <div class="kpi-label">Đã ký HĐ</div>
            <div class="kpi-value green">{signed}</div>
            <div class="kpi-sub">xác suất 100%</div>
          </div>
          <div class="kpi-card orange">
            <div class="kpi-label">Đang sản xuất</div>
            <div class="kpi-value orange">{in_progress}</div>
            <div class="kpi-sub">tháng {td.strftime("%m/%Y")}</div>
          </div>
          <div class="kpi-card" style="border-top-color:#00cec9">
            <div class="kpi-label">Tổng khối lượng</div>
            <div class="kpi-value" style="color:#00cec9;font-size:1.4rem">{total_m2:,.0f}</div>
            <div class="kpi-sub">m² (kỳ vọng)</div>
          </div>
          <div class="kpi-card" style="border-top-color:#00b894">
            <div class="kpi-label">Năng lực SX</div>
            <div class="kpi-value" style="color:#00b894;font-size:1.4rem">{cap_monthly:,}</div>
            <div class="kpi-sub">m²/tháng {"· 🔴 " + (sx_data["bottleneck"] or "") if sx_data else "· (mặc định)"}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Build monthly load
        mnths = []
        cur = td
        for _ in range(18):
            mnths.append(cur); cur = _nx(cur)
        ml = [m.strftime("%m/%Y") for m in mnths]
        load = {x: 0.0 for x in ml}
        for p in projs:
            pm_list = _mrange(p["start"], p["end"])
            n = len(pm_list)
            if n == 0: continue
            for pm in pm_list:
                k = pm.strftime("%m/%Y")
                if k in load:
                    load[k] += p["m2"] * (p["prob"] / 100) / n

        pcts = [round(load[x] / cap_monthly * 100, 1) for x in ml]
        n_over  = sum(1 for v in pcts if v >= 100)
        n_tight = sum(1 for v in pcts if 90 <= v < 100)
        n_warn  = sum(1 for v in pcts if 70 <= v < 90)
        n_ok    = sum(1 for v in pcts if v < 70)
        avg_u   = round(sum(pcts)/len(pcts), 1) if pcts else 0

        # KPI row 2
        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi-card red">
            <div class="kpi-label">Tháng quá tải</div>
            <div class="kpi-value red">{n_over}</div>
            <div class="kpi-sub">≥ 100% năng suất</div>
          </div>
          <div class="kpi-card orange">
            <div class="kpi-label">Căng tiến độ</div>
            <div class="kpi-value orange">{n_tight}</div>
            <div class="kpi-sub">90–100%</div>
          </div>
          <div class="kpi-card yellow">
            <div class="kpi-label">Cần chú ý</div>
            <div class="kpi-value yellow">{n_warn}</div>
            <div class="kpi-sub">70–90%</div>
          </div>
          <div class="kpi-card green">
            <div class="kpi-label">Bình thường</div>
            <div class="kpi-value green">{n_ok}</div>
            <div class="kpi-sub">&lt; 70%</div>
          </div>
          <div class="kpi-card" style="border-top-color:#a29bfe">
            <div class="kpi-label">Tải TB 18 tháng</div>
            <div class="kpi-value" style="color:#6c5ce7;font-size:1.4rem">{avg_u}%</div>
            <div class="kpi-sub">năng suất sử dụng</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Charts
        col_l, col_r = st.columns([2, 1])
        with col_l:
            def _bar_color(v):
                if v >= 100: return "#e17055"
                if v >= 90:  return "#fd9644"
                if v >= 70:  return "#f9ca24"
                return "#00b894"

            bar_c = [_bar_color(v) for v in pcts]
            fig_bar = go.Figure(go.Bar(
                x=ml, y=[round(load[x]) if load[x] == load[x] else 0 for x in ml],
                marker_color=bar_c,
                hovertemplate="<b>%{x}</b><br>%{y:,.0f} m²  (%{customdata:.1f}%)<extra></extra>",
                customdata=pcts,
            ))
            fig_bar.add_hline(y=cap_monthly, line_dash="dash", line_color="red",
                              annotation_text=f"Max: {cap_monthly:,} m²",
                              annotation_font_color="red", annotation_position="right")
            fig_bar.update_layout(
                height=300, title="Tải trọng nhà máy 18 tháng tới",
                plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
                xaxis=dict(tickangle=45, tickfont=dict(size=9), gridcolor="#e0e0e0",
                           color="#444"),
                yaxis=dict(gridcolor="#e0e0e0", color="#444"),
                margin=dict(t=40, r=120, b=60, l=60),
                font=dict(family="Inter, sans-serif", color="#333"),
            )
            st.plotly_chart(fig_bar, use_container_width=True)
            st.caption("🟢 < 70% bình thường  ·  🟡 70–90% cần chú ý  ·  🟠 90–100% căng tiến độ  ·  🔴 ≥ 100% quá tải")

        with col_r:
            groups = {"Đã ký HĐ": 0, "Khả năng cao": 0, "Đang xét": 0, "Khác": 0}
            for p in projs:
                if p["prob"] == 100:   groups["Đã ký HĐ"]      += p["m2"]
                elif p["prob"] >= 90:  groups["Khả năng cao"]   += p["m2"]
                elif p["prob"] >= 50:  groups["Đang xét"]        += p["m2"]
                else:                  groups["Khác"]             += p["m2"]
            labs = [k for k, v in groups.items() if v > 0]
            vals = [v for v in groups.values()  if v > 0]
            fig_pie = go.Figure(go.Pie(
                labels=labs, values=vals, hole=0.55,
                marker_colors=["#00b894","#6c5ce7","#fd9644","#b2bec3"],
                textinfo="percent",
                hovertemplate="%{label}<br>%{value:,.0f} m²<extra></extra>",
            ))
            fig_pie.update_layout(
                height=300, title="Phân bổ KL theo XS%",
                plot_bgcolor="white", paper_bgcolor="white",
                legend=dict(orientation="v", font=dict(size=10, color="#333")),
                margin=dict(t=40, r=10, b=10, l=10),
                font=dict(family="Inter, sans-serif", color="#333"),
                title_font=dict(color="#333"),
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        # Gantt
        st.subheader("📅 Tiến độ dự án")
        gdata = []
        cmap = {"Đã ký HĐ": "#00b894", "Khả năng cao": "#6c5ce7",
                "Đang xét": "#fd9644", "Khác": "#b2bec3"}
        for p in projs:
            if p["prob"]==100:   g="Đã ký HĐ"
            elif p["prob"]>=90:  g="Khả năng cao"
            elif p["prob"]>=50:  g="Đang xét"
            else:                g="Khác"
            gdata.append({"Dự án": p["name"], "Bắt đầu": p["start"],
                          "Kết thúc": _nx(p["end"]), "Nhóm": g,
                          "m²": p["m2"], "XS%": p["prob"]})
        df_g = pd.DataFrame(gdata).sort_values("Bắt đầu")
        fig_g = px.timeline(df_g, x_start="Bắt đầu", x_end="Kết thúc",
                            y="Dự án", color="Nhóm", color_discrete_map=cmap,
                            custom_data=["m²","XS%"],
                            template="plotly_white")
        fig_g.update_traces(
            hovertemplate="<b>%{y}</b><br>%{customdata[0]:,.0f} m²  |  XS: %{customdata[1]}%<extra></extra>"
        )
        all_starts = [p["start"] for p in projs]
        all_ends   = [_nx(p["end"]) for p in projs]
        x_min = min(all_starts).strftime("%Y-%m-%d")
        x_max = max(all_ends).strftime("%Y-%m-%d")

        fig_g.add_vline(x=td.timestamp()*1000, line_color="red", line_width=2,
                        line_dash="dash", annotation_text="  Hôm nay",
                        annotation_font_color="red", annotation_position="top right")
        _all_proj_names_g = df_g["Dự án"].tolist()
        fig_g.update_yaxes(
            autorange="reversed", automargin=True,
            tickmode="array",
            tickvals=_all_proj_names_g,
            ticktext=_all_proj_names_g,
            tickfont=dict(size=9, color="#333"), color="#333",
        )
        fig_g.update_xaxes(
            range=[x_min, x_max],
            dtick="M1", tickformat="%m/%Y", tickangle=45,
            tickfont=dict(size=8, color="#333"), color="#333",
        )
        fig_g.update_layout(
            height=max(150, len(projs) * 26 + 80),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0,
                        font=dict(size=10, color="#333")),
            margin=dict(l=10, r=20, t=36, b=50),
            font=dict(family="Inter, sans-serif", color="#333"),
        )
        st.plotly_chart(fig_g, use_container_width=True)

        # ── Nhân lực & Mặt bằng (từ dữ liệu tuần đã lưu) ────────────────────
        _wl   = st.session_state.get("t1_w_labels", [])
        _wload = st.session_state.get("t1_weekly_load", {})
        _tw   = st.session_state.get("t1_tech_weekly", {})

        if _wl and _wload:
            col_nl, col_mb = st.columns(2)

            with col_nl:
                st.markdown("**👥 Nhân lực kỹ thuật gián tiếp theo tuần**")
                if _tw:
                    _eng = [_tw.get(wl, 0) for wl in _wl]
                    _eng_c = []
                    for v in _eng:
                        pct = v / max_indirect * 100 if max_indirect > 0 else 0
                        if pct >= 100:  _eng_c.append("#e17055")
                        elif pct >= 90: _eng_c.append("#fd9644")
                        elif pct >= 70: _eng_c.append("#f9ca24")
                        else:           _eng_c.append("#00b894")
                    fig_nl = go.Figure(go.Bar(
                        x=_wl, y=_eng, marker_color=_eng_c,
                        hovertemplate="<b>%{x}</b><br>%{y:.0f} người<extra></extra>",
                    ))
                    fig_nl.add_hline(y=max_indirect, line_dash="dash", line_color="red",
                                     annotation_text=f"  Max: {max_indirect} người",
                                     annotation_font_color="red", annotation_position="right")
                    fig_nl.update_layout(
                        height=260, showlegend=False,
                        plot_bgcolor="white", paper_bgcolor="white",
                        xaxis=dict(tickangle=45, tickfont=dict(size=7), color="#333"),
                        yaxis=dict(title="Người", color="#333"),
                        margin=dict(t=20, r=120, b=70, l=55),
                        font=dict(family="Inter, sans-serif", color="#333"),
                    )
                    st.plotly_chart(fig_nl, use_container_width=True)
                else:
                    st.info("File chưa có sheet **Tech input**.")

            with col_mb:
                st.markdown("**🏗️ Mặt bằng xưởng & kho TP theo tuần**")
                _xu = [sum(_wload.get(wl, {}).values()) * coeff_xuong for wl in _wl]
                _kh = [sum(_wload.get(wl, {}).values()) * coeff_kho_tp for wl in _wl]
                fig_mb2 = go.Figure()
                for vals, avail, color, label in [
                    (_xu, avail_xuong,  "#1f77b4", "Xưởng SX"),
                    (_kh, avail_kho_tp, "#2ca02c", "Kho TP"),
                ]:
                    fig_mb2.add_trace(go.Scatter(
                        x=_wl, y=vals, mode="lines+markers", name=f"{label} (yêu cầu)",
                        line=dict(color=color, width=2), marker=dict(size=4),
                        hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{y:,.0f}} m²<extra></extra>",
                    ))
                    fig_mb2.add_hline(y=avail, line_dash="dash", line_color=color,
                                      annotation_text=f"  {label}: {avail:,} m²",
                                      annotation_font_color=color, annotation_position="right")
                fig_mb2.update_layout(
                    height=260,
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                                font=dict(size=10, color="#333")),
                    xaxis=dict(tickangle=45, tickfont=dict(size=7), color="#333"),
                    yaxis=dict(title="m²", color="#333"),
                    margin=dict(t=20, r=120, b=70, l=60),
                    font=dict(family="Inter, sans-serif", color="#333"),
                )
                st.plotly_chart(fig_mb2, use_container_width=True)

        # ── Nút xuất HTML để in ─────────────────────────────────────────────
        st.divider()
        if st.button("🖨️ Xuất HTML để in (A3 ngang)", key="dash_print"):
            import plotly.io as pio

            def _fig_html(fig, w, h):
                f2 = fig
                f2.update_layout(width=w, height=h,
                                 paper_bgcolor="white", plot_bgcolor="white",
                                 font=dict(color="#333"))
                return pio.to_html(f2, include_plotlyjs=False, full_html=False,
                                   config={"displayModeBar": False})

            # ── Kích thước cột cho bố cục 3 cột A3 landscape ────────────────
            # A3 landscape usable: ~1527px × 1062px (sau margin 8mm)
            # Trừ banner 52px + KPI row 62px + gaps: content cao ~930px
            _SIDE_W   = 420   # px chiều rộng cột trái/phải
            _CENTER_W = 650   # px chiều rộng cột giữa (Gantt)
            _SIDE_CH  = 420   # px chiều cao mỗi biểu đồ nhỏ (2 cái/cột)
            _GANTT_H  = max(500, len(projs) * 28 + 80)   # đủ chỗ cho tất cả dự án

            # Rebuild figures
            _fb_p = go.Figure(fig_bar)
            _fb_p.update_layout(width=_SIDE_W, height=_SIDE_CH,
                                margin=dict(t=30, r=10, b=55, l=55),
                                paper_bgcolor="white", plot_bgcolor="white",
                                font=dict(color="#333"), showlegend=True,
                                legend=dict(orientation="h", y=1.05, x=0, font=dict(size=8)))
            _fb_h = pio.to_html(_fb_p, include_plotlyjs=False, full_html=False,
                                config={"displayModeBar": False})

            _fp_p = go.Figure(fig_pie)
            _fp_p.update_layout(width=_SIDE_W, height=_SIDE_CH,
                                margin=dict(t=30, r=10, b=10, l=10),
                                paper_bgcolor="white", plot_bgcolor="white",
                                font=dict(color="#333"))
            _fp_h = pio.to_html(_fp_p, include_plotlyjs=False, full_html=False,
                                config={"displayModeBar": False})

            _nl_p = go.Figure(fig_nl) if _tw else None
            if _nl_p:
                _nl_p.update_layout(width=_SIDE_W, height=_SIDE_CH,
                                    margin=dict(t=20, r=80, b=60, l=50),
                                    paper_bgcolor="white", plot_bgcolor="white",
                                    font=dict(color="#333"), showlegend=False)
                _nl_h = pio.to_html(_nl_p, include_plotlyjs=False, full_html=False,
                                    config={"displayModeBar": False})
            else:
                _nl_h = "<p style='color:#888;font-size:11px;padding:16px'>Không có dữ liệu nhân lực</p>"

            _mb_p = go.Figure(fig_mb2)
            _mb_p.update_layout(width=_SIDE_W, height=_SIDE_CH,
                                margin=dict(t=20, r=80, b=60, l=50),
                                paper_bgcolor="white", plot_bgcolor="white",
                                font=dict(color="#333"),
                                legend=dict(orientation="h", y=1.05, x=0, font=dict(size=8)))
            _mb_h = pio.to_html(_mb_p, include_plotlyjs=False, full_html=False,
                                config={"displayModeBar": False})

            # Gantt: full chiều cao, hiển thị đủ tên tất cả dự án
            _fg_print = go.Figure(fig_g)
            _pnames_g = df_g["Dự án"].tolist()
            _fg_print.update_layout(
                width=_CENTER_W, height=_GANTT_H,
                paper_bgcolor="white", plot_bgcolor="white",
                font=dict(color="#333"),
                yaxis=dict(
                    autorange="reversed",
                    tickmode="array",
                    tickvals=_pnames_g,
                    ticktext=_pnames_g,
                    tickfont=dict(size=9, color="#333"),
                    automargin=True,
                ),
                legend=dict(orientation="h", y=1.02, x=0, font=dict(size=9)),
                margin=dict(l=10, r=10, t=45, b=70),
            )
            _fg_h = pio.to_html(_fg_print, include_plotlyjs=False, full_html=False,
                                config={"displayModeBar": False})

            # ── Biểu đồ Năng Lực SX (nếu đã tải file SX) ─────────────────────
            _fstage_h = ""
            _sx_print = st.session_state.get("sx_data")
            if _sx_print and _sx_print.get("stage_caps"):
                _bns_p = _sx_print.get("bottlenecks") or [_sx_print.get("bottleneck","")]
                _cap_p = _sx_print["cap_monthly"] or 0
                _sc_names = [s["cong_doan"] for s in _sx_print["stage_caps"]]
                _sc_caps  = [s["cap_m2_thang"] for s in _sx_print["stage_caps"]]
                _sc_colors = ["#e17055" if n in _bns_p else "#00b894" for n in _sc_names]
                _fstage_p = go.Figure()
                _fstage_p.add_trace(go.Bar(
                    x=_sc_names, y=_sc_caps,
                    marker_color=_sc_colors,
                    text=[f"{v:,}" for v in _sc_caps],
                    textposition="outside",
                    hovertemplate="<b>%{x}</b><br>%{y:,} m²/tháng<extra></extra>",
                ))
                _fstage_p.add_hline(y=_cap_p, line_dash="dash", line_color="#e17055", line_width=1.5,
                                    annotation_text=f"  Nút thắt: {_cap_p:,} m²",
                                    annotation_font=dict(color="#e17055", size=9))
                _fstage_p.update_layout(
                    height=300, showlegend=False,
                    margin=dict(t=20, r=120, b=50, l=10),
                    paper_bgcolor="white", plot_bgcolor="white",
                    font=dict(color="#333", size=9),
                    xaxis=dict(tickfont=dict(size=8)),
                    yaxis=dict(title=dict(text="m²/tháng", font=dict(size=8))),
                )
                _fstage_h = pio.to_html(_fstage_p, include_plotlyjs=False, full_html=False,
                                        config={"displayModeBar": False})

            html_out = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>Dashboard KHSX – QDP</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
  @page {{ size: A3 landscape; margin: 8mm; }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Inter, sans-serif; background: #fff;
          color: #333; font-size: 11px; height: 100%; }}
  /* ── Banner ── */
  .banner {{ background: linear-gradient(135deg,#6c5ce7,#a29bfe); color:#fff;
             padding: 7px 16px; border-radius: 10px; margin-bottom: 6px;
             display:flex; align-items:center; gap:10px; }}
  .banner h1 {{ font-size:1.05rem; font-weight:700; margin:0; }}
  .banner p  {{ font-size:0.70rem; opacity:.88; margin:0; }}
  /* ── KPI bar ── */
  .kpi-bar {{ display:flex; gap:6px; margin-bottom:6px; }}
  .kc {{ flex:1; background:#fff; border-radius:8px; padding:5px 10px;
          box-shadow:0 1px 4px rgba(0,0,0,.10); border-top:3px solid #6c5ce7;
          display:flex; flex-direction:column; justify-content:center; }}
  .kc.g {{ border-top-color:#00b894; }}
  .kc.o {{ border-top-color:#fd9644; }}
  .kc.r {{ border-top-color:#e17055; }}
  .kc.y {{ border-top-color:#e1b000; }}
  .kc.c {{ border-top-color:#00cec9; }}
  .kl {{ font-size:.52rem; color:#555; font-weight:700; text-transform:uppercase;
          letter-spacing:.4px; }}
  .kv {{ font-size:1.15rem; font-weight:800; line-height:1.1; }}
  .ks {{ font-size:.50rem; color:#777; }}
  /* ── 3-column main layout ── */
  .main {{ display:flex; gap:8px; align-items:stretch; }}
  .col-side {{ width:{_SIDE_W}px; flex-shrink:0; display:flex;
               flex-direction:column; gap:6px; }}
  .col-center {{ flex:1; display:flex; flex-direction:column; }}
  .slabel {{ font-size:.65rem; font-weight:700; color:#444;
             margin-bottom:2px; padding-left:2px; }}
  .caption {{ font-size:.55rem; color:#666; margin-top:2px; }}
  .plotly-graph-div {{ max-width:100% !important; }}
  .chart-box {{ background:#fff; border-radius:8px;
                box-shadow:0 1px 4px rgba(0,0,0,.08); padding:4px; }}
</style>
</head>
<body>
<!-- Banner -->
<div class="banner">
  <div style="font-size:1.6rem">🏭</div>
  <div style="flex:1">
    <h1>Kế Hoạch Sản Xuất – Nhà Máy QDP <span style="font-size:0.5em;color:#aaaaaa;font-weight:400">v5.6.26</span></h1>
    <p>Dashboard quản lý tải trọng &amp; tiến độ sản xuất nhôm kính &nbsp;|&nbsp; In ngày: {td.strftime("%d/%m/%Y")}</p>
  </div>
</div>
<!-- KPI bar compact -->
<div class="kpi-bar">
  <div class="kc"><div class="kl">Tổng dự án</div><div class="kv" style="color:#6c5ce7">{n_proj}</div><div class="ks">đang theo dõi</div></div>
  <div class="kc g"><div class="kl">Đã ký HĐ</div><div class="kv" style="color:#00b894">{signed}</div><div class="ks">xác suất 100%</div></div>
  <div class="kc o"><div class="kl">Đang SX</div><div class="kv" style="color:#fd9644">{in_progress}</div><div class="ks">tháng {td.strftime("%m/%Y")}</div></div>
  <div class="kc c"><div class="kl">Tổng KL</div><div class="kv" style="color:#00cec9">{total_m2:,.0f}</div><div class="ks">m² kỳ vọng</div></div>
  <div class="kc r"><div class="kl">Quá tải</div><div class="kv" style="color:#e17055">{n_over}</div><div class="ks">tháng ≥100%</div></div>
  <div class="kc o"><div class="kl">Căng TĐ</div><div class="kv" style="color:#fd9644">{n_tight}</div><div class="ks">90–100%</div></div>
  <div class="kc y"><div class="kl">Cần chú ý</div><div class="kv" style="color:#c49000">{n_warn}</div><div class="ks">70–90%</div></div>
  <div class="kc g"><div class="kl">Bình thường</div><div class="kv" style="color:#00b894">{n_ok}</div><div class="ks">&lt;70%</div></div>
  <div class="kc"><div class="kl">Tải TB</div><div class="kv" style="color:#6c5ce7">{avg_u}%</div><div class="ks">năng suất 18T</div></div>
  <div class="kc g"><div class="kl">Năng lực SX</div><div class="kv" style="color:#00b894">{cap_monthly:,}</div><div class="ks">m²/tháng · {" · ".join((_sx_print.get("bottlenecks") or [_sx_print["bottleneck"]]) if _sx_print else ["mặc định"])}</div></div>
</div>
<!-- 3-column layout -->
<div class="main">
  <!-- Cột trái: Tải trọng + Nhân lực + Mặt bằng -->
  <div class="col-side">
    <div class="chart-box">
      <div class="slabel">📊 Tải trọng nhà máy 18 tháng tới</div>
      {_fb_h}
      <div class="caption">🟢 &lt;70% bình thường · 🟡 70–90% cần chú ý · 🟠 90–100% căng · 🔴 ≥100% quá tải</div>
    </div>
    <div class="chart-box">
      <div class="slabel">👥 Nhân lực kỹ thuật gián tiếp theo tuần</div>
      {_nl_h}
    </div>
    <div class="chart-box">
      <div class="slabel">🏗️ Mặt bằng xưởng &amp; kho TP theo tuần</div>
      {_mb_h}
    </div>
  </div>
  <!-- Cột giữa: Gantt -->
  <div class="col-center">
    <div class="chart-box" style="flex:1">
      <div class="slabel">📅 Tiến độ dự án (Gantt) – {len(projs)} dự án</div>
      {_fg_h}
      <div class="caption">🟢 Đã ký HĐ · 🟣 Khả năng cao · 🟠 Đang xét · 🔴 Hôm nay</div>
    </div>
  </div>
  <!-- Cột phải: Pie + NLSX -->
  <div class="col-side">
    <div class="chart-box">
      <div class="slabel">🍩 Phân bổ KL theo XS%</div>
      {_fp_h}
    </div>
    {'<div class="chart-box" style="flex:1"><div class="slabel">🔧 Năng lực SX theo công đoạn (m²/tháng)</div>' + _fstage_h + '<div class="caption">🔴 Nút thắt · 🟢 Bình thường</div></div>' if _fstage_h else ''}
  </div>
</div>
</body>
</html>"""

            st.download_button(
                "💾 Tải file HTML (mở bằng Chrome → Ctrl+P → A3 Landscape → PDF)",
                data=html_out.encode("utf-8"),
                file_name=f"dashboard_QDP_{td.strftime('%m%Y')}.html",
                mime="text/html",
            )

# ══════════════════════════════════════════════════════════════════════════════
# TAB KẾ HOẠCH SẢN XUẤT (toàn bộ giao diện gốc, giữ nguyên)
# ══════════════════════════════════════════════════════════════════════════════
with tab_plan:
    st.header("📊 Kế hoạch sản xuất nhà máy QDP")

    # ══════════════════════════════════════════════════════════════════════════════
    # Planning Dashboard
    # ══════════════════════════════════════════════════════════════════════════════
    if True:
        uploaded = st.file_uploader(
            "📂 Tải lên file Excel (Dự án input · Tech input · Mặt bằng yêu cầu)",
            type=["xlsx", "xls"], key="plan_file"
        )

        if uploaded is None:
            st.info("⬅️ Vui lòng tải lên file chính để bắt đầu.")
        else:
            xl = pd.ExcelFile(uploaded)

        if uploaded is not None:
            tech_weekly = parse_tech_input(xl)
            st.session_state["t1_tech_weekly"] = tech_weekly
            if "Dự án input" in xl.sheet_names:
                sheet = "Dự án input"
            else:
                sheet = st.selectbox("Chọn sheet dữ liệu", xl.sheet_names, key="plan_sheet")
            try:
                df_raw = xl.parse(sheet, header=None)
            except Exception as e:
                st.error(f"Lỗi đọc sheet: {e}")
                df_raw = None

            if df_raw is not None:
                projects, skipped_rows = parse_projects(df_raw)

                if skipped_rows:
                    st.warning(
                        f"⚠️ **Có {len(skipped_rows)} dòng bị BỎ QUA vì thiếu dữ liệu — "
                        "các dự án này sẽ KHÔNG xuất hiện trong kế hoạch:**\n\n"
                        + "\n".join(f"- **{r['name']}** — {r['reason']}" for r in skipped_rows)
                        + "\n\nVui lòng bổ sung đầy đủ **Tên dự án, Khối lượng (m²), "
                        "Ngày bắt đầu, Ngày hoàn thành** trong file Excel rồi tải lại."
                    )

                if not projects:
                    st.error("Không parse được dự án nào. Kiểm tra lại file.")
                else:
                    file_id = f"{uploaded.name}_{uploaded.size}_{sheet}"
                    if st.session_state.get("t1_file_id") != file_id:
                        orig_rows = [{
                            "name":  p["name"],
                            "m2":    round(p["m2"], 4),
                            "start": date(p["start"].year, p["start"].month, 1),
                            "end":   date(p["end"].year, p["end"].month, 1),
                            "prob":  p["prob"],
                        } for p in projects]
                        st.session_state.t1_file_id        = file_id
                        st.session_state.t1_orig           = orig_rows
                        st.session_state.t1_rows           = [dict(r) for r in orig_rows]
                        st.session_state.t1_editor_ver     = st.session_state.get("t1_editor_ver", 0) + 1
                        st.session_state.t1_reset_ask      = False
                        st.session_state.t1_custom         = {}
                        st.session_state.t1_weekly_custom  = {}
                        st.rerun()  # Dashboard tab cần chạy lại để đọc t1_rows vừa khởi tạo

                    rows     = st.session_state.t1_rows
                    orig_map = {r["name"]: r for r in st.session_state.t1_orig}

                    # ── Toolbar ───────────────────────────────────────────────
                    st.subheader("⚙️ Quản lý danh sách dự án")
                    tb1, tb2, tb3 = st.columns([1, 1, 4])

                    with tb1:
                        exp_rows = [{
                            "Tên dự án":       r["name"],
                            "Khối lượng (m²)": r["m2"],
                            "Bắt đầu":         r["start"].strftime("%d/%m/%Y"),
                            "Kết thúc":        r["end"].strftime("%d/%m/%Y"),
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
                                    st.session_state.t1_rows           = [dict(r) for r in st.session_state.t1_orig]
                                    st.session_state.t1_editor_ver    += 1
                                    st.session_state.t1_reset_ask      = False
                                    st.session_state.t1_custom         = {}
                                    st.session_state.t1_weekly_custom  = {}
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
                                "Bắt đầu SX", format="DD/MM/YYYY",
                            ),
                            "Kết thúc":        st.column_config.DateColumn(
                                "Kết thúc SX", format="DD/MM/YYYY",
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
                                or df_row["Bắt đầu"] != orig["start"].strftime("%d/%m/%Y")
                                or df_row["Kết thúc"] != orig["end"].strftime("%d/%m/%Y")
                                or df_row["XS%"] != orig["prob"]
                            )
                            if changed:
                                return ["background-color:#fff3cd;color:#856404"] * len(df_row)
                            return ["color:#333333"] * len(df_row)

                        display_df = pd.DataFrame([{
                            "Tên dự án":       r["name"],
                            "Khối lượng (m²)": r["m2"],
                            "Bắt đầu":         r["start"].strftime("%d/%m/%Y"),
                            "Kết thúc":        r["end"].strftime("%d/%m/%Y"),
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

                    # ── Tùy chỉnh m²/tháng từng dự án ────────────────────────────
                    with st.expander("📅 Tùy chỉnh m²/tháng từng dự án", expanded=False):
                        cust_proj_sel = st.selectbox(
                            "Chọn dự án:", [p["name"] for p in projects], key="t1_csel"
                        )
                        p_csel = next(p for p in projects if p["name"] == cust_proj_sel)
                        cust_months = months_range(p_csel["start"], p_csel["end"])
                        _cur_custom  = st.session_state.get("t1_custom", {})
                        _p_cust      = _cur_custom.get(cust_proj_sel, {})
                        _n_cm        = len(cust_months)
                        _base_m      = round(p_csel["m2"] / _n_cm, 1) if _n_cm else 0
                        _cust_rows   = [
                            {"Tháng": m.strftime("%m/%Y"),
                             "m² kế hoạch": _p_cust.get(m.strftime("%m/%Y"), _base_m)}
                            for m in cust_months
                        ]
                        _cc1, _cc2 = st.columns([3, 1])
                        with _cc1:
                            _edited_c = st.data_editor(
                                pd.DataFrame(_cust_rows),
                                column_config={
                                    "Tháng":        st.column_config.TextColumn("Tháng", disabled=True),
                                    "m² kế hoạch":  st.column_config.NumberColumn(
                                        "m² kế hoạch", min_value=0, format="%,.1f"
                                    ),
                                },
                                hide_index=True, use_container_width=True,
                                num_rows="fixed",
                                key=f"t1_ced_{cust_proj_sel}",
                            )
                        with _cc2:
                            _tot_c = _edited_c["m² kế hoạch"].sum()
                            _diff_c = abs(p_csel["m2"] - _tot_c)
                            st.metric("Tổng nhập", f"{_tot_c:,.1f} m²")
                            st.metric("Tổng HĐ",   f"{p_csel['m2']:,.0f} m²")
                            if _diff_c < 0.5:
                                st.success("✅ Khớp")
                            else:
                                st.warning(f"⚠️ Chênh {_diff_c:,.1f} m²")
                            if st.button("💾 Lưu", key=f"t1_csave_{cust_proj_sel}"):
                                if "t1_custom" not in st.session_state:
                                    st.session_state.t1_custom = {}
                                st.session_state.t1_custom[cust_proj_sel] = {
                                    row["Tháng"]: row["m² kế hoạch"]
                                    for _, row in _edited_c.iterrows()
                                }
                                st.rerun()
                            if _diff_c >= 0.1 and st.button("⚖️ Bù chênh đều", key=f"t1_cauto_{cust_proj_sel}",
                                                              help="Giữ các tháng đã chỉnh, chia đều phần chênh lệch cho các tháng còn lại"):
                                # Tháng "tự do" = vẫn ở giá trị gốc (chưa chỉnh tay)
                                vals = _edited_c["m² kế hoạch"].copy()
                                pinned = abs(vals - _base_m) > 0.05
                                free   = ~pinned
                                diff_c = p_csel["m2"] - vals.sum()
                                n_free = int(free.sum())
                                if n_free > 0:
                                    vals[free] = (vals[free] + diff_c / n_free).round(1)
                                else:
                                    # Tất cả đã chỉnh → chia đều cho tất cả
                                    vals = (vals + diff_c / len(vals)).round(1)
                                # Bù sai số làm tròn vào tháng cuối
                                vals.iloc[-1] = round(p_csel["m2"] - vals.iloc[:-1].sum(), 1)
                                if "t1_custom" not in st.session_state:
                                    st.session_state.t1_custom = {}
                                st.session_state.t1_custom[cust_proj_sel] = {
                                    r["Tháng"]: float(v)
                                    for r, v in zip(_edited_c.to_dict("records"), vals)
                                }
                                st.rerun()
                            if _p_cust and st.button("⚡ Phân bổ đều", key=f"t1_cclear_{cust_proj_sel}"):
                                st.session_state.t1_custom.pop(cust_proj_sel, None)
                                st.rerun()
                        _cust_list = [nm for nm in st.session_state.get("t1_custom", {})]
                        if _cust_list:
                            st.caption("📌 Đã tùy chỉnh tháng: " + " · ".join(f"**{nm}**" for nm in _cust_list))

                    st.divider()

                    # ── Load matrix: load_data[month_label][project_name] = m² ──
                    # Xây đồng thời 2 ma trận: kỳ vọng (×XS%) và xấu nhất (100%)
                    load_data_w    = {ml: {} for ml in month_labels}  # kỳ vọng
                    load_data_full = {ml: {} for ml in month_labels}  # xấu nhất
                    _t1_custom = st.session_state.get("t1_custom", {})
                    for p in projects:
                        p_custom = _t1_custom.get(p["name"], {})
                        if p_custom:
                            for ml, v in p_custom.items():
                                if ml in load_data_w:
                                    load_data_w[ml][p["name"]]    = round(v * (p["prob"] / 100), 1)
                                    load_data_full[ml][p["name"]] = round(v, 1)
                        else:
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

                    # ── Lưu BASE peak demand vào session state ──────────────
                    _monthly_totals_w = {ml: round(sum(load_data_w[ml].values())) for ml in month_labels}
                    _peak_month_label = max(_monthly_totals_w, key=_monthly_totals_w.get) if _monthly_totals_w else None
                    _peak_demand      = _monthly_totals_w.get(_peak_month_label, 0)
                    _overload_months  = [ml for ml, v in _monthly_totals_w.items() if v > cap_monthly]
                    # Lưu base (luôn cập nhật theo file)
                    st.session_state["khsx_base_peak"]         = _peak_demand
                    st.session_state["khsx_base_month"]        = _peak_month_label
                    st.session_state["khsx_base_overload"]     = _overload_months
                    st.session_state["khsx_base_totals"]       = _monthly_totals_w
                    st.session_state["khsx_cap_monthly"]       = cap_monthly
                    # Effective = max(base, whatif) — whatif chỉ bị ghi đè khi click button
                    _wif_peak_saved = st.session_state.get("khsx_whatif_peak", 0)
                    if _peak_demand >= _wif_peak_saved:
                        # Base lớn hơn what-if → reset whatif
                        st.session_state["khsx_whatif_peak"]    = 0
                        st.session_state["khsx_whatif_totals"]  = {}
                    # Effective values for warnings
                    _eff_totals  = st.session_state.get("khsx_whatif_totals") or _monthly_totals_w
                    _eff_peak    = max(_peak_demand, st.session_state.get("khsx_whatif_peak", 0))
                    _eff_month   = st.session_state.get("khsx_whatif_month", _peak_month_label) if _eff_peak > _peak_demand else _peak_month_label
                    _eff_over    = [ml for ml, v in _eff_totals.items() if v > cap_monthly]
                    st.session_state["khsx_peak_demand"]     = _eff_peak
                    st.session_state["khsx_peak_month"]      = _eff_month
                    st.session_state["khsx_overload_months"] = _eff_over
                    st.session_state["khsx_monthly_totals"]  = _eff_totals

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
                            "Bắt đầu SX":     p["start"].strftime("%d/%m/%Y"),
                            "Kết thúc SX":    p["end"].strftime("%d/%m/%Y"),
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

                    # Trục X bao phủ toàn bộ dữ liệu để hiển thị đủ tất cả dự án
                    all_starts = [p["start"] for p in projects]
                    all_ends   = [p["end"]   for p in projects]
                    data_x_start = min(all_starts).strftime("%Y-%m-%d")
                    data_x_end   = next_month(max(all_ends)).strftime("%Y-%m-%d")

                    # Đánh dấu ranh giới cuối tầm nhìn kế hoạch
                    horizon_end_dt = next_month(horizon_months[-1])
                    fig_gantt1.add_vline(
                        x=pd.Timestamp(horizon_end_dt).timestamp() * 1000,
                        line_color="gray",
                        line_width=1.5,
                        line_dash="dot",
                        annotation_text=f"  Cuối tầm nhìn ({horizon_end_dt.strftime('%m/%Y')})",
                        annotation_font_color="gray",
                        annotation_position="top right",
                        annotation_font_size=10,
                    )

                    all_project_names = df_gantt["Dự án"].tolist()
                    fig_gantt1.update_xaxes(
                        range=[data_x_start, data_x_end],
                        dtick="M1",
                        tickformat="%m/%Y",
                        tickangle=45,
                        tickfont=dict(size=8),
                    )
                    fig_gantt1.update_layout(
                        yaxis=dict(
                            autorange="reversed",
                            tickmode="array",
                            tickvals=all_project_names,
                            ticktext=all_project_names,
                            tickfont=dict(size=9),
                            automargin=True,
                        ),
                        height=max(380, len(projects) * 28 + 100),
                        margin=dict(l=10, r=20, t=36, b=60),
                        legend=dict(
                            title="Xác suất",
                            orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0,
                            font=dict(size=10),
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
                        sim_load  = load_data_w
                        sim_label = "Tải nền kỳ vọng (m²)"

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

                                # ── Lưu what-if peak vào key RIÊNG — không bị rerun ghi đè ──
                                _sim_monthly_totals = {r["Tháng"]: r["Tổng (m²)"] for r in sim_rows}
                                _sim_peak_month = max(_sim_monthly_totals, key=_sim_monthly_totals.get) if _sim_monthly_totals else None
                                _sim_peak       = _sim_monthly_totals.get(_sim_peak_month, 0)
                                _sim_over       = [ml for ml, v in _sim_monthly_totals.items() if v > cap_monthly]
                                # Luôn lưu what-if riêng (kể cả nhỏ hơn base)
                                st.session_state["khsx_whatif_peak"]    = _sim_peak
                                st.session_state["khsx_whatif_month"]   = _sim_peak_month
                                st.session_state["khsx_whatif_totals"]  = _sim_monthly_totals
                                # Cập nhật effective values
                                st.session_state["khsx_peak_demand"]     = _sim_peak
                                st.session_state["khsx_peak_month"]      = _sim_peak_month
                                st.session_state["khsx_overload_months"] = _sim_over
                                st.session_state["khsx_monthly_totals"]  = _sim_monthly_totals
                                if _sim_peak > st.session_state.get("khsx_base_peak", 0):
                                    st.info(f"📊 Đã cập nhật sang tab **Năng Lực SX** — peak mới: **{_sim_peak:,} m²/tháng**")

                    # (Download ở cuối trang, sau khi weekly_load được tính)

                    # ══════════════════════════════════════════════════════════════════
                    # PHẦN KẾ HOẠCH THEO TUẦN
                    # ══════════════════════════════════════════════════════════════════
                    st.divider()
                    st.subheader("📅 Kế hoạch & Cảnh báo theo tuần")

                    weekly_load, week_labels_all, week_dates_all = build_weekly_production(
                        projects, show_weighted,
                        weekly_custom=st.session_state.get("t1_weekly_custom", {})
                    )

                    if not week_labels_all:
                        st.warning("Không có dữ liệu tuần.")
                    else:
                        # Giới hạn theo horizon (cùng tầm nhìn với biểu đồ tháng)
                        cutoff = pd.Timestamp(today) + pd.DateOffset(months=horizon)
                        wl_wd_pairs = [
                            (wl, wd) for wl, wd in zip(week_labels_all, week_dates_all)
                            if wd <= cutoff
                        ] or list(zip(week_labels_all, week_dates_all))
                        w_labels = [p[0] for p in wl_wd_pairs]

                        # ── Tùy chỉnh m²/tuần từng dự án ────────────────────────────
                        with st.expander("📅 Tùy chỉnh m²/tuần từng dự án", expanded=False):
                            wk_proj_sel = st.selectbox(
                                "Chọn dự án:", [p["name"] for p in projects], key="t1_wsel"
                            )
                            p_wsel = next(p for p in projects if p["name"] == wk_proj_sel)
                            # Tính tuần của dự án này
                            p_start = pd.Timestamp(p_wsel["start"])
                            p_end   = pd.Timestamp(p_wsel["end"]) + pd.DateOffset(months=1) - pd.Timedelta(days=1)
                            proj_weeks_w = [
                                (d, f"{d.isocalendar()[0]}-W{d.isocalendar()[1]:02d}", d.strftime("%d/%m/%Y"))
                                for d in week_dates_all
                                if p_start <= d + pd.Timedelta(days=6) and d <= p_end
                            ]
                            n_weeks = len(proj_weeks_w)
                            _base_w = round(p_wsel["m2"] / n_weeks, 1) if n_weeks else 0
                            _cur_custom_w = st.session_state.get("t1_weekly_custom", {})
                            _p_cust_w = _cur_custom_w.get(wk_proj_sel, {})
                            _wk_rows = [
                                {"Tuần": wl, "Ngày bắt đầu": wd, "m² kế hoạch": _p_cust_w.get(wl, _base_w)}
                                for _, wl, wd in proj_weeks_w
                            ]
                            _wc1, _wc2 = st.columns([3, 1])
                            with _wc1:
                                _edited_w = st.data_editor(
                                    pd.DataFrame(_wk_rows),
                                    column_config={
                                        "Tuần":       st.column_config.TextColumn("Tuần", disabled=True),
                                        "Ngày bắt đầu": st.column_config.TextColumn("Ngày bắt đầu", disabled=True),
                                        "m² kế hoạch": st.column_config.NumberColumn(
                                            "m² kế hoạch", min_value=0, format="%,.1f"
                                        ),
                                    },
                                    hide_index=True, use_container_width=True,
                                    num_rows="fixed",
                                    key=f"t1_wed_{wk_proj_sel}",
                                )
                            with _wc2:
                                _tot_w = _edited_w["m² kế hoạch"].sum()
                                _diff_w = abs(p_wsel["m2"] - _tot_w)
                                st.metric("Tổng nhập", f"{_tot_w:,.1f} m²")
                                st.metric("Tổng HĐ",   f"{p_wsel['m2']:,.0f} m²")
                                if _diff_w < 0.5:
                                    st.success("✅ Khớp")
                                else:
                                    st.warning(f"⚠️ Chênh {_diff_w:,.1f} m²")
                                if st.button("💾 Lưu", key=f"t1_wsave_{wk_proj_sel}"):
                                    if "t1_weekly_custom" not in st.session_state:
                                        st.session_state.t1_weekly_custom = {}
                                    st.session_state.t1_weekly_custom[wk_proj_sel] = {
                                        row["Tuần"]: row["m² kế hoạch"]
                                        for _, row in _edited_w.iterrows()
                                    }
                                    st.rerun()
                                if _diff_w >= 0.1 and st.button("⚖️ Bù chênh đều", key=f"t1_wauto_{wk_proj_sel}",
                                                                  help="Giữ các tuần đã chỉnh, chia đều phần chênh lệch cho các tuần còn lại"):
                                    vals_w = _edited_w["m² kế hoạch"].copy()
                                    pinned_w = abs(vals_w - _base_w) > 0.05
                                    free_w   = ~pinned_w
                                    diff_w   = p_wsel["m2"] - vals_w.sum()
                                    n_free_w = int(free_w.sum())
                                    if n_free_w > 0:
                                        vals_w[free_w] = (vals_w[free_w] + diff_w / n_free_w).round(1)
                                    else:
                                        vals_w = (vals_w + diff_w / len(vals_w)).round(1)
                                    vals_w.iloc[-1] = round(p_wsel["m2"] - vals_w.iloc[:-1].sum(), 1)
                                    if "t1_weekly_custom" not in st.session_state:
                                        st.session_state.t1_weekly_custom = {}
                                    st.session_state.t1_weekly_custom[wk_proj_sel] = {
                                        r["Tuần"]: float(v)
                                        for r, v in zip(_edited_w.to_dict("records"), vals_w)
                                    }
                                    st.rerun()
                                if _p_cust_w and st.button("⚡ Phân bổ đều", key=f"t1_wclear_{wk_proj_sel}"):
                                    st.session_state.t1_weekly_custom.pop(wk_proj_sel, None)
                                    st.rerun()

                        # ── Biểu đồ 1: Tải SX theo tuần ──────────────────────────
                        st.markdown("**🏭 Tải sản xuất theo tuần**")
                        w_totals  = [sum(weekly_load.get(wl, {}).values()) for wl in w_labels]
                        y_max_w   = max(max(w_totals) * 1.15, cap_weekly * 1.15) if w_totals else cap_weekly * 1.2

                        fig_w = go.Figure()
                        fig_w.add_hrect(y0=0,              y1=cap_weekly * 0.7,
                                        fillcolor="rgba(0,180,0,0.06)",   line_width=0)
                        fig_w.add_hrect(y0=cap_weekly * 0.7, y1=cap_weekly * 0.9,
                                        fillcolor="rgba(255,200,0,0.12)", line_width=0)
                        fig_w.add_hrect(y0=cap_weekly * 0.9, y1=y_max_w,
                                        fillcolor="rgba(255,60,0,0.09)",  line_width=0)

                        for i, p in enumerate(projects):
                            y_vals = [weekly_load.get(wl, {}).get(p["name"], 0) for wl in w_labels]
                            if any(v > 0 for v in y_vals):
                                fig_w.add_trace(go.Bar(
                                    name=p["name"],
                                    x=w_labels,
                                    y=y_vals,
                                    marker_color=COLORS[i % len(COLORS)],
                                    hovertemplate="<b>%{x}</b><br>" + p["name"] + ": %{y:,.0f} m²<extra></extra>",
                                ))

                        fig_w.add_hline(
                            y=cap_weekly, line_dash="dash", line_color="red", line_width=2,
                            annotation_text=f"  Tối đa: {cap_weekly:,} m²/tuần",
                            annotation_font_color="red", annotation_position="right",
                        )
                        fig_w.add_hline(
                            y=cap_weekly * 0.9, line_dash="dot", line_color="darkorange", line_width=1.5,
                            annotation_text="  Cảnh báo 90%",
                            annotation_font_color="darkorange", annotation_position="right",
                        )
                        n_overload_w = sum(1 for v in w_totals if v > cap_weekly)
                        if n_overload_w:
                            st.warning(f"⚠️ **{n_overload_w} tuần** vượt năng suất tối đa ({cap_weekly:,} m²/tuần).")
                        fig_w.update_layout(
                            barmode="stack", height=480,
                            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
                            xaxis_title="Tuần", yaxis_title="Khối lượng sản xuất (m²)",
                            yaxis=dict(range=[0, y_max_w]),
                            xaxis=dict(tickangle=45, tickfont=dict(size=8)),
                            margin=dict(t=20, r=200, b=80, l=70),
                        )
                        st.plotly_chart(fig_w, use_container_width=True)

                        # ── Biểu đồ 2: Nhân lực kỹ thuật gián tiếp ───────────────
                        st.markdown("**👥 Nhân lực kỹ thuật gián tiếp theo tuần**")
                        if tech_weekly:
                            w_eng = [tech_weekly.get(wl, 0) for wl in w_labels]
                            bar_colors_eng = []
                            for v in w_eng:
                                pct_e = v / max_indirect * 100 if max_indirect > 0 else 0
                                if pct_e >= 100:   bar_colors_eng.append("#ff4d4d")
                                elif pct_e >= 90:  bar_colors_eng.append("#ff8c00")
                                elif pct_e >= 70:  bar_colors_eng.append("#ffd700")
                                else:              bar_colors_eng.append("#2ca02c")

                            n_overload_eng = sum(1 for v in w_eng if v > max_indirect)
                            if n_overload_eng:
                                st.warning(
                                    f"⚠️ **{n_overload_eng} tuần** vượt ngưỡng nhân lực gián tiếp "
                                    f"({max_indirect} người)."
                                )

                            fig_eng = go.Figure()
                            fig_eng.add_trace(go.Bar(
                                x=w_labels, y=w_eng,
                                marker_color=bar_colors_eng,
                                hovertemplate="<b>%{x}</b><br>%{y:.0f} người<extra></extra>",
                            ))
                            fig_eng.add_hline(
                                y=max_indirect, line_dash="dash", line_color="red", line_width=2,
                                annotation_text=f"  Tối đa: {max_indirect} người",
                                annotation_font_color="red", annotation_position="right",
                            )
                            fig_eng.update_layout(
                                height=350, showlegend=False,
                                xaxis_title="Tuần", yaxis_title="Số kỹ sư",
                                xaxis=dict(tickangle=45, tickfont=dict(size=8)),
                                margin=dict(t=20, r=200, b=80, l=60),
                            )
                            st.plotly_chart(fig_eng, use_container_width=True)
                        else:
                            st.info(
                                "ℹ️ File chính không có sheet **'Tech input'** — "
                                "bỏ qua biểu đồ nhân lực gián tiếp."
                            )

                        # ── Biểu đồ 3: Mặt bằng & Kho theo tuần ─────────────────
                        st.markdown("**🏗️ Diện tích mặt bằng & kho theo tuần**")
                        xuong_vals = [
                            sum(weekly_load.get(wl, {}).values()) * coeff_xuong
                            for wl in w_labels
                        ]
                        kho_tp_vals = [
                            sum(weekly_load.get(wl, {}).values()) * coeff_kho_tp
                            for wl in w_labels
                        ]
                        fig_mb = go.Figure()
                        area_cfg_mb = [
                            (xuong_vals,  avail_xuong,  "#1f77b4", "Xưởng SX"),
                            (kho_tp_vals, avail_kho_tp, "#2ca02c", "Kho TP"),
                        ]
                        warnings_mb = []
                        for vals, avail, color, label in area_cfg_mb:
                            fig_mb.add_trace(go.Scatter(
                                x=w_labels, y=vals, mode="lines+markers",
                                name=f"{label} (yêu cầu)",
                                line=dict(color=color, width=2),
                                marker=dict(size=5),
                                hovertemplate=f"<b>%{{x}}</b><br>{label}: %{{y:,.0f}} m²<extra></extra>",
                            ))
                            fig_mb.add_hline(
                                y=avail, line_dash="dash", line_color=color, line_width=1.5,
                                annotation_text=f"  {label} hiện có: {avail:,} m²",
                                annotation_font_color=color, annotation_position="right",
                            )
                            n_exceed = sum(1 for v in vals if v > avail)
                            if n_exceed:
                                warnings_mb.append(
                                    f"⚠️ **{label}**: {n_exceed} tuần yêu cầu vượt diện tích hiện có "
                                    f"({avail:,} m²)."
                                )
                        for w in warnings_mb:
                            st.warning(w)
                        fig_mb.update_layout(
                            height=400,
                            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
                            xaxis_title="Tuần", yaxis_title="Diện tích (m²)",
                            xaxis=dict(tickangle=45, tickfont=dict(size=8)),
                            margin=dict(t=20, r=220, b=80, l=70),
                        )
                        st.plotly_chart(fig_mb, use_container_width=True)

                    # ── Xuất Excel đầy đủ ────────────────────────────────────────
                    st.divider()
                    _exp_proj = [{
                        "Tên dự án":       r["name"],
                        "Khối lượng (m²)": r["m2"],
                        "Bắt đầu":         r["start"].strftime("%d/%m/%Y"),
                        "Kết thúc":        r["end"].strftime("%d/%m/%Y"),
                        "XS%":             r["prob"],
                    } for r in st.session_state.t1_rows]

                    _full_buf = io.BytesIO()
                    with pd.ExcelWriter(_full_buf, engine="openpyxl") as _writer:
                        # Sheet 1: dự án (có thể đã chỉnh sửa trong app)
                        pd.DataFrame(_exp_proj).to_excel(
                            _writer, sheet_name="Dự án input", index=False
                        )
                        # Sheet 2 & 3: copy nguyên từ file input chính
                        for _sn in ("Tech input", "Mặt bằng yêu cầu"):
                            if _sn in xl.sheet_names:
                                xl.parse(_sn).to_excel(
                                    _writer, sheet_name=_sn, index=False
                                )
                        # Sheet 4: tổng hợp tháng (app tính)
                        df_summary.to_excel(
                            _writer, sheet_name="Tổng hợp tháng", index=False
                        )
                        # Sheet 5-6: Năng lực SX (nếu đã tải file SX)
                        _sx = st.session_state.get("sx_data")
                        if _sx:
                            if _sx["may_moc"]:
                                pd.DataFrame(_sx["may_moc"]).rename(columns={
                                    "ma": "Mã máy", "ten": "Tên máy", "sl": "SL máy",
                                    "sl_nguoi": "NV/máy", "nang_suat_m2_gio": "NS m²/máy/giờ",
                                    "ca": "Ca/ngày", "gio_ca": "Giờ/ca",
                                    "hs_pct": "Hiệu suất%", "cap_m2_thang": "Năng suất (m²/tháng)"
                                }).to_excel(_writer, sheet_name="Máy móc TB", index=False)
                            if _sx["nhan_luc"]:
                                _df_nl_export = pd.DataFrame(_sx["nhan_luc"])
                                _df_nl_export["nang_suat_m2_gio"] = _df_nl_export["nang_suat_m2_gio"].apply(
                                    lambda x: x if x is not None else "theo máy"
                                )
                                _df_nl_export.rename(columns={
                                    "to": "Tổ", "cong_doan": "Công đoạn",
                                    "nv": "Số NV", "nang_suat_m2_gio": "NS m²/người/giờ",
                                    "ca": "Ca/ngày", "gio_ca": "Giờ/ca",
                                    "hs_pct": "Hiệu suất%", "cap_m2_thang": "Năng suất (m²/tháng)"
                                }).to_excel(_writer, sheet_name="Nhân lực", index=False)
                            if _sx["stage_caps"]:
                                pd.DataFrame(_sx["stage_caps"]).rename(columns={
                                    "cong_doan": "Công đoạn", "to": "Tổ SX",
                                    "nv": "Số NV", "cap_m2_thang": "Năng suất (m²/tháng)",
                                    "cap_m2_tuan": "Năng suất (m²/tuần)"
                                }).to_excel(_writer, sheet_name="Năng lực CĐ", index=False)

                    _sx_label = " · Năng lực SX" if st.session_state.get("sx_data") else ""
                    st.download_button(
                        f"💾 Xuất Excel (Dự án input · Tech input · Mặt bằng yêu cầu · Tổng hợp tháng{_sx_label})",
                        data=_full_buf.getvalue(),
                        file_name="ke_hoach_sx.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="t1_full_export",
                    )

# ══════════════════════════════════════════════════════════════════════════════
# TAB NĂNG LỰC SX
# ══════════════════════════════════════════════════════════════════════════════
with tab_nanluc:
    st.header("🔧 Phân Tích Năng Lực Sản Xuất")

    _sx = st.session_state.get("sx_data")

    if _sx is None:
        st.info(
            "⬅️ Vui lòng **tải file SX input.xlsx** ở sidebar bên trái  \n"
            "*(2 sheet: máy móc TB · nhân lực)*"
        )
    else:
        # ── Cảnh báo KHSX vs Năng Lực ─────────────────────────────────────────
        _peak   = st.session_state.get("khsx_peak_demand", 0)
        _peak_m = st.session_state.get("khsx_peak_month", "")
        _over   = st.session_state.get("khsx_overload_months", [])
        _cap    = st.session_state.get("khsx_cap_monthly", cap_monthly)
        _totals = st.session_state.get("khsx_monthly_totals", {})

        if _peak > 0:
            _pct = round(_peak / _cap * 100, 1) if _cap > 0 else 0
            if len(_over) > 0:
                st.error(
                    f"🚨 **CẢNH BÁO: KHSX VƯỢT NĂNG LỰC SX!**  \n"
                    f"📈 Đỉnh tải: **{_peak:,} m²/tháng** ({_peak_m}) = **{_pct}%** năng lực  \n"
                    f"🔴 **{len(_over)} tháng quá tải:** {', '.join(_over)}  \n"
                    f"⚠️ Cần tăng năng lực SX lên ít nhất **{_peak:,} m²/tháng** "
                    f"(hiện tại: {_cap:,} m²/tháng — thiếu **{_peak - _cap:,} m²/tháng**)"
                )
            elif _pct >= 90:
                st.warning(
                    f"⚠️ **Lưu ý: KHSX gần chạm giới hạn năng lực!**  \n"
                    f"📈 Đỉnh tải: **{_peak:,} m²/tháng** ({_peak_m}) = **{_pct}%** năng lực  \n"
                    f"💡 Khuyến nghị tăng năng lực hoặc điều chỉnh KHSX để có buffer an toàn."
                )
            else:
                st.success(
                    f"✅ **Năng lực SX đáp ứng đủ KHSX hiện tại**  \n"
                    f"📈 Đỉnh tải: **{_peak:,} m²/tháng** ({_peak_m}) = **{_pct}%** năng lực  \n"
                    f"🟢 Còn dư: **{_cap - _peak:,} m²/tháng**"
                )
        else:
            st.info("ℹ️ Chưa có dữ liệu KHSX. Vui lòng tải file kế hoạch ở tab **📊 Kế Hoạch Sản Xuất** trước.")

        st.divider()

        # ── KPI row ───────────────────────────────────────────────────────────
        _total_nv  = sum(nl["nv"]  for nl in _sx["nhan_luc"])
        _total_may = sum(mm["sl"]  for mm in _sx["may_moc"])
        _cap_m     = _sx["cap_monthly"] or 0
        _cap_w     = _sx["cap_weekly"]  or 0
        _bn        = _sx["bottleneck"]  or "—"
        _bns_list  = _sx.get("bottlenecks") or ([_bn] if _bn != "—" else [])

        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi-card" style="border-top-color:#6c5ce7">
            <div class="kpi-label">Tổng công nhân SX</div>
            <div class="kpi-value" style="color:#6c5ce7">{_total_nv}</div>
            <div class="kpi-sub">người ({len(_sx["nhan_luc"])} tổ)</div>
          </div>
          <div class="kpi-card" style="border-top-color:#0984e3">
            <div class="kpi-label">Tổng máy móc TB</div>
            <div class="kpi-value" style="color:#0984e3">{_total_may}</div>
            <div class="kpi-sub">máy ({len(_sx["may_moc"])} loại)</div>
          </div>
          <div class="kpi-card" style="border-top-color:#00b894">
            <div class="kpi-label">Năng suất tháng</div>
            <div class="kpi-value green">{_cap_m:,}</div>
            <div class="kpi-sub">m²/tháng (nút thắt)</div>
          </div>
          <div class="kpi-card" style="border-top-color:#fdcb6e">
            <div class="kpi-label">Năng suất tuần</div>
            <div class="kpi-value" style="color:#c49000">{_cap_w:,}</div>
            <div class="kpi-sub">m²/tuần (≈ tháng÷4.33)</div>
          </div>
          <div class="kpi-card red">
            <div class="kpi-label">Nút thắt chuỗi SX</div>
            <div class="kpi-value red" style="font-size:1.0rem">{"<br>".join(_bns_list) if _bns_list else "—"}</div>
            <div class="kpi-sub">công đoạn giới hạn năng suất</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        # ── Bảng năng suất từng công đoạn ────────────────────────────────────
        st.subheader("📊 Năng suất theo công đoạn")
        if _sx["stage_caps"]:
            _df_stage = pd.DataFrame(_sx["stage_caps"])
            _df_stage = _df_stage.rename(columns={
                "cong_doan":    "Công đoạn",
                "to":           "Tổ SX",
                "nv":           "Số NV",
                "cap_m2_thang": "Năng suất (m²/tháng)",
                "cap_m2_tuan":  "Năng suất (m²/tuần)",
            })
            # Format NV as rounded number, hide machine_code
            _df_stage["Số NV"] = _df_stage["Số NV"].apply(
                lambda x: round(x, 1) if x != int(x) else int(x)
            )
            _df_stage = _df_stage.drop(columns=["machine_code"], errors="ignore")

            _df_stage["% so nút thắt"] = (
                _df_stage["Năng suất (m²/tháng)"] / _cap_m * 100
            ).round(1)

            def _style_stage(row):
                if row["Công đoạn"] in _bns_list:
                    return ["background-color:#fff0f0;color:#c0392b;font-weight:bold"] * len(row)
                if row["% so nút thắt"] < 150:
                    return ["background-color:#fff8e1;color:#555555"] * len(row)
                return ["color:#dddddd"] * len(row)

            st.dataframe(
                _df_stage.style.apply(_style_stage, axis=1)
                    .format({
                        "Năng suất (m²/tháng)": "{:,}",
                        "Năng suất (m²/tuần)":  "{:,}",
                        "% so nút thắt":        "{:.1f}%",
                    }),
                use_container_width=True, hide_index=True,
            )

            # ── Biểu đồ bar: năng suất từng công đoạn ────────────────────────
            st.markdown("**📈 Biểu đồ năng suất công đoạn**")
            _colors_bar = [
                "#e17055" if r in _bns_list else "#00b894"
                for r in _df_stage["Công đoạn"]
            ]
            _fig_stage = go.Figure()
            _fig_stage.add_trace(go.Bar(
                x=_df_stage["Công đoạn"],
                y=_df_stage["Năng suất (m²/tháng)"],
                marker_color=_colors_bar,
                text=_df_stage["Năng suất (m²/tháng)"].apply(lambda v: f"{v:,}"),
                textposition="outside",
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    "Năng suất: %{y:,} m²/tháng<extra></extra>"
                ),
            ))
            _fig_stage.add_hline(
                y=_cap_m, line_dash="dash", line_color="#e17055", line_width=2,
                annotation_text=f"  Nút thắt: {_cap_m:,} m²/tháng",
                annotation_font_color="#e17055",
                annotation_position="right",
            )
            _fig_stage.update_layout(
                height=400, showlegend=False,
                xaxis_title="Công đoạn",
                yaxis_title="Năng suất (m²/tháng)",
                margin=dict(t=30, r=200, b=60, l=80),
            )
            st.plotly_chart(_fig_stage, use_container_width=True)

            # ── Chú thích giải thích ──────────────────────────────────────────
            st.info(
                "🔴 **Nút thắt (màu đỏ)** = công đoạn có năng suất thấp nhất → giới hạn toàn chuỗi sản xuất.  \n"
                "Để tăng năng suất nhà máy, cần tăng cường tổ/máy tại công đoạn này trước tiên."
            )
        else:
            st.warning("Không đủ dữ liệu để tính năng suất theo công đoạn.")

        st.divider()

        # ── Hai cột: Nhân lực & Máy móc ──────────────────────────────────────
        _col_nl, _col_mm = st.columns(2)

        with _col_nl:
            st.subheader("👷 Chi tiết nhân lực")
            if _sx["nhan_luc"]:
                _df_nl = pd.DataFrame(_sx["nhan_luc"]).rename(columns={
                    "to": "Tổ", "cong_doan": "Công đoạn",
                    "nv": "Số NV", "nang_suat_m2_gio": "NS m²/người/giờ",
                    "ca": "Ca/ngày", "gio_ca": "Giờ/ca", "hs_pct": "Hiệu suất%",
                    "cap_m2_thang": "Năng suất (m²/tháng)",
                })
                # Handle None values for nang_suat_m2_gio
                _df_nl["NS m²/người/giờ"] = _df_nl["NS m²/người/giờ"].apply(
                    lambda x: f"{x:.2f}" if x is not None else "theo máy"
                )
                st.dataframe(
                    _df_nl.style.format({
                        "Năng suất (m²/tháng)": "{:,}",
                    }),
                    use_container_width=True, hide_index=True,
                )
                # Note: Don't sum capacities - they're sequential stages, not additive
            else:
                st.info("Chưa có dữ liệu nhân lực.")

        with _col_mm:
            st.subheader("⚙️ Chi tiết máy móc thiết bị")
            if _sx["may_moc"]:
                _df_mm = pd.DataFrame(_sx["may_moc"]).rename(columns={
                    "ma": "Mã máy", "ten": "Tên máy",
                    "sl": "SL máy", "sl_nguoi": "NV/máy",
                    "nang_suat_m2_gio": "NS m²/máy/giờ",
                    "ca": "Ca/ngày", "gio_ca": "Giờ/ca",
                    "hs_pct": "Hiệu suất%", "cap_m2_thang": "Năng suất (m²/tháng)",
                })
                st.dataframe(
                    _df_mm.style.format({
                        "NS m²/máy/giờ": "{:.2f}",
                        "Năng suất (m²/tháng)": "{:,}",
                    }),
                    use_container_width=True, hide_index=True,
                )
                # Note: Don't sum capacities - machines work sequentially, not in parallel
            else:
                st.info("Chưa có dữ liệu máy móc.")

        st.divider()

        st.divider()

        # ── What-if Simulation ────────────────────────────────────────────────
        st.subheader("🎛️ What-if: Điều chỉnh máy móc & nhân lực lắp ráp")
        
        _col_hdr, _col_reset = st.columns([6, 1])
        with _col_hdr:
            st.caption("Thay đổi số lượng máy hoặc số NV lắp ráp. Số công nhân vận hành máy tự động tính theo tỷ lệ NV/máy.")
        with _col_reset:
            if st.button("🔄 Reset", key="wif_reset", help="Phục hồi về giá trị ban đầu"):
                # Xóa tất cả what-if keys khỏi session state
                for mm in _sx["may_moc"]:
                    st.session_state.pop(f"wif_m_{mm['ma']}", None)
                for _nl in _sx["nhan_luc"]:
                    if _nl["nang_suat_m2_gio"] is not None:
                        st.session_state.pop(f"wif_nv_{_nl['to']}", None)
                st.rerun()

        # Find assembly team = team có "lắp" trong tên (hoặc team cuối cùng có NS trực tiếp)
        _assembly_team = next(
            (nl for nl in _sx["nhan_luc"] if "lắp" in nl["to"].lower() and nl["nang_suat_m2_gio"] is not None),
            next((nl for nl in reversed(_sx["nhan_luc"]) if nl["nang_suat_m2_gio"] is not None), None)
        )
        
        # Input controls: all machines + assembly workers
        _n_machines = len(_sx["may_moc"])
        _n_inputs = _n_machines + (1 if _assembly_team else 0)
        _wif_cols = st.columns(min(_n_inputs, 4))
        
        _wif_machines = {}  # {machine_name: new_count}
        _wif_assembly_nv = None
        
        _col_idx = 0
        # Show all machines with operator count caption
        for mm in _sx["may_moc"]:
            with _wif_cols[_col_idx % len(_wif_cols)]:
                _new_count = st.number_input(
                    f"Máy – {mm['ten']}",
                    value=mm["sl"], min_value=1, max_value=100, step=1,
                    key=f"wif_m_{mm['ma']}",
                    help=f"{mm['sl_nguoi']} NV/máy · {mm['nang_suat_m2_gio']} m²/giờ/máy"
                )
                _wif_machines[mm["ma"]] = _new_count
                _ops = _new_count * mm["sl_nguoi"]
                st.caption(f"👷 {_ops:.1f} CN vận hành")
                _col_idx += 1

        # Show sliders for ALL direct-productivity teams (not just lắp ráp)
        _wif_direct_nv = {}   # {to_name: new_nv}
        for _nl in _sx["nhan_luc"]:
            if _nl["nang_suat_m2_gio"] is not None:   # direct productivity team
                with _wif_cols[_col_idx % len(_wif_cols)]:
                    _new_nv = st.number_input(
                        f"NV – {_nl['to'].strip()}",
                        value=int(_nl["nv"]), min_value=1, max_value=1000, step=1,
                        key=f"wif_nv_{_nl['to']}",
                        help=f"Năng suất: {_nl['nang_suat_m2_gio']} m²/người/giờ"
                    )
                    _wif_direct_nv[_nl["to"]] = _new_nv
                    _col_idx += 1
        
        # Keep _wif_assembly_nv for backward compat
        if _assembly_team:
            _wif_assembly_nv = _wif_direct_nv.get(_assembly_team["to"], _assembly_team["nv"])
        
        # Recalculate capacity for each stage based on what-if inputs
        _wif_stage_caps = []
        _wif_operator_counts = {}  # Track operator counts for display
        
        # Process each original stage
        for orig_stage in _sx["stage_caps"]:
            _stage_name = orig_stage["cong_doan"]
            _orig_cap = orig_stage["cap_m2_thang"]
            _machine_code = orig_stage.get("machine_code")  # Get machine code if exists
            
            if _machine_code:
                # Machine-based stage: use machine code for precise matching
                matching_machine = next((mm for mm in _sx["may_moc"] if mm["ma"] == _machine_code), None)
                if matching_machine:
                    _new_count = _wif_machines.get(matching_machine["ma"], matching_machine["sl"])
                    _new_cap = round(_new_count * matching_machine["nang_suat_m2_gio"] * 
                                    matching_machine["ca"] * matching_machine["gio_ca"] * 26 * 
                                    (matching_machine["hs_pct"] / 100))
                    _new_operators = _new_count * matching_machine["sl_nguoi"]
                    _wif_operator_counts[_stage_name] = _new_operators
                else:
                    _new_cap = _orig_cap
            else:
                # Direct productivity team: use new NV from slider
                _nl_orig = next((nl for nl in _sx["nhan_luc"]
                                 if nl["cong_doan"] == _stage_name), None)
                if _nl_orig and _nl_orig["nang_suat_m2_gio"] is not None:
                    _new_nv = _wif_direct_nv.get(_nl_orig["to"], int(_nl_orig["nv"]))
                    _new_cap = round(_new_nv * _nl_orig["nang_suat_m2_gio"] *
                                    _nl_orig["ca"] * _nl_orig["gio_ca"] * 26 *
                                    (_nl_orig["hs_pct"] / 100))
                else:
                    _new_cap = _orig_cap
            
            _wif_stage_caps.append({
                "cong_doan": _stage_name,
                "orig_cap": _orig_cap,
                "wif_cap": _new_cap,
            })
        
        # Calculate new bottleneck
        _wif_min = min(s["wif_cap"] for s in _wif_stage_caps)
        _wif_bns = [s["cong_doan"] for s in _wif_stage_caps if s["wif_cap"] == _wif_min]
        _delta_cap = _wif_min - _cap_m

        # Calculate new total NV
        _wif_total_nv = 0
        # 1. Machine-based teams: new_count × NV/machine
        for mm in _sx["may_moc"]:
            _new_count = _wif_machines.get(mm["ma"], mm["sl"])
            _wif_total_nv += _new_count * mm["sl_nguoi"]
        # 2. ALL direct-productivity teams (dập, lắp, etc.)
        for _nl_d in _sx["nhan_luc"]:
            if _nl_d["nang_suat_m2_gio"] is not None:
                _new_nv_d = _wif_direct_nv.get(_nl_d["to"], int(_nl_d["nv"]))
                _wif_total_nv += _new_nv_d
        _delta_nv = _wif_total_nv - _total_nv

        _wc1, _wc2, _wc3, _wc4 = st.columns(4)
        _wc1.metric("Năng suất mới (m²/tháng)", f"{_wif_min:,}",
                    delta=f"{_delta_cap:+,}", delta_color="normal")
        _wc2.metric("Nút thắt mới", " · ".join(_wif_bns))
        _wc3.metric("Tuần (ước tính)", f"{round(_wif_min/4.33):,} m²")
        _wc4.metric("Tổng công nhân SX", f"{_wif_total_nv} người",
                    delta=f"{_delta_nv:+} người", delta_color="normal")

        # ── Cảnh báo What-if vs KHSX ─────────────────────────────────────────
        _wif_peak = st.session_state.get("khsx_peak_demand", 0)
        _wif_over = [ml for ml, v in st.session_state.get("khsx_monthly_totals", {}).items()
                     if v > _wif_min]
        if _wif_peak > 0:
            _wif_pct = round(_wif_peak / _wif_min * 100, 1) if _wif_min > 0 else 0
            if len(_wif_over) > 0:
                st.error(
                    f"🚨 **Năng lực mới {_wif_min:,} m²/tháng — vẫn CHƯA đáp ứng KHSX!**  \n"
                    f"📈 Đỉnh tải KHSX: **{_wif_peak:,} m²/tháng** ({_wif_pct}% năng lực mới)  \n"
                    f"🔴 Còn **{len(_wif_over)} tháng quá tải:** {', '.join(_wif_over)}  \n"
                    f"💡 Cần đạt ít nhất **{_wif_peak:,} m²/tháng** — "
                    f"thiếu **{_wif_peak - _wif_min:,} m²/tháng** nữa"
                )
            elif _wif_pct >= 90:
                st.warning(
                    f"⚠️ **Năng lực mới gần chạm giới hạn KHSX** ({_wif_pct}%)  \n"
                    f"💡 Khuyến nghị tăng thêm để có buffer an toàn."
                )
            else:
                st.success(
                    f"✅ **Năng lực mới {_wif_min:,} m²/tháng ĐÁP ỨNG đủ KHSX!**  \n"
                    f"📈 Đỉnh tải: {_wif_peak:,} m² = {_wif_pct}% — "
                    f"Dư: **{_wif_min - _wif_peak:,} m²/tháng**"
                )
        else:
            st.info("ℹ️ Chưa có dữ liệu KHSX. Tải file kế hoạch ở tab **📊 Kế Hoạch Sản Xuất** để so sánh.")

        # Chart: green (current) vs blue (what-if)
        _stage_names = [s["cong_doan"] for s in _wif_stage_caps]
        _orig_caps = [s["orig_cap"] for s in _wif_stage_caps]
        _new_caps = [s["wif_cap"] for s in _wif_stage_caps]
        
        _wif_fig = go.Figure()
        
        # Add current capacity bars (GREEN FILLED)
        _wif_fig.add_trace(go.Bar(
            x=_stage_names, 
            y=_orig_caps, 
            name="Năng suất hiện tại",
            marker=dict(color="#00b894"),  # Green filled
            text=[f"{v:,}" for v in _orig_caps],
            textposition="outside",
            textfont=dict(size=9),
        ))
        
        # Add what-if capacity bars (BLUE FILLED)
        _wif_fig.add_trace(go.Bar(
            x=_stage_names, 
            y=_new_caps, 
            name="Năng suất sau khi điều chỉnh",
            marker=dict(color="#74b9ff"),  # Blue filled
            text=[f"{v:,}" for v in _new_caps],
            textposition="outside",
            textfont=dict(size=9),
        ))
        
        # Add bottleneck line
        _wif_fig.add_hline(
            y=_wif_min, 
            line_dash="dash", 
            line_color="#e74c3c", 
            line_width=2,
            annotation_text=f"Nút thắt mới: {_wif_min:,} m²",
            annotation_position="right",
            annotation_font=dict(color="#e74c3c", size=11),
        )
        
        _wif_fig.update_layout(
            barmode="group",  # Group bars side by side
            height=400,
            showlegend=True,
            legend=dict(
                title="",
                orientation="h", 
                yanchor="bottom", 
                y=1.01, 
                xanchor="left", 
                x=0,
                font=dict(size=12, color="#333"),  # Explicit text color
                bgcolor="rgba(255, 255, 255, 0.9)",
                bordercolor="#999",
                borderwidth=1,
            ),
            xaxis=dict(
                title="Công đoạn",
                tickangle=-45,
                tickfont=dict(size=10),
            ),
            yaxis=dict(
                title="Năng suất (m²/tháng)",
                tickfont=dict(size=10),
            ),
            margin=dict(t=60, b=100, l=60, r=20),
            plot_bgcolor="white", 
            paper_bgcolor="white", 
            font=dict(color="#333"),
        )
        _wif_fig.update_xaxes(showgrid=False)
        _wif_fig.update_yaxes(showgrid=True, gridcolor="#e0e0e0", gridwidth=1)
        st.plotly_chart(_wif_fig, use_container_width=True)

        # ── Phân tích tải trọng theo dự án (nếu đã có file dự án) ─────────────
        _rows_proj = st.session_state.get("t1_rows", [])
        if _rows_proj and _sx["stage_caps"]:
            st.subheader("📋 Phân tích tải trọng công đoạn theo tháng hiện tại")
            # Tháng hiện tại
            _today = datetime(datetime.today().year, datetime.today().month, 1)
            _m2_thang = sum(
                r["m2"] / max(1, (
                    (datetime(r["end"].year, r["end"].month, 1) -
                     datetime(r["start"].year, r["start"].month, 1)).days // 30 + 1
                ))
                for r in _rows_proj
                if (
                    datetime(r["start"].year, r["start"].month, 1) <= _today
                    <= datetime(r["end"].year, r["end"].month, 1)
                )
            ) * (show_weighted and 1 or 1)

            if _m2_thang > 0:
                st.caption(
                    f"Khối lượng SX ước tính tháng hiện tại: **{_m2_thang:,.0f} m²**  "
                    f"(trung bình từ {len(_rows_proj)} dự án đang chạy)"
                )
                _load_rows = []
                for sc in _sx["stage_caps"]:
                    _pct = _m2_thang / sc["cap_m2_thang"] * 100 if sc["cap_m2_thang"] > 0 else 0
                    _load_rows.append({
                        "Công đoạn":             sc["cong_doan"],
                        "Tổ SX":                 sc["to"],
                        "Năng suất (m²/tháng)":  sc["cap_m2_thang"],
                        "Nhu cầu tháng này (m²)": round(_m2_thang),
                        "% Tải trọng":            round(_pct, 1),
                        "Trạng thái":             (
                            "🔴 QUÁ TẢI"       if _pct >= 100 else
                            "🟠 Căng tiến độ"  if _pct >= 85  else
                            "🟡 Cần chú ý"     if _pct >= 70  else
                            "🟢 Bình thường"
                        ),
                    })
                _df_load = pd.DataFrame(_load_rows)
                st.dataframe(
                    _df_load.style.format({
                        "Năng suất (m²/tháng)":   "{:,}",
                        "Nhu cầu tháng này (m²)": "{:,}",
                        "% Tải trọng":            "{:.1f}%",
                    }),
                    use_container_width=True, hide_index=True,
                )
            else:
                st.info("Không có dự án nào đang chạy trong tháng hiện tại.")
        elif not _rows_proj:
            st.caption("💡 Tải thêm file dự án ở tab **📊 Kế Hoạch Sản Xuất** để xem phân tích tải trọng theo công đoạn.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB NHU CẦU VS NĂNG LỰC
# ══════════════════════════════════════════════════════════════════════════════
with tab_nvcnl:
    st.header("📈 Nhu Cầu vs Năng Lực Sản Xuất theo tháng")
    rows_nvcnl = st.session_state.get("t1_rows", [])
    if not rows_nvcnl:
        st.info("⬅️ Vui lòng vào tab **📊 Kế Hoạch Sản Xuất** và tải file Excel trước.")
    else:
        monthly_demand: dict = {}
        for r in rows_nvcnl:
            start = datetime(r["start"].year, r["start"].month, 1)
            end   = datetime(r["end"].year,   r["end"].month,   1)
            factor = (r["prob"] / 100) if show_weighted else 1.0
            ms = months_range(start, end)
            if not ms:
                continue
            m2pm = r["m2"] * factor / len(ms)
            for m in ms:
                monthly_demand[m] = monthly_demand.get(m, 0) + m2pm

        if not monthly_demand:
            st.warning("Không có dữ liệu nhu cầu.")
        else:
            today_m = datetime(datetime.today().year, datetime.today().month, 1)
            cutoff  = today_m + pd.DateOffset(months=horizon)
            months_sorted = sorted(k for k in monthly_demand if k <= cutoff)
            labels  = [m.strftime("%m/%Y") for m in months_sorted]
            demands = [monthly_demand[m] for m in months_sorted]

            fig_nvcnl = go.Figure()
            bar_colors = [
                "#e17055" if d > cap_monthly else
                "#fd9644" if d > cap_monthly * 0.85 else
                "#00b894"
                for d in demands
            ]
            fig_nvcnl.add_trace(go.Bar(
                x=labels, y=demands, name="Nhu cầu (m²/tháng)",
                marker_color=bar_colors,
                text=[f"{d:,.0f}" for d in demands], textposition="outside",
                hovertemplate="<b>%{x}</b><br>Nhu cầu: %{y:,.0f} m²<extra></extra>",
            ))
            fig_nvcnl.add_trace(go.Scatter(
                x=labels, y=[cap_monthly] * len(months_sorted),
                name=f"Năng lực tối đa ({cap_monthly:,} m²)",
                mode="lines", line=dict(color="#e17055", dash="dash", width=2),
            ))
            fig_nvcnl.update_layout(
                height=480,
                legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
                xaxis_title="Tháng", yaxis_title="Khối lượng (m²)",
                xaxis=dict(tickangle=45), margin=dict(t=50, b=80),
                plot_bgcolor="white", paper_bgcolor="white", font=dict(color="#333"),
            )
            st.plotly_chart(fig_nvcnl, use_container_width=True)

            st.subheader("📊 Thặng dư / Thiếu hụt năng lực theo tháng")
            nvcnl_rows = []
            for lbl, d in zip(labels, demands):
                diff = cap_monthly - d
                pct  = d / cap_monthly * 100
                nvcnl_rows.append({
                    "Tháng": lbl,
                    "Nhu cầu (m²)": round(d),
                    "Năng lực (m²)": cap_monthly,
                    "Thặng dư / Thiếu hụt (m²)": round(diff),
                    "% Tải trọng": round(pct, 1),
                    "Trạng thái": (
                        "🔴 QUÁ TẢI"      if pct >= 100 else
                        "🟠 Căng tiến độ" if pct >= 85  else
                        "🟡 Cần chú ý"    if pct >= 70  else
                        "🟢 Bình thường"
                    ),
                })
            df_nvcnl = pd.DataFrame(nvcnl_rows)
            st.dataframe(
                df_nvcnl.style.format({
                    "Nhu cầu (m²)":               "{:,}",
                    "Năng lực (m²)":              "{:,}",
                    "Thặng dư / Thiếu hụt (m²)":  "{:+,}",
                    "% Tải trọng":                "{:.1f}%",
                }),
                use_container_width=True, hide_index=True,
            )

            n_over = sum(1 for d in demands if d > cap_monthly)
            c1, c2, c3 = st.columns(3)
            c1.metric("Tháng quá tải",    f"{n_over} tháng",
                      delta=f"-{n_over}" if n_over else None, delta_color="inverse")
            c2.metric("Tháng bình thường", f"{len(demands) - n_over} tháng")
            c3.metric("Năng lực tối đa",  f"{cap_monthly:,} m²/tháng")

# ══════════════════════════════════════════════════════════════════════════════
# TAB TIẾN ĐỘ THỰC TẾ
# ══════════════════════════════════════════════════════════════════════════════
with tab_thucte:
    st.header("📋 Tiến Độ Thực Tế vs Kế Hoạch")
    rows_tt = st.session_state.get("t1_rows", [])
    if not rows_tt:
        st.info("⬅️ Vui lòng vào tab **📊 Kế Hoạch Sản Xuất** và tải file Excel trước.")
    else:
        monthly_plan: dict = {}
        for r in rows_tt:
            start = datetime(r["start"].year, r["start"].month, 1)
            end   = datetime(r["end"].year,   r["end"].month,   1)
            factor = (r["prob"] / 100) if show_weighted else 1.0
            ms = months_range(start, end)
            if not ms:
                continue
            m2pm = r["m2"] * factor / len(ms)
            for m in ms:
                monthly_plan[m] = monthly_plan.get(m, 0) + m2pm

        today_m  = datetime(datetime.today().year, datetime.today().month, 1)
        cutoff_t = today_m + pd.DateOffset(months=12)
        months_tt = sorted(k for k in monthly_plan if k <= cutoff_t)

        if not months_tt:
            st.warning("Không có dữ liệu kế hoạch.")
        else:
            labels_tt = [m.strftime("%m/%Y") for m in months_tt]
            plans_tt  = [round(monthly_plan[m]) for m in months_tt]
            actuals_saved = st.session_state.get("tt_actuals", {})

            input_rows = [
                {"Tháng": lbl, "Kế hoạch (m²)": p, "Thực tế (m²)": actuals_saved.get(lbl, 0)}
                for lbl, p in zip(labels_tt, plans_tt)
            ]
            st.subheader("✏️ Nhập sản lượng thực tế")
            edited = st.data_editor(
                pd.DataFrame(input_rows),
                column_config={
                    "Tháng":         st.column_config.TextColumn("Tháng", disabled=True),
                    "Kế hoạch (m²)": st.column_config.NumberColumn("KH (m²)", disabled=True, format="%d"),
                    "Thực tế (m²)":  st.column_config.NumberColumn("Thực tế (m²)", min_value=0, step=100, format="%d"),
                },
                use_container_width=True, hide_index=True, key="tt_editor",
            )
            st.session_state["tt_actuals"] = {row["Tháng"]: row["Thực tế (m²)"] for _, row in edited.iterrows()}

            actuals_list = [st.session_state["tt_actuals"].get(lbl, 0) for lbl in labels_tt]
            cum_plan, cum_act = [], []
            cp = ca_val = 0
            for p, a in zip(plans_tt, actuals_list):
                cp += p; ca_val += a
                cum_plan.append(cp); cum_act.append(ca_val)

            fig_tt = go.Figure()
            fig_tt.add_trace(go.Bar(x=labels_tt, y=plans_tt, name="Kế hoạch (m²/tháng)",
                                    marker_color="#a29bfe", opacity=0.7))
            fig_tt.add_trace(go.Bar(x=labels_tt, y=actuals_list, name="Thực tế (m²/tháng)",
                                    marker_color="#00b894", opacity=0.85))
            fig_tt.add_trace(go.Scatter(x=labels_tt, y=cum_plan, name="Lũy kế kế hoạch",
                                        yaxis="y2", mode="lines+markers",
                                        line=dict(color="#6c5ce7", dash="dot", width=2)))
            fig_tt.add_trace(go.Scatter(x=labels_tt, y=cum_act, name="Lũy kế thực tế",
                                        yaxis="y2", mode="lines+markers",
                                        line=dict(color="#e17055", width=2), marker=dict(size=7)))
            fig_tt.update_layout(
                barmode="group", height=480,
                yaxis=dict(title="m²/tháng"),
                yaxis2=dict(title="Lũy kế (m²)", overlaying="y", side="right", showgrid=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
                xaxis=dict(tickangle=45), margin=dict(t=50, b=80),
                plot_bgcolor="white", paper_bgcolor="white", font=dict(color="#333"),
            )
            st.plotly_chart(fig_tt, use_container_width=True)

            total_plan_tt   = sum(plans_tt)
            total_actual_tt = sum(actuals_list)
            done_months     = sum(1 for a in actuals_list if a > 0)
            pct_done        = total_actual_tt / total_plan_tt * 100 if total_plan_tt else 0
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tổng KH (m²)",      f"{total_plan_tt:,}")
            c2.metric("Tổng thực tế (m²)",  f"{total_actual_tt:,}")
            c3.metric("% Hoàn thành",       f"{pct_done:.1f}%")
            c4.metric("Tháng đã nhập",      f"{done_months}/{len(labels_tt)}")
