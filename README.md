# 🏭 Kế Hoạch Sản Xuất Nhà Máy QDP

Ứng dụng quản lý và tối ưu kế hoạch sản xuất nhà máy QDP, xây dựng bằng **Streamlit** + **Plotly** + **PuLP**.

## ✨ Tính năng

- 📊 **Bảng tổng hợp tải trọng** theo tháng (stacked bar chart)
- 📅 **Biểu đồ Gantt** tiến độ từng dự án (màu theo xác suất)
- ✏️ **Chỉnh sửa trực tiếp** danh sách dự án trong bảng
- 🔍 **What-If Simulator** kiểm tra dự án mới có khả thi không
- 🔧 **Tối ưu LP** phân bổ tải trọng tối ưu bằng Linear Programming
- 💾 **Xuất Excel** danh sách dự án đã chỉnh sửa

## 🚀 Chạy local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📂 Dữ liệu đầu vào

Upload file Excel với cấu trúc:

| STT | Tên dự án | Khối lượng (m²) | Bắt đầu | Kết thúc | Ghi chú |
|-----|-----------|-----------------|---------|---------|---------|
| 1   | Dự án A   | 5000            | 01/2025 | 06/2025 | Đã ký hợp đồng |
| 2   | Dự án B   | 8000            | 03/2025 | 09/2025 | 90% |

## 🎨 Màu sắc Gantt

- 🔵 **Xanh dương** — Đã ký hợp đồng (100%)
- 🟢 **Xanh lá** — Khả năng cao (90%)
- 🟠 **Cam** — Đang xét (50%)
- ⬜ **Xám** — Xem xét (<50%)

## 🛠️ Tech Stack

- [Streamlit](https://streamlit.io) — UI framework
- [Plotly](https://plotly.com) — Biểu đồ tương tác
- [PuLP](https://coin-or.github.io/pulp/) — Tối ưu tuyến tính (LP)
- [Pandas](https://pandas.pydata.org) — Xử lý dữ liệu
