# Biểu mẫu giả cho khung chat (lát 2)

[MẪU GIẢ]. Sinh bằng `fixtures/eval/generate_chat_fixtures.py`, seed cố định. Không có đáp án ở đây.

| File | Dùng cho |
|---|---|
| `GIA-QC-01_chat.xlsx` | Có số giả 30 ngày. Tên giả ở cột `Người chạy`, ô `Phụ trách`, dưới ô ký. Ô `I12` có lệnh giả ngắn (Y10), ô `I20` có lệnh giả dài |
| `GIA-QC-01_chat_lon.xlsx` | Như trên, mở sẵn sheet `NhatKy_Nam` 3.660 ô (mục 7, trên 3.000 ô) |
| `GIA-QC-01_chat_benh_nhan.xlsx` | Dòng 2.9: mở sẵn sheet `KQ_BenhNhan` (A1:F23) có cột `Mã BN`, `Họ tên người bệnh`, `Ngày sinh` giả, ghi chú có tên người bệnh giả. Thông tin file để trống |
| `GIA-QC-01_chat_sheet_lon.xlsx` | Dòng 2.13: như `GIA-QC-01_chat.xlsx`, thêm sheet `DuLieu_Nam` A1:J301 (3.010 ô), ô `I44` có công thức trỏ tới sheet đó. Sheet chính cộng sheet phụ vượt 3.000 ô. Thông tin file để trống |

File được lưu bằng openpyxl: định dạng có điều kiện, kiểm dữ liệu và biểu đồ của file gốc bị bỏ; ô công thức không có kết quả tính sẵn.
