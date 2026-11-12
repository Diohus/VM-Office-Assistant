# Biểu mẫu giả cho khung chat (lát 2)

[MẪU GIẢ]. Sinh bằng `fixtures/eval/generate_chat_fixtures.py`, seed cố định. Không có đáp án ở đây.

| File | Dùng cho |
|---|---|
| `GIA-QC-01_chat.xlsx` | Có số giả 30 ngày. Tên giả ở cột `Người chạy`, ô `Phụ trách`, dưới ô ký. Ô `I12` có lệnh giả ngắn (Y10), ô `I20` có lệnh giả dài |
| `GIA-QC-01_chat_lon.xlsx` | Như trên, mở sẵn sheet `NhatKy_Nam` 3.660 ô (mục 7, trên 3.000 ô) |

File được lưu bằng openpyxl: định dạng có điều kiện, kiểm dữ liệu và biểu đồ của file gốc bị bỏ; ô công thức không có kết quả tính sẵn.
