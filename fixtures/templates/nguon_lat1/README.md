# Nguồn giả cho nghiệm thu lát 1 [MẪU GIẢ]

Toàn bộ là dữ liệu giả do đội tự tạo. Mọi số, tên người ("... Giả ..."), mã người bệnh ("BN-GIA-...") đều giả. Dùng với biểu mẫu `../GIA-QC-01_mau.xlsx`, mô tả `GIA-QC-01`. Theo `docs/product/acceptance_lat1.md` mục 0 và `docs/product/qc_form_spec.md`.

Sinh lại, từ gốc repo, seed cố định: `vmai/.venv/Scripts/python.exe fixtures/eval/generate_lat1_sources.py`. Gốc của mọi file là `../GIA-NGUON_2026-09.xlsx` (X01 chuẩn, tháng 9/2026); mỗi file chỉ đổi đúng điều ghi dưới đây.

README này **không có đáp án số**. PO/BA tự tính đáp án bằng script riêng.

| File | Tình huống | App phải làm gì |
|---|---|---|
| `../GIA-NGUON_2026-09.xlsx` | X01: nguồn chuẩn, đủ ngày, đúng tên cột | Điền mọi ô số của bảng. Cột Người chạy để trống, báo cho nhân viên tự viết |
| `X02_doi_ten_cot.xlsx` | Đổi tên cột: `Kết quả` thành `KQ`, `Mã XN` thành `Tên XN`; giá trị `GLU` ghi là `Đường máu (giả)`, tên **chưa khai** trong mô tả | Ghép được cột theo tên khác. Không đoán xét nghiệm có tên lạ, không gửi giá trị ô ra ngoài. Báo rõ và không là ngõ cụt (dòng 2.5). Không xanh |
| `X02a_ten_da_khai.xlsx` | Như X02, nhưng giá trị `GLU` ghi là `Đường huyết`, tên **đã khai** là tên khác của GLU | Ghép được cả cột và tên xét nghiệm theo bảng đã khai. Không xanh |
| `X03_thieu_3_ngay.xlsx` | Thiếu toàn bộ dòng của 3 ngày | Để trống các ô của 3 ngày đó, báo rõ thiếu ngày nào, không tự suy số |
| `X04_mot_ngay_hai_lan_chay.xlsx` | Ngày 10/09 có 2 lần chạy GLU L1, khác số | Không tự chọn. Một đề xuất đỏ cho ô đó, hiện cả hai số để người dùng chọn một hoặc tự gõ (D10) |
| `X07_so_dang_chu_dau_phay.xlsx` | Các ô GLU L1 lưu dạng chữ, dấu phẩy thập phân (kiểu `5,47`) | Chuyển thành số đúng, giữ chữ gốc trong nguồn, báo đã đổi chữ thành số |
| `X10_so_vo_ly.xlsx` | 5 ô GLU L1 có số vô lý so với đích: `50`, `55`, `527`, `0,527`, `52,7` | Giữ nguyên số, không tự sửa, cảnh báo. Các ô đó đỏ |
| `X13_sai_thang_10.xlsx` | Mọi dòng là tháng 10/2026, biểu mẫu là tháng 9 | Báo lệch tháng, không điền ô nào |
| `nguoi_chay_ten_gia.xlsx` | Cột `Người chạy` ghi tên người giả | Không điền cột Người chạy. Tên người không có trong đề xuất, lô lưu hay log |
| `X17_cot_nguoi_benh_gia.xlsx` | Thêm cột `Họ tên người bệnh`, `Mã BN` với giá trị giả | Không dùng hai cột này, báo có cột giống thông tin người bệnh. Giá trị không có trong đề xuất, lô lưu hay log; không gửi ra ngoài. Phần số vẫn điền |

Lát 1 (Q133): không có hồ sơ, không model, mọi ô tối đa vàng. Màu "xanh" ghi trong `qc_form_spec.md` cho một số ca (ví dụ X07) không áp ở lát 1.
