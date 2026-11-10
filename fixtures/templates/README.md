# Fixture biểu mẫu [MẪU GIẢ]

Toàn bộ file trong thư mục này là dữ liệu giả do đội tự tạo để thử phần mềm, không dùng cho chuyên môn. Mọi tên người, khoa, mã lô đều là giả. Mô tả theo `docs/product/qc_form_spec.md`.

Sinh lại bằng `spikes/editor/gen_fixtures.py` và `run_x19.py`, `run_lo.py`. Các file được Excel thật lưu lại để XML giống file thực tế. Excel ghi đường dẫn máy vào `xl/workbook.xml`, nên sau khi sinh phải chạy `tools/excel_check/clean_fixture_paths.py`; test `test_commit_scan.py` sẽ trượt nếu quên. File vượt giới hạn sinh bằng `tools/fixtures/make_limit_files.py`, nằm ở `fixtures/generated/`, ngoài git.

| File | Dùng để |
|---|---|
| `GIA-QC-01_mau.xlsx` | Biểu mẫu QC trống theo spec |
| `GIA-QC-01_da_dien.xlsx` | Cùng biểu mẫu đã điền số giả, để so công thức |
| `GIA-NGUON_2026-09.xlsx` | File nguồn chuẩn (X01), số giả, seed cố định |
| `GIA-KHO-02_tieu_hao.xlsx` | Biểu mẫu khó: Table, ảnh, shape, tên vùng, `INDIRECT`, `OFFSET`, tham chiếu ba chiều, công thức mảng, thanh dữ liệu, 71 hàm |
| `GIA-QC-01_da_dien_doi_cu.xls` | Thử chuyển `.xls` |
| `GIA-QT-03_quy_trinh_doi_cu.doc` | Thử chuyển `.doc` |
| `GIA-X18_ham_mang.xlsx` | X18: `WEBSERVICE`, liên kết ngoài. Không dùng trên máy có mạng ngoài kiểm soát |
| `GIA-X18b_lien_ket_ngoai.xlsx` | X18: liên kết ngoài thật, có phần `xl/externalLinks`, `I9` chỉ lộ liên kết qua tên vùng. Sinh bằng `tools/excel_check/make_external_link.py` |
| `GIA-X19_co_mat_khau.xlsx` | X19: file có mật khẩu giả `matkhau-gia` |
| `GIA-X19_xls_mat_khau.xls` | X19: file đời cũ có mật khẩu giả `matkhau-gia`. Sinh bằng `tools/excel_check/make_xls_password.py` |
| `GIA-X19_khong_co_vba.xlsm` | X19: đuôi `.xlsm`, **không có VBA thật** |
| `GIA-X19_hong.xlsx` | X19: file bị cắt ngang, không mở được |
| `GIA-X19_sheet_khoa.xlsx` | Sheet có khóa, mật khẩu giả `khoa-gia` |

`GIA-KHO-02_kho.xlsx` và `~$GIA-KHO-02_kho.xlsx` là bản lỗi cũ, bị một tiến trình Excel ẩn giữ khóa. Xóa khi tiến trình đó đã tắt. Đừng dùng bản này: tên sheet `Quy1` trùng dạng địa chỉ ô.
