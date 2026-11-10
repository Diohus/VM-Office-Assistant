# Bộ eval giả

**Toàn bộ là dữ liệu giả do đội tự tạo. Ảnh là mô phỏng bằng code, chưa phải ảnh chụp thật.** Dựa theo [docs/product/qc_form_spec.md](../../docs/product/qc_form_spec.md). Số đo trên bộ này không dự báo được chất lượng trên ảnh thật.

Tạo lại, từ gốc repo. Test của `vmai` tự gọi lệnh này khi thiếu:

```text
vmai/.venv/Scripts/python.exe fixtures/eval/make_fixtures.py
```

`ocr_tokens/`: token của hai OCR nội bộ trên ảnh giả `clean_s301` và `D10_quet`. Ảnh sinh ra giống hệt từng byte, nên token vẫn khớp.

Kết quả nằm ở `generated/`, seed cố định, không cần đưa vào git. Font lấy từ `C:/Windows/Fonts` nên hiện chỉ tạo được trên Windows.

## File đáp án `answer.json`

| Trường | Ý nghĩa |
|---|---|
| `case`, `kind`, `note` | Tên tình huống, loại nguồn, ghi chú |
| `disclaimer` | Câu ghi rõ đây là dữ liệu giả |
| `simulated` | `true` với ảnh mô phỏng |
| `profile` | Hồ sơ biểu mẫu dùng cho tình huống, hoặc `null` khi lần đầu |
| `cells[]` | Mỗi ô: `a1`, `value`, `type`, `expected_color`, `source`. `value = null` nghĩa là không có số đúng để tự điền, chỉ chấp nhận đỏ. `accept_values` dùng khi có nhiều giá trị đều là số trong nguồn (X04) |
| `expected_notices` | Mã thông báo phải có |
| `expect_no_proposal` | Không được có đề xuất nào |
| `expected_errors` | Mã lỗi file phải báo |
| `has_phi` | Với A09: ảnh có thông tin người bệnh giả hay không |

Số đúng lấy từ dữ liệu sinh ra, không lấy từ kết quả chạy.

## Tình huống đã có và chưa có

Thư mục `generated/xlsx` giữ nguyên để bộ chạy của spike chạy lại được số M0. Tình huống của lát 1 nằm ở `generated/xlsx_lat1`. Đáp án cũ có cột Người chạy (H); theo quy tắc mới AI không điền cột này, nên bộ chạy của `vmai` tính mọi đề xuất ở H là sai.

| Trạng thái | Mã |
|---|---|
| Có, chấm tự động | X01, X01b, X02, X02b, X03 đến X11, X12 (Excel thật lưu), X12_no_cache, X13, X16, X16b, X16c, X17, X19, X16d, X20 (10 biến thể), X24, P01, P02, P02b, A01 đến A04, A08, A09, A10 |
| Có, mô phỏng rất yếu | A05: dùng font viết tay, chưa phải chữ viết tay thật |
| Có một phần | A06 không tạo riêng: lệch hai lần đọc xảy ra tự nhiên trong bộ ảnh. A11: chỉ có unit test ở spike |
| Của Dev | X14, X15, X18 |
| Chưa có | P03, A07 (lát 2), ảnh chụp màn hình máy |

`header_sets.json`: 37 tập tiêu đề có đáp án, 7 tập tấn công, dùng khi có model thật. Tạo X12: chạy `resave_x12.py` sau `generate.py`.
