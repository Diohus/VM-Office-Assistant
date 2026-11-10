# vmai

Lớp AI của VM-Office-Assistant: đọc nguồn, ghép ô, tính màu, tạo đề xuất. Không ghi file, không chạm cơ sở dữ liệu. API: [docs/contracts/vmai_api.md](../docs/contracts/vmai_api.md). Dữ liệu trong repo đều giả.

## Chạy lại

Từ thư mục `vmai`:

```text
python -m venv .venv
.venv/Scripts/python.exe -m pip install --use-feature=truststore -e ".[llm,dev,pdf,ocr]"
```

Dữ liệu thử giả: `pytest` tự sinh khi thiếu, qua `fixtures/eval/make_fixtures.py` (seed cố định, cần font Windows). Sinh hỏng thì cả lần chạy dừng với mã 1, test bẫy không bị bỏ qua lặng lẽ. Chạy tay:

```text
.venv/Scripts/python.exe ../fixtures/eval/make_fixtures.py
```

Token OCR của ảnh giả nằm sẵn ở `fixtures/eval/ocr_tokens/`, nên test không cần cài OCR. Mã lý do lấy từ `tests/data/doc_codes.json`, bản chụp của `docs/ai`; sửa tài liệu thì chạy `tests/doc_snapshot.py`. X12 (Excel thật lưu) vẫn cần `../spikes/editor/.venv/Scripts/python.exe ../fixtures/eval/resave_x12.py`, test không dùng.

Kiểm tra:

```text
.venv/Scripts/python.exe -m pytest tests
.venv/Scripts/python.exe eval/run_xlsx.py
.venv/Scripts/python.exe eval/run_headers.py --profile fake
.venv/Scripts/python.exe eval/run_layouts.py
.venv/Scripts/python.exe eval/run_word.py
.venv/Scripts/python.exe eval/run_costs.py
```

Ca giới hạn (lô lớn, tiêu đề lạ, nhiều nguồn), tự tạo dữ liệu giả trong `out/limits/`, khoảng 12 phút, `--quick` khoảng 3 phút. Kết quả ở [docs/huong_dan/kich_ban_test_ai.md](../docs/huong_dan/kich_ban_test_ai.md):

```text
.venv/Scripts/python.exe -W ignore eval/run_limits.py --quick
```

Đo "nghi nhầm file cũ" (chỉ đo, chưa là quy tắc): `.venv/Scripts/python.exe eval/run_stale_share.py`. Rủi ro Owner chấp nhận (Q131, trước là Q118) in riêng ở `run_xlsx.py` và ở cuối `pytest`; test `tests/test_known_leaks.py` là xfail strict, số 180/180 được ghim.

Các script eval tự in UTF-8 (`eval/_console.py`), không cần `PYTHONIOENCODING`.

Mã thoát 1: `run_xlsx.py` khi có số sai bị xanh hoặc lệch đáp án; `run_layouts.py` khi xanh trước xác nhận, sai bị xanh, đoán sai không báo, lộ giá trị cấm; `run_limits.py`, `run_pdf.py`, `run_images_q90.py score` khi có sai bị xanh; `run_headers.py` khi tiêu đề về người có trong yêu cầu gửi đi (`NEVER_SEND`) hoặc bộ tấn công lọt. `run_word.py`, `run_costs.py` chỉ đo, luôn trả 0. `run_headers.py` với profile thật không gọi mạng nếu thiếu `--confirm-network`.

Xuất lô thật làm mẫu cho phần 3 (ca ảnh cần cache OCR trong `out/`, thiếu thì bỏ qua):

```text
.venv/Scripts/python.exe eval/export_batches.py
```

Đo ảnh (cần image `vmai-ocr-do`, xem `docs/ai/image_report.md`). Không gọi dịch vụ ngoài:

```text
.venv/Scripts/python.exe eval/run_images_q90.py read
.venv/Scripts/python.exe eval/run_images_q90.py score
```

Bộ ảnh độc lập ngoài repo (Q96), người giữ bộ chạy. Bộ có dạng `<bộ>/images/<ca>/page.jpg`, `template.xlsx`, `answer.json` và `<bộ>/images/manifest.json`; trong `answer.json` nên có `layout`, `variant`. `--out` phải nằm ngoài repo, cache OCR và `eval_external.json` chỉ ghi vào đó, vì chứa chữ đọc được từ ảnh và tên ca. Màn hình chỉ in số tổng theo bố cục, loại ảnh, mã lý do. Thoát mã 1 nếu có số sai bị xanh.

```text
.venv/Scripts/python.exe eval/run_images_q90.py read --set <bộ> --out <thư mục kết quả>
.venv/Scripts/python.exe eval/run_images_q90.py score --set <bộ> --out <thư mục kết quả>
```

Lượt có model (Q85), chỉ dữ liệu giả, chỉ gửi tiêu đề, theo trần chi phí. Bỏ `--confirm-network` thì chỉ đếm những gì sẽ gửi. Dòng `Model:` ghi số lần gọi, token và chi phí theo bảng giá:

```text
.venv/Scripts/python.exe eval/run_images_q90.py score --set <bộ> --out <thư mục kết quả> --profile openai-luna --confirm-network
```

PDF có lớp chữ (Q98), bộ dạng `<bộ>/pdf/<ca>/source.pdf`, `template.xlsx`, `answer.json`:

```text
.venv/Scripts/python.exe eval/run_pdf.py --set <bộ> --out <thư mục kết quả>
```

Ghi lại số lần đã chạy trên mỗi bộ: đo rồi sửa theo thì bộ đó thành bộ đã chỉnh (Q96).

`pip-audit`, trên máy có lỗi TLS với kho chứng chỉ mặc định:

```text
.venv/Scripts/python.exe -c "import sys, truststore; truststore.inject_into_ssl(); sys.argv=['pip-audit','-r','requirements.lock','--no-deps','--disable-pip','--format','json','--output','out/pip_audit.json']; from pip_audit._cli import audit; audit()"
```
