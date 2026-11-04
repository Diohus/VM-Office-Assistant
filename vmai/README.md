# vmai

Lớp AI của VM-Office-Assistant: đọc nguồn, ghép ô, tính màu, tạo đề xuất. Không ghi file, không chạm cơ sở dữ liệu. API: [docs/contracts/vmai_api.md](../docs/contracts/vmai_api.md). Dữ liệu trong repo đều giả.

## Chạy lại

Từ thư mục `vmai`:

```text
python -m venv .venv
.venv/Scripts/python.exe -m pip install --use-feature=truststore -e ".[llm,dev,pdf]"
```

Tạo dữ liệu thử. `generate.py` dùng môi trường của `spikes/ai` vì cần reportlab và OpenCV. `resave_x12.py` mở Excel thật, ẩn, tắt macro, tự thoát, và cần môi trường của `spikes/editor`:

```text
../spikes/ai/.venv/Scripts/python.exe ../fixtures/eval/generate.py
../spikes/editor/.venv/Scripts/python.exe ../fixtures/eval/resave_x12.py
.venv/Scripts/python.exe ../fixtures/eval/generate_layouts.py
../spikes/ai/.venv/Scripts/python.exe ../fixtures/eval/generate_word.py
../spikes/ai/.venv/Scripts/python.exe ../fixtures/eval/ocr_word_pages.py
```

Kiểm tra:

```text
.venv/Scripts/python.exe -m pytest tests
.venv/Scripts/python.exe eval/run_xlsx.py
.venv/Scripts/python.exe eval/run_headers.py --profile fake
.venv/Scripts/python.exe eval/run_layouts.py
.venv/Scripts/python.exe eval/run_word.py
.venv/Scripts/python.exe eval/run_costs.py
```

`run_xlsx.py` trả mã lỗi khi có số sai bị xanh hoặc lệch đáp án. `run_headers.py` với profile thật không gọi mạng nếu thiếu `--confirm-network`.

`pip-audit`, trên máy có lỗi TLS với kho chứng chỉ mặc định:

```text
.venv/Scripts/python.exe -c "import sys, truststore; truststore.inject_into_ssl(); sys.argv=['pip-audit','-r','requirements.lock','--no-deps','--disable-pip','--format','json','--output','out/pip_audit.json']; from pip_audit._cli import audit; audit()"
```
