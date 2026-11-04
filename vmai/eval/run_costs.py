"""Ước tính lượng dữ liệu gửi lên model theo loại việc, chưa gọi model nào.

Đây là ƯỚC TÍNH từ số ký tự, không phải số đo. Số tiền thật chỉ có sau khi có key và chốt model.
"""
from __future__ import annotations

import json
from pathlib import Path

from vmai.document import read_pdf_lines
from vmai.llm.budget import CHARS_PER_TOKEN_LOW
from vmai.llm_tasks import MAX_HEADER_LENGTH
from vmai.outline import CHUNK_LINES, _line_payload

ROOT = Path(__file__).resolve().parents[2]
WORD = ROOT / "fixtures" / "eval" / "generated" / "word" / "D01_so_va_dam" / "source.pdf"
OUT = ROOT / "vmai" / "out"
SYSTEM_CHARS = 300


def main() -> None:
    document = read_pdf_lines(WORD)
    lines = [l for l in document.lines if not l.repeated and not l.in_table]
    body = sum(len(json.dumps(_line_payload(l), ensure_ascii=False)) for l in lines)
    word_pages = len(document.pages)
    calls = len(lines) // CHUNK_LINES + 1
    rows = {
        "Điền từ Excel, tiêu đề khớp bảng khai báo": {"calls": 0, "tokens_in": 0, "note": "Không gọi model"},
        "Điền từ Excel, thiếu cột phải hỏi model": {"calls": 1, "tokens_in": int((SYSTEM_CHARS + 10 * MAX_HEADER_LENGTH) / CHARS_PER_TOKEN_LOW), "note": "Tối đa khoảng 10 tiêu đề"},
        f"Word đầu mục, tài liệu giả {word_pages} trang": {
            "calls": calls,
            "tokens_in": int((SYSTEM_CHARS * calls + body) / CHARS_PER_TOKEN_LOW),
            "tokens_in_per_page": int((SYSTEM_CHARS * calls + body) / CHARS_PER_TOKEN_LOW / word_pages),
            "note": "Dòng cắt còn 120 ký tự. Tài liệu hãng thật có thể hàng trăm trang",
        },
    }
    print("ƯỚC TÍNH, chưa gọi model nào. Số token thật phải đo khi có key.")
    for name, row in rows.items():
        print(f"- {name}: {json.dumps(row, ensure_ascii=False)}")
    OUT.mkdir(exist_ok=True)
    (OUT / "cost_estimate.json").write_text(json.dumps({"basis": "uoc_tinh", "rows": rows}, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
