"""Bản chụp mã lý do từ tài liệu nội bộ, để test chạy được trên kho công khai không có docs/ (T-AI-17, Q128).

Tạo lại sau khi sửa docs/ai/confidence_rules.md hoặc docs/ai/reason_texts.md, từ `vmai/`:
    .venv/Scripts/python.exe tests/doc_snapshot.py
Trên cây đầy đủ, `test_rules.py::test_doc_snapshot_matches_the_docs` báo lỗi nếu bản chụp cũ.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs" / "ai"
SNAPSHOT = Path(__file__).resolve().parent / "data" / "doc_codes.json"
_CODE = re.compile(r"`([A-Z]{3,}_[A-Z0-9_]+)`")
_ROW_CODE = re.compile(r"^`([A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+)`$")


def _sentences(text: str) -> dict[str, str]:
    """Cột "Câu hiện" của mỗi bảng, theo dòng tiêu đề của bảng đó. Dòng đầu tiên của một mã thắng."""
    out: dict[str, str] = {}
    column = None
    for line in text.splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")] if line.startswith("|") else None
        if cells is None:
            column = None
        elif "Câu hiện" in cells:
            column = cells.index("Câu hiện")
        elif column is not None and (m := _ROW_CODE.match(cells[0])) and column < len(cells):
            out.setdefault(m.group(1), cells[column])
    return out


def _ui_texts(text: str) -> dict[str, str]:
    """Bảng "Câu không gắn mã": cột "Chỗ hiện" -> "Câu hiện". Giao diện so câu với bản chụp này (phần 3)."""
    out: dict[str, str] = {}
    columns = None
    for line in text.splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")] if line.startswith("|") else None
        if cells is None:
            columns = None
        elif "Chỗ hiện" in cells and "Câu hiện" in cells:
            columns = (cells.index("Chỗ hiện"), cells.index("Câu hiện"))
        elif columns and not set(cells[0]) <= set("-: ") and max(columns) < len(cells):
            out.setdefault(cells[columns[0]], cells[columns[1]])
    return out


def build() -> dict:
    rules = (DOCS / "confidence_rules.md").read_text(encoding="utf-8")
    texts = (DOCS / "reason_texts.md").read_text(encoding="utf-8")
    sentences = _sentences(texts)
    mentioned = set(re.findall(r"`([A-Z][A-Z0-9]*(?:_[A-Z0-9]+)+)`", texts))
    return {
        "note": "Sinh bởi vmai/tests/doc_snapshot.py từ docs/ai. Không sửa tay.",
        "confidence_rules_codes": sorted(set(_CODE.findall(rules))),
        "reason_texts": {code: sentences.get(code, "") for code in sorted(mentioned)},
        "ui_texts": _ui_texts(texts),
    }


def load() -> dict:
    return json.loads(SNAPSHOT.read_text(encoding="utf-8"))


if __name__ == "__main__":
    SNAPSHOT.parent.mkdir(exist_ok=True)
    SNAPSHOT.write_text(json.dumps(build(), ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    print(f"Đã ghi {SNAPSHOT}")
