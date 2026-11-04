"""Phân loại công thức khi MỞ biểu mẫu: công thức gọi ra ngoài không được đưa cho bộ tính của lưới.

Quy tắc nằm ở vmai.security, không lặp lại ở đây. INDIRECT, OFFSET không gọi ra ngoài nên vẫn hiện.
"""
from __future__ import annotations

from vmai.security import check_formula


def blocked_reason(formula: str) -> dict | None:
    """Trả lý do nếu công thức không được đưa cho lưới. Công thức trong định dạng có điều kiện không có dấu = đầu."""
    check = check_formula("=" + formula.lstrip("="))
    if check.block_on_open:
        return {"kind": check.risk.value, "detail": check.reason}
    return None
