"""So số LibreOffice tính được với số trên lưới, trước khi xuất PDF (Q48c). Hàm đổi theo thời điểm thì không so."""
from __future__ import annotations

from datetime import date, datetime, time

TOLERANCE = 1e-6
_EPOCH = datetime(1899, 12, 30)


def _serial(text: str) -> float | None:
    """Reader trả ngày dạng ISO, lưới giữ số ngày như Excel. Đổi về số ngày để so, không phải ngày thì trả None."""
    for parse in (datetime.fromisoformat, date.fromisoformat):
        try:
            value = parse(text)
        except ValueError:
            continue
        moment = value if isinstance(value, datetime) else datetime.combine(value, time())
        return (moment - _EPOCH).total_seconds() / 86400
    try:
        t = time.fromisoformat(text)
    except ValueError:
        return None
    return (t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6) / 86400


def _same(pdf, grid) -> bool:
    if isinstance(pdf, str) and isinstance(grid, (int, float)) and not isinstance(grid, bool):
        pdf = _serial(pdf) if _serial(pdf) is not None else pdf
    if pdf in (None, "") and grid in (None, ""):
        return True
    if isinstance(pdf, bool) or isinstance(grid, bool):
        return bool(pdf) == bool(grid) and pdf not in (None, "") and grid not in (None, "")
    if isinstance(pdf, (int, float)) and isinstance(grid, (int, float)):
        return abs(pdf - grid) <= TOLERANCE * max(1, abs(pdf), abs(grid))
    return str(pdf) == str(grid)


def compare(computed: list[dict], grid: list[dict]) -> list[dict]:
    """Trả các ô lệch. Ô có trên lưới mà LibreOffice không có cũng tính là lệch: không xuất khi chưa chắc."""
    pdf = {(c["sheet"], c["cell"]): c for c in computed}
    out = []
    for g in grid:
        key = (g["sheet"], g["cell"])
        c = pdf.get(key)
        if c is not None and c.get("volatile"):
            continue
        value = c["value"] if c is not None else None
        if not _same(value, g.get("value")):
            out.append({"sheet": g["sheet"], "cell": g["cell"], "grid": g.get("value"), "pdf": value})
    return out
