"""So tiêu đề và giá trị với bảng khai báo. Không gọi LLM."""
from __future__ import annotations

import difflib
import re
import unicodedata
from collections.abc import Mapping
from dataclasses import dataclass

from .confidence import ColumnMatch

FUZZY_MIN = 0.84
FUZZY_MARGIN = 0.08


def normalize_text(text: str) -> str:
    """Bỏ dấu, chữ thường, gom khoảng trắng và dấu câu."""
    decomposed = unicodedata.normalize("NFD", str(text).replace("đ", "d").replace("Đ", "D"))
    stripped = "".join(c for c in decomposed if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", " ", stripped.lower()).strip()


_HEADER_UNIT = re.compile(r"^(.*\S)\s*[(\[]([^()\[\]]+)[)\]]\s*$")
_BARE_UNITS = frozenset({"%", "g", "mg", "ug", "ng", "pg", "kg", "l", "dl", "ml", "ul", "fl", "mol", "mmol", "umol", "nmol", "iu", "u", "meq", "vnd"})


def split_header_unit(text: str) -> tuple[str, str | None]:
    """Tách `HGB (g/L)` thành tên và đơn vị. Ngoặc không giống đơn vị, như `Natri (máy 2)`, vẫn thuộc về tên.

    Bỏ đơn vị chỉ để so tên. Đơn vị phải được so tiếp bằng `same_unit`, nếu không số sai gấp 10 lần sẽ ghép im lặng (R9b).
    """
    match = _HEADER_UNIT.match(str(text).strip())
    if not match:
        return str(text).strip(), None
    inner = match.group(2).strip()
    lowered = inner.lower().replace("µ", "u").replace("μ", "u")
    if ("/" in inner and any(c.isalpha() or c == "%" for c in inner)) or lowered in _BARE_UNITS:
        return match.group(1).strip(), inner
    return str(text).strip(), None


def same_unit(a: str, b: str) -> bool:
    """So đơn vị, giữ chữ hoa thường: `G/L` (10^9/L) khác `g/L` (gam/L). Riêng lít viết `l` hay `L` là một."""

    def canon(text: str) -> str:
        cleaned = unicodedata.normalize("NFKC", text).replace("µ", "u").replace("μ", "u")
        return "".join(cleaned.split()).replace("l", "L")

    return canon(a) == canon(b)


@dataclass(frozen=True)
class FieldMatch:
    match: ColumnMatch
    field: str | None
    candidates: tuple[tuple[str, float], ...] = ()


def match_field(
    text: str,
    fields: Mapping[str, tuple[str, ...]],
    aliases: Mapping[str, tuple[str, ...]],
    *,
    fuzzy: bool = True,
) -> FieldMatch:
    """So một chuỗi với các trường đã biết. `aliases` là từ điển viết tắt, luôn thấp hơn khớp chính xác."""
    exact = [f for f, names in fields.items() if text in names]
    if len(exact) == 1:
        return FieldMatch(ColumnMatch.EXACT, exact[0])
    if len(exact) > 1:
        return FieldMatch(ColumnMatch.AMBIGUOUS, None)

    norm = normalize_text(text)
    if not norm:
        return FieldMatch(ColumnMatch.NONE, None)

    hits = [f for f, names in fields.items() if norm in {normalize_text(n) for n in names}]
    if len(hits) == 1:
        return FieldMatch(ColumnMatch.NORMALIZED, hits[0])
    if len(hits) > 1:
        return FieldMatch(ColumnMatch.AMBIGUOUS, None)

    hits = [f for f, names in aliases.items() if norm in {normalize_text(n) for n in names}]
    if len(hits) == 1:
        return FieldMatch(ColumnMatch.ABBREVIATION, hits[0])
    if len(hits) > 1:
        return FieldMatch(ColumnMatch.AMBIGUOUS, None)

    if not fuzzy or not fields:
        return FieldMatch(ColumnMatch.NONE, None)

    def closeness(name: str) -> float:
        # `HGB (g/dL)` rất giống `HGB (g/L)` theo chữ nhưng là số khác 10 lần: khác đơn vị thì không khớp gần đúng.
        ours, theirs = split_header_unit(text)[1], split_header_unit(name)[1]
        if ours and theirs and not same_unit(ours, theirs):
            return 0.0
        # `Mức 1` rất giống `Mức 2` theo chữ nhưng là cột khác. Số trong tên phải giống hệt, nếu không số vào nhầm cột.
        if re.findall(r"\d+", norm) != re.findall(r"\d+", normalize_text(name)):
            return 0.0
        return difflib.SequenceMatcher(None, norm, normalize_text(name)).ratio()

    scored = sorted(
        ((f, max(closeness(n) for n in names)) for f, names in fields.items()),
        key=lambda item: -item[1],
    )
    if scored[0][1] < FUZZY_MIN:
        return FieldMatch(ColumnMatch.NONE, None, tuple(scored[:3]))
    if len(scored) > 1 and scored[0][1] - scored[1][1] < FUZZY_MARGIN:
        return FieldMatch(ColumnMatch.AMBIGUOUS, None, tuple(scored[:3]))
    return FieldMatch(ColumnMatch.FUZZY, scored[0][0], tuple(scored[:3]))
