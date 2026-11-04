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
    scored = sorted(
        (
            (f, max(difflib.SequenceMatcher(None, norm, normalize_text(n)).ratio() for n in names))
            for f, names in fields.items()
        ),
        key=lambda item: -item[1],
    )
    if scored[0][1] < FUZZY_MIN:
        return FieldMatch(ColumnMatch.NONE, None, tuple(scored[:3]))
    if len(scored) > 1 and scored[0][1] - scored[1][1] < FUZZY_MARGIN:
        return FieldMatch(ColumnMatch.AMBIGUOUS, None, tuple(scored[:3]))
    return FieldMatch(ColumnMatch.FUZZY, scored[0][0], tuple(scored[:3]))
