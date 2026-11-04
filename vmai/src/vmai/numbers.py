"""Chuẩn hóa số theo docs/ai/confidence_rules.md, mục Chuẩn hóa số."""
from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from enum import Enum


class NumberStyle(str, Enum):
    COMMA_DECIMAL = "comma_decimal"  # 1.234,5
    DOT_DECIMAL = "dot_decimal"  # 1,234.5


class ParseStatus(str, Enum):
    OK = "ok"
    AMBIGUOUS = "ambiguous"
    INVALID = "invalid"


@dataclass(frozen=True)
class ParsedNumber:
    status: ParseStatus
    normalized: str | None = None
    style: NumberStyle | None = None  # chỉ có khi cách viết tự nó rõ ràng


_SHAPE = re.compile(r"^[+-]?\d[\d.,]*$")


def _ok(digits: str, style: NumberStyle | None) -> ParsedNumber:
    try:
        value = Decimal(digits)
    except InvalidOperation:
        return ParsedNumber(ParseStatus.INVALID)
    return ParsedNumber(ParseStatus.OK, format(value, "f"), style)


def _valid_grouping(integer: str, sep: str) -> bool:
    groups = integer.lstrip("+-").split(sep)
    return 1 <= len(groups[0]) <= 3 and all(len(g) == 3 for g in groups[1:])


def parse_number(text: str, column_style: NumberStyle | None = None) -> ParsedNumber:
    """Đọc một chuỗi thành số. Không đoán khi nhập nhằng."""
    s = text.strip().replace(" ", "").replace(" ", "")
    if not _SHAPE.match(s):
        return ParsedNumber(ParseStatus.INVALID)

    has_dot, has_comma = "." in s, "," in s
    if not has_dot and not has_comma:
        return _ok(s, None)

    if has_dot and has_comma:
        if s.rfind(",") > s.rfind("."):
            decimal_sep, group_sep, style = ",", ".", NumberStyle.COMMA_DECIMAL
        else:
            decimal_sep, group_sep, style = ".", ",", NumberStyle.DOT_DECIMAL
        if s.count(decimal_sep) != 1:
            return ParsedNumber(ParseStatus.INVALID)
        integer, fraction = s.split(decimal_sep)
        if not fraction or not _valid_grouping(integer, group_sep):
            return ParsedNumber(ParseStatus.INVALID)
        return _ok(integer.replace(group_sep, "") + "." + fraction, style)

    sep = "." if has_dot else ","
    same_style = NumberStyle.DOT_DECIMAL if sep == "." else NumberStyle.COMMA_DECIMAL
    group_style = NumberStyle.COMMA_DECIMAL if sep == "." else NumberStyle.DOT_DECIMAL

    if s.count(sep) > 1:
        if not _valid_grouping(s, sep):
            return ParsedNumber(ParseStatus.INVALID)
        return _ok(s.replace(sep, ""), group_style)

    integer, fraction = s.split(sep)
    if not fraction:
        return ParsedNumber(ParseStatus.INVALID)
    if len(fraction) != 3 or integer.lstrip("+-") == "0":
        return _ok(f"{integer}.{fraction}", same_style)

    # 1.234 hoặc 1,234: chỉ gỡ được khi cột có bằng chứng về cách viết.
    if column_style is None:
        return ParsedNumber(ParseStatus.AMBIGUOUS)
    if column_style is same_style:
        return _ok(f"{integer}.{fraction}", same_style)
    if not _valid_grouping(s, sep):
        return ParsedNumber(ParseStatus.AMBIGUOUS)
    return _ok(s.replace(sep, ""), group_style)


def infer_column_style(texts: Iterable[str]) -> NumberStyle | None:
    """Trả về cách viết của cột khi mọi ô rõ ràng cùng một kiểu, ngược lại None."""
    styles = {
        parsed.style
        for parsed in (parse_number(t) for t in texts)
        if parsed.status is ParseStatus.OK and parsed.style is not None
    }
    return styles.pop() if len(styles) == 1 else None
