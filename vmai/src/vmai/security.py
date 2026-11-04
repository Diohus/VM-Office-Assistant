"""Kiểm tra công thức. Backend chạy trên mọi công thức AI đề xuất và mọi biểu mẫu mở lên, không tin lớp AI.

Chỉ dùng thư viện chuẩn, để cổng mở file gọi được mà không kéo theo gì khác.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum


class FormulaRisk(str, Enum):
    NOT_FORMULA = "not_formula"
    NETWORK = "network"  # gọi hoặc mở địa chỉ ngoài
    EXTERNAL_LINK = "external_link"  # tham chiếu file khác
    DDE = "dde"  # gọi chương trình khác
    ENVIRONMENT = "environment"  # đọc thông tin máy, đường dẫn file
    UNDETERMINABLE_INPUTS = "undeterminable_inputs"  # không xác định được ô đầu vào để tính màu


# Nhóm chặn hẳn khi mở biểu mẫu. UNDETERMINABLE_INPUTS không nguy hiểm tự thân: vẫn hiện, nhưng AI không được đề xuất.
BLOCK_ON_OPEN = frozenset({FormulaRisk.NETWORK, FormulaRisk.EXTERNAL_LINK, FormulaRisk.DDE, FormulaRisk.ENVIRONMENT})

_FUNCTIONS: dict[FormulaRisk, frozenset[str]] = {
    FormulaRisk.NETWORK: frozenset(
        {"WEBSERVICE", "FILTERXML", "HYPERLINK", "IMPORTXML", "IMPORTDATA", "IMPORTHTML", "IMPORTFEED", "IMPORTRANGE", "IMAGE"}
    ),
    FormulaRisk.DDE: frozenset({"DDE", "DDEAUTH", "RTD", "CALL", "REGISTER", "REGISTER.ID", "EXEC"}),
    FormulaRisk.ENVIRONMENT: frozenset({"INFO", "CELL"}),
    FormulaRisk.UNDETERMINABLE_INPUTS: frozenset({"INDIRECT", "OFFSET"}),
}
_RISK_BY_FUNCTION = {name: risk for risk, names in _FUNCTIONS.items() for name in names}
_STRING = re.compile(r'"(?:[^"]|"")*"')
_FUNCTION = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
_URL = re.compile(r"\b(?:https?|ftp|file)://", re.IGNORECASE)
# Table của Excel dùng dấu ngoặc vuông như BangTieuHao[Mã], không phải liên kết ngoài.
_EXTERNAL_BOOK = re.compile(r"(?:^|[=+\-*/(,'\s])\[[^\]]+\][^\[\]!]*!|'[^']*\[[^\]]+\][^']*'!")


@dataclass(frozen=True)
class FormulaCheck:
    allowed: bool
    risk: FormulaRisk | None = None
    reason: str = ""

    @property
    def block_on_open(self) -> bool:
        return self.risk in BLOCK_ON_OPEN


def check_formula(expression: str) -> FormulaCheck:
    """`allowed = False` nghĩa là AI không được đề xuất công thức này. Danh sách chặn nên có thể thiếu."""
    if not expression.startswith("="):
        return FormulaCheck(False, FormulaRisk.NOT_FORMULA, "Không phải công thức")
    # Bỏ chuỗi trong dấu nháy kép để tên hàm trong chuỗi không bị tính nhầm,
    # nhưng tên sheet trong nháy đơn vẫn được xét vì liên kết ngoài nằm ở đó.
    bare = _STRING.sub('""', expression)
    if "|" in bare:
        return FormulaCheck(False, FormulaRisk.DDE, "Có dạng gọi DDE")
    if _URL.search(bare):
        return FormulaCheck(False, FormulaRisk.NETWORK, "Có địa chỉ mạng")
    if _EXTERNAL_BOOK.search(bare):
        return FormulaCheck(False, FormulaRisk.EXTERNAL_LINK, "Có liên kết tới file ngoài")
    for name in _FUNCTION.findall(bare):
        risk = _RISK_BY_FUNCTION.get(name.upper().removeprefix("_XLFN.").removeprefix("_XLWS."))
        if risk is not None:
            return FormulaCheck(False, risk, f"Hàm bị cấm: {name.upper()}")
    return FormulaCheck(True)
