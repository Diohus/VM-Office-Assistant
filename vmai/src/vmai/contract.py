"""Hợp đồng đề xuất dạng code, theo docs/contracts/proposal_schema.md (phiên bản 0.3)."""
from __future__ import annotations

from enum import Enum
from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from .confidence import Color, SourceKind

SCHEMA_VERSION = "0.3"


class _Model(BaseModel):
    model_config = ConfigDict(extra="forbid")


class RowKeyPart(_Model):
    column_path: list[str]
    value: str


class CellTarget(_Model):
    sheet: str
    table_anchor: str
    row_key: list[RowKeyPart]
    row_offset: int | None = None
    col_path: list[str]
    a1_at_write: str


class Value(_Model):
    type: Literal["number", "text", "date", "boolean"]
    normalized: str
    raw_text: str
    unit: str | None = None


class Previous(_Model):
    value: Value | None = None
    formula: str | None = None


class Formula(_Model):
    expression: str
    inputs: list[CellTarget]
    proposed_by: Literal["code", "llm", "template"]


class XlsxCellLocator(_Model):
    type: Literal["xlsx_cell"] = "xlsx_cell"
    sheet: str
    a1: str


class PdfTextLocator(_Model):
    type: Literal["pdf_text"] = "pdf_text"
    page: int
    bbox: tuple[float, float, float, float]


class ImageRegionLocator(_Model):
    type: Literal["image_region"] = "image_region"
    page: int
    bbox: tuple[float, float, float, float]


Locator = Annotated[XlsxCellLocator | PdfTextLocator | ImageRegionLocator, Field(discriminator="type")]


class Evidence(_Model):
    source_id: str
    reader: str
    read_text: str
    locator: Locator


class MappingMethod(str, Enum):
    EXACT = "exact"
    PROFILE = "profile"
    NORMALIZED = "normalized"
    ABBREVIATION = "abbreviation"
    INFERRED = "inferred"
    FUZZY = "fuzzy"
    LLM_INFERRED = "llm_inferred"
    POSITIONAL = "positional"
    NONE = "none"


class Candidate(_Model):
    col_path: list[str]
    score: float


class Mapping(_Model):
    method: MappingMethod
    source_header: str | None = None
    profile_id: str | None = None
    profile_version: int | None = None
    candidates: list[Candidate] = []


class ReasonModel(_Model):
    code: str
    detail: str


class ConfidenceModel(_Model):
    read_color: Color
    map_color: Color
    final_color: Color
    reasons: list[ReasonModel]


class CrossCheck(_Model):
    check: Literal["printed_total", "decimal_places", "historical_range", "valid_range"]
    status: Literal["pass", "fail", "not_applicable"]
    detail: str = ""


class Review(_Model):
    status: Literal["pending", "approved", "rejected", "edited"] = "pending"
    unit: Literal["group", "row", "cell"]
    by: str | None = None
    at: str | None = None
    edited_value: Value | None = None


class Change(_Model):
    change_id: str
    op: Literal["set_value", "set_formula", "add_row", "add_column"]
    target: CellTarget
    previous: Previous = Previous()
    value: Value | None = None
    formula: Formula | None = None
    evidence: list[Evidence]
    mapping: Mapping
    confidence: ConfidenceModel
    cross_checks: list[CrossCheck] = []
    review: Review

    @model_validator(mode="after")
    def _final_not_above_inputs(self) -> Change:
        ranks = {Color.RED: 0, Color.YELLOW: 1, Color.GREEN: 2}
        c = self.confidence
        if self.op == "set_value" and ranks[c.final_color] > min(ranks[c.read_color], ranks[c.map_color]):
            raise ValueError("final_color không được cao hơn màu đọc và màu ghép")
        if self.op == "set_value" and not self.evidence:
            if not any(r.code == "NO_SOURCE" for r in c.reasons):
                raise ValueError("set_value không có evidence phải có lý do NO_SOURCE")
        return self


class Page(_Model):
    page: int
    width: int
    height: int
    page_image_sha256: str


class Source(_Model):
    source_id: str
    file_name: str
    source_file_sha256: str
    kind: SourceKind
    pages: list[Page] = []


class Template(_Model):
    document_uuid: str | None = None
    template_sha256: str
    file_name: str
    spec_id: str | None = None
    spec_version: int | None = None


class Notice(_Model):
    reason: str
    count: int
    change_ids: list[str] = []
    detail: str


class LlmCall(_Model):
    purpose: str
    provider: str
    model: str
    request_id: str | None = None
    input_tokens: int | None = None
    output_tokens: int | None = None
    latency_ms: int | None = None
    cost_usd: float | None = None


class ProposalBatch(_Model):
    schema_version: Literal["0.3"] = SCHEMA_VERSION
    rule_version: str
    proposal_id: str
    created_at: str
    template: Template
    sources: list[Source]
    changes: list[Change]
    notices: list[Notice] = []
    llm_calls: list[LlmCall] = []
