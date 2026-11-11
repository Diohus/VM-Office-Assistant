"""Hàm vào cho backend, lát 1 (Q133, T-AI-19): Excel sang Excel theo mô tả khai sẵn, không model, không hồ sơ.

Chạy trong tiến trình con của reader: không mạng, không mở tiến trình hay luồng, không ghi file.
Lỗi luôn là `VmaiError` có `.code`, `.detail`, `.message`. `.message` là câu tiếng Việt cho người dùng, lấy từ
`docs/ai/reason_texts.md` (bản chụp `data/doc_codes.json`). Tên file và nội dung file nguồn không bao giờ nằm trong
`.detail` hay `.message`: tên file chỉ có ở `template.file_name` và `sources[].file_name` của lô (Q131).
"""
from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Literal

from .confidence import Color
from .contract import ConfidenceModel, ProposalBatch, ReasonModel
from .errors import MissingFields, SourceFileError, TemplateSpecError, VmaiError
from .headers import plan_headers
from .propose import _review_unit, propose
from .source_xlsx import read_source
from .template import load_spec, read_form

SPECS = Path(__file__).resolve().parent / "specs"
TEXTS = Path(__file__).resolve().parent / "data" / "doc_codes.json"
MAX_CHANGES = 5000  # khớp MAX_EDITS của lần lưu (Dev)
MAX_BATCH_BYTES = 5 * 1024 * 1024
UNCHANGED = "Biểu mẫu không bị đổi."
_SPEC_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$")


@lru_cache(maxsize=1)
def _texts() -> dict:
    return json.loads(TEXTS.read_text(encoding="utf-8"))


def user_message(code: str, values: dict[str, str] | None = None, *, lat1: bool = True) -> str:
    """Câu cho người dùng: câu hiện, "Biểu mẫu không bị đổi." nếu chưa có, rồi việc làm tiếp. Không có chỗ trống `{...}`.

    `lat1`: dùng cột "Làm gì ở lát 1" khi mã có ở đó, vì lát 1 chưa có màn xác nhận, màn chọn cột (Q133).
    """
    texts = _texts()
    if code not in texts["reason_texts"]:
        code = "FILL_FAILED"
    sentence = (texts.get("texts_lat1", {}).get(code) if lat1 else None) or texts["reason_texts"][code]
    lat1_actions = texts.get("actions_lat1", {})
    # Có dòng ở bảng lát 1 mà ô "Làm gì" trống thì không có câu làm gì (N3), giống giao diện dùng `??`.
    action = lat1_actions[code] if lat1 and code in lat1_actions else texts["actions"].get(code, "")

    def fill(text: str) -> str:
        for key, value in (values or {}).items():
            text = text.replace("{" + key + "}", value)
        return re.sub(r"\s*:?\s*\{[^}]*\}", "", text)  # chỗ trống không có giá trị thì bỏ, không để lộ `{...}`

    sentence, action = fill(sentence), fill(action)
    parts = [sentence.strip()]
    if UNCHANGED not in sentence:
        parts.append(UNCHANGED)
    if action.strip():
        parts.append(action.strip())
    return " ".join(p if p.endswith((".", "?", "!")) else p + "." for p in parts)


class FillError(VmaiError):
    """Lỗi của `fill_from_excel`. `detail` không chứa tên file hay nội dung file nguồn."""

    def __init__(self, code: str, detail: str, values: dict[str, str] | None = None):
        super().__init__(code, detail)
        self.message = user_message(code, values)


def _scrub(text: str, secrets: list[str]) -> str:
    for secret in sorted({s for s in secrets if s}, key=len, reverse=True):
        text = text.replace(secret, "file")
    return text


def _spec(spec_id: str):
    path = SPECS / f"{spec_id}.yaml"
    if not _SPEC_ID.match(spec_id) or not path.is_file():
        raise FillError("SPEC_INVALID", "Không có mô tả biểu mẫu với mã này trong vmai/specs")
    return load_spec(path)


def _cap_green(batch: ProposalBatch) -> ProposalBatch:
    """Q133: lát 1 không có ô xanh. Hạ xanh xuống vàng kèm `MAP_GREEN_DISABLED`, không bao giờ nâng."""
    changes = []
    for change in batch.changes:
        c = change.confidence
        if c.final_color is Color.GREEN:
            confidence = ConfidenceModel(
                read_color=c.read_color,
                map_color=c.map_color,
                final_color=Color.YELLOW,
                reasons=[*c.reasons, ReasonModel(code="MAP_GREEN_DISABLED", detail="Lát 1 chưa đánh dấu Khớp file đã chọn")],
            )
            has_existing = change.previous.value is not None or change.previous.formula is not None
            review = change.review.model_copy(update={"unit": _review_unit(Color.YELLOW, has_existing)})
            change = change.model_copy(update={"confidence": confidence, "review": review})
        changes.append(change)
    return batch.model_copy(update={"changes": changes})


def fill_from_excel(
    template_path: Path | str,
    source_path: Path | str,
    *,
    spec_id: str,
    template_name: str | None = None,
    source_name: str | None = None,
    max_color: Literal["yellow", "green"] = "yellow",
) -> ProposalBatch:
    """Đọc biểu mẫu theo mô tả khai sẵn `spec_id`, đọc file nguồn Excel, trả lô đề xuất. Không model, không hồ sơ.

    `template_name`, `source_name`: tên người dùng thấy, khi reader dùng đường dẫn tạm. `max_color` mặc định "yellow".
    """
    if max_color not in ("yellow", "green"):
        raise FillError("FILL_FAILED", "max_color không hợp lệ")
    template_path, source_path = Path(template_path), Path(source_path)
    secrets = [str(template_path), str(source_path), template_path.name, source_path.name, template_path.stem, source_path.stem, template_name or "", source_name or ""]
    spec = _spec(spec_id)
    try:
        form = read_form(template_path, spec)
        source = read_source(source_path)
        batch = propose(form, source, plan_headers(form, source))
    except MissingFields as exc:
        # Chỉ nhãn của phía biểu mẫu (tên cột mô tả chờ đợi), không tiêu đề hay giá trị lấy từ nguồn.
        labels = ", ".join(spec.source_fields[f].names[0] if f in spec.source_fields and spec.source_fields[f].names else f for f in exc.fields)
        raise FillError("MISSING_FIELDS", f"Chưa ghép được: {labels}", {"cột": labels, "tên cột": labels}) from None
    except SourceFileError as exc:
        # Chi tiết gốc có thể chứa tên sheet hay chữ của file nguồn: không đưa ra.
        raise FillError(exc.code, "Lỗi đọc file nguồn") from None
    except TemplateSpecError as exc:
        raise FillError(exc.code, _scrub(exc.detail, secrets)) from None
    except VmaiError as exc:
        raise FillError(exc.code, _scrub(exc.detail, secrets)) from None
    except Exception as exc:  # noqa: BLE001 - không để lộ nội dung lỗi gốc
        raise FillError("FILL_FAILED", type(exc).__name__) from None

    if max_color != "green":
        batch = _cap_green(batch)
    update = {}
    if template_name:
        update["template"] = batch.template.model_copy(update={"file_name": template_name})
    if source_name:
        update["sources"] = [s.model_copy(update={"file_name": source_name}) for s in batch.sources]
    if update:
        batch = batch.model_copy(update=update)
    if len(batch.changes) > MAX_CHANGES or len(batch.model_dump_json().encode("utf-8")) > MAX_BATCH_BYTES:
        limit = f"{MAX_CHANGES:,}".replace(",", ".")
        raise FillError("BATCH_TOO_LARGE", f"{len(batch.changes)} ô", {"giới hạn": limit})
    return ProposalBatch.model_validate_json(batch.model_dump_json())  # kiểm lại toàn bộ lô theo hợp đồng
