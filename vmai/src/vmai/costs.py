"""Chi phí mỗi biểu mẫu, cho trang số liệu (Q50, Q55). Mọi số phải ghi là đo hay ước tính.

Thuần: nhận danh sách `LlmCall` đã ghi. Không gọi gì.
"""
from __future__ import annotations

import statistics
from collections import defaultdict
from collections.abc import Iterable
from typing import Literal

from pydantic import BaseModel, ConfigDict

from .contract import LlmCall
from .llm.budget import Price


class FormCost(BaseModel):
    model_config = ConfigDict(frozen=True)

    form_id: str
    form_type: str
    calls: int
    input_tokens: int
    output_tokens: int
    usd: float | None
    basis: Literal["do", "tinh_theo_bang_gia", "khong_goi_model", "thieu_so_lieu"]
    models: tuple[str, ...] = ()


def cost_of_form(form_id: str, form_type: str, calls: Iterable[LlmCall], price: Price | None = None) -> FormCost:
    items = list(calls)
    if not items:
        return FormCost(form_id=form_id, form_type=form_type, calls=0, input_tokens=0, output_tokens=0, usd=0.0, basis="khong_goi_model")
    input_tokens = sum(c.input_tokens or 0 for c in items)
    output_tokens = sum(c.output_tokens or 0 for c in items)
    models = tuple(sorted({c.model for c in items}))
    if any(c.input_tokens is None or c.output_tokens is None for c in items) and not all(c.cost_usd is not None for c in items):
        return FormCost(form_id=form_id, form_type=form_type, calls=len(items), input_tokens=input_tokens, output_tokens=output_tokens, usd=None, basis="thieu_so_lieu", models=models)
    if all(c.cost_usd is not None for c in items):
        usd, basis = sum(c.cost_usd for c in items), "do"
    elif price is not None:
        usd, basis = input_tokens / 1e6 * price.input_per_mtok + output_tokens / 1e6 * price.output_per_mtok, "tinh_theo_bang_gia"
    else:
        usd, basis = None, "thieu_so_lieu"
    return FormCost(form_id=form_id, form_type=form_type, calls=len(items), input_tokens=input_tokens, output_tokens=output_tokens, usd=usd, basis=basis, models=models)


def summarize_by_type(costs: Iterable[FormCost]) -> dict[str, dict]:
    """Trung vị và cao nhất theo loại biểu mẫu. Không gộp các loại, vì Excel và Word khác nhau rất xa."""
    groups: dict[str, list[FormCost]] = defaultdict(list)
    for cost in costs:
        groups[cost.form_type].append(cost)
    out = {}
    for form_type, items in groups.items():
        usd = [c.usd for c in items if c.usd is not None]
        out[form_type] = {
            "forms": len(items),
            "median_usd": statistics.median(usd) if usd else None,
            "max_usd": max(usd) if usd else None,
            "median_tokens": statistics.median(c.input_tokens + c.output_tokens for c in items),
            "missing": sum(c.usd is None for c in items),
            "bases": sorted({c.basis for c in items}),
            "models": sorted({m for c in items for m in c.models}),
        }
    return out
