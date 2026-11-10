"""Chạy các ca X02, X16, X16b, X16c, X16d, X17 với model thật (Q26, Q43). Chỉ dữ liệu giả, chỉ gửi tiêu đề.

Không có --confirm-network thì chỉ in những gì SẼ gửi. Ca không cần hỏi model thì không gọi model.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import vmai
from vmai import MissingFields, load_spec, map_headers, plan_headers, propose, read_form, read_source
from vmai.llm.factory import create_provider, load_profile
from vmai.llm.keys import read_api_key

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _env import load_key  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
GENERATED = ROOT / "fixtures" / "eval" / "generated"
OUT = ROOT / "vmai" / "out"
SPEC = load_spec(Path(vmai.__file__).parent / "specs" / "GIA-QC-01.yaml")
CASES = {
    "X02": GENERATED / "xlsx" / "X02",
    "X16": GENERATED / "xlsx" / "X16",
    "X16b": GENERATED / "xlsx_lat1" / "X16b",
    "X16c": GENERATED / "xlsx_lat1" / "X16c",
    "X16d": GENERATED / "xlsx_lat1" / "X16d",
    "X17": GENERATED / "xlsx" / "X17",
}
FORBIDDEN = ("Nguyễn Giả Một", "Lê Giả", "Phạm Giả", "Hoàng Giả", "vi-du.invalid", "9.99")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", default=None)
    parser.add_argument("--confirm-network", action="store_true")
    args = parser.parse_args()
    profile = load_profile(args.profile)
    provider = None
    if args.confirm_network:
        load_key(profile.api_key_env)
        read_api_key(profile)
        provider = create_provider(profile)
    results = {}
    for name, directory in CASES.items():
        form, source = read_form(directory / "template.xlsx", SPEC), read_source(directory / "source.xlsx")
        plan = plan_headers(form, source)
        row: dict = {"sent_headers": list(plan.llm_headers), "called_model": False, "phi_headers": list(plan.phi_headers), "person_headers": list(plan.person_headers)}
        mapping, calls = None, ()
        if plan.llm_headers and provider is not None:
            mapping, call = map_headers(provider, list(plan.llm_headers), plan.llm_fields)
            calls = (call,)
            row.update(called_model=True, mapping=mapping, input_tokens=call.input_tokens, output_tokens=call.output_tokens, latency_ms=call.latency_ms)
        try:
            batch = propose(form, source, plan, mapping, calls)
        except MissingFields as exc:
            row["outcome"] = f"Hỏi người dùng: thiếu {exc.fields}"
        else:
            colors: dict[str, int] = {}
            for change in batch.changes:
                colors[change.confidence.final_color.value] = colors.get(change.confidence.final_color.value, 0) + 1
            dumped = batch.model_dump_json()
            row.update(outcome="Có đề xuất", colors=colors, proposals=len(batch.changes), leaks=[f for f in FORBIDDEN if f in dumped])
            row["all_values_from_source"] = _values_from_source(batch, source)
        results[name] = row
        print(name, json.dumps(row, ensure_ascii=False))
    if provider is not None:
        results["_total"] = {"model": profile.model, "calls": len(provider.calls), "tokens": provider.tokens_used, "usd_by_price_table": provider.usd_used}
        print("Tổng:", json.dumps(results["_total"], ensure_ascii=False))
        OUT.mkdir(exist_ok=True)
        (OUT / f"model_flows_{profile.name}.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")


def _values_from_source(batch, source) -> bool:
    """Mọi số đề xuất phải có trong nguồn: model không có đường tạo ra số."""
    values = {c.value for sheet in source.sheets for row in sheet.rows for c in row}
    return all(c.value.raw_text in values for c in batch.changes)


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
