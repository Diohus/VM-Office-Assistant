"""Eval tầng cổng của khung chat (Q135): model giả trả kết quả xấu, xem 7 cổng có chặn đủ không. Không gọi mạng.

Chạy từ `vmai/`:  .venv/Scripts/python.exe eval/run_chat_gates.py
Thoát mã 1 khi có op lọt (qua cổng mà đáp án nói phải chặn), lệch đáp án, hoặc chuỗi cấm có trong nội dung gửi đi.
Đáp án do AI Engineer đặt, là giả định. Đạt chỉ cho biết các ca đã nghĩ ra được chặn đúng, không chứng minh model thật an toàn.
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
from collections import Counter
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _console import utf8  # noqa: E402

from vmai.chat import ChatInputError, ChatRequest, check_chat, parse_cell, parse_range, prepare_chat, to_llm_request  # noqa: E402
from vmai.contract import LlmCall  # noqa: E402
from vmai.llm.base import LLMResponse, Profile  # noqa: E402
from vmai.llm.fake import FakeProvider  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CASES = ROOT / "fixtures" / "eval" / "chat_cases.json"
OUT = Path(__file__).resolve().parents[1] / "out" / "eval_chat_gates.json"


def _cell(c, cached, dates: dict[str, dt.date]) -> dict | None:
    """Giống trình duyệt gửi: t, v là kết quả đang hiện, f là công thức."""
    value, formula = c.value, None
    if value is None:
        return None
    if isinstance(value, str) and value.startswith("="):
        formula, value = value, cached
        if c.coordinate in dates:
            value = dates[c.coordinate]
        if value is None:
            value = 0  # file do openpyxl lưu không có kết quả tính sẵn; ô công thức vẫn bị chặn ghi đè
    if isinstance(value, bool):
        t, v = "boolean", str(value).lower()
    elif isinstance(value, (dt.datetime, dt.date)):
        t, v = "date", (value.date() if isinstance(value, dt.datetime) else value).isoformat()
    elif isinstance(value, (int, float)):
        t, v = "number", repr(value) if isinstance(value, float) else str(value)
    else:
        t, v = "text", str(value)
    return {"a1": c.coordinate, "t": t, "v": v, "f": formula}


def _date_chain(ws) -> dict[str, dt.date]:
    """Cột ngày của GIA-QC-01 (=$B$4, =A8+1...): file openpyxl không có kết quả, nên tính ở đây như lưới tính."""
    start = ws["B4"].value
    if not isinstance(start, (dt.date, dt.datetime)):
        return {}
    out, day = {}, start.date() if isinstance(start, dt.datetime) else start
    for row in range(8, 38):
        f = ws[f"A{row}"].value
        if isinstance(f, str) and (f == "=$B$4" or re.fullmatch(rf"=A{row - 1}\+1", f)):
            out[f"A{row}"] = day + dt.timedelta(days=row - 8)
    return out


def request_from_xlsx(path: Path, active: str, region: str, extra: dict[str, str], message: str) -> dict:
    wb = openpyxl.load_workbook(path)
    cached = openpyxl.load_workbook(path, data_only=True)
    sheets = []
    for name, rng in [(active, region), *extra.items()]:
        ws, wc = wb[name], cached[name]
        dates = _date_chain(ws) if name == active else {}
        cells = []
        for row in ws[rng]:
            for c in row:
                item = _cell(c, wc[c.coordinate].value, dates)
                if item is not None:
                    cells.append(item)
        r1, c1, r2, c2 = parse_range(rng)
        merges = [str(m) for m in ws.merged_cells.ranges
                  if m.min_row <= r2 and m.max_row >= r1 and m.min_col <= c2 and m.max_col >= c1]
        sheets.append({"name": name, "region": rng, "cells": cells, "merges": merges})
    return {"base_version": "1", "active_sheet": active, "sheets": sheets, "message": message}


def _fake(output: dict) -> FakeProvider:
    profile = Profile(name="fake", provider="fake", model="fake-1")
    return FakeProvider(profile, allow_unverified=True, handler=lambda req: LLMResponse(
        provider="fake", model="fake-1", text=json.dumps(output, ensure_ascii=False), input_tokens=100, output_tokens=50))


def run_case(spec: dict, case: dict) -> dict:
    spec = {**spec, **case.get("source", {})}  # ca dùng file khác (2.9: cột người bệnh giả), chỉ ở tầng cổng
    req = apply_case(spec, case, request_from_xlsx(ROOT / spec["file"], spec["active_sheet"], spec["region"],
                                                    spec["extra_sheets"], case["message"]))
    expect, problems = case["expect"], []
    return _run_prepared(spec, case, req, expect, problems)


def apply_case(spec: dict, case: dict, req: dict) -> dict:
    """Thay đổi riêng của ca (bỏ ô, sheet ẩn, thêm ô chữ). Tầng cổng và tầng model dùng chung, để hai tầng đo cùng một đầu vào."""
    for sheet in req["sheets"]:
        drop = {ref.split("!")[1] for ref in case.get("drop_cells", []) if ref.split("!")[0] == sheet["name"]}
        sheet["cells"] = [c for c in sheet["cells"] if c["a1"] not in drop]
        sheet["hidden"] = sheet["name"] in case.get("hidden_sheets", [])
        if sheet["name"] == spec["active_sheet"]:
            added = case.get("add_cells", {})  # ô chữ thêm vào bản chụp, như người dùng vừa gõ trên lưới
            sheet["cells"] = [c for c in sheet["cells"] if c["a1"] not in added]
            sheet["cells"] += [{"a1": a1, "t": "text", "v": text, "f": None} for a1, text in added.items()]
            if "edit_region" in case:  # 0.2 (N2): người dùng chọn vùng; gửi kèm context_region chỉ để đọc
                ctx = parse_range(case["context_region"])
                sheet["cells"] = [c for c in sheet["cells"] if _in_range(c["a1"], ctx)]
                sheet["merges"] = [m for m in sheet.get("merges", []) if _overlaps(parse_range(m), ctx)]
                sheet["region"], sheet["context_region"] = case["edit_region"], case["context_region"]
    return req


def _col(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _in_range(a1: str, rng) -> bool:
    row, col = parse_cell(a1)
    return rng[0] <= row <= rng[2] and rng[1] <= col <= rng[3]


def _overlaps(a, b) -> bool:
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def _run_prepared(spec: dict, case: dict, req: dict, expect: dict, problems: list) -> dict:
    provider = _fake(case["model"])
    try:
        prepared = prepare_chat(ChatRequest.model_validate(req))
    except ChatInputError as err:
        if expect.get("prepare_error") != err.code:
            problems.append(f"lệch: prepare báo {err.code}")
        return {"id": case["id"], "group": case["group"], "leaks": 0, "problems": problems, "sent": False}
    if "prepare_error" in expect:
        problems.append("lọt: câu đáng lẽ không được gửi")
    for bad in spec["never_send"] + expect.get("prompt_has_no", []):
        if bad in prepared.prompt or bad in prepared.system:
            problems.append(f"lọt: chuỗi cấm trong nội dung gửi: {bad}")
    active_masked = {a for sh in prepared.context.sheets if sh.name == prepared.context.active_sheet for a in sh.masked}
    for rng in expect.get("masked_has", []):  # vùng phải bị che (2.9: cột người bệnh)
        r = parse_range(rng)
        missing = [f"{_col(c)}{row}" for row in range(r[0], r[2] + 1) for c in range(r[1], r[3] + 1)
                   if f"{_col(c)}{row}" not in active_masked]
        if missing:
            problems.append(f"lọt: ô không bị che: {missing[:3]}")
    response = provider.generate(to_llm_request(prepared))
    call = LlmCall(purpose="chat_edit", provider=response.provider, model=response.model,
                   input_tokens=response.input_tokens, output_tokens=response.output_tokens)
    plan = check_chat(prepared.context, response.text, call, prepared.masked)

    accepted = [op.cell for op in plan.ops]
    rejected = {r.cell: r.reason for r in plan.rejected_ops}
    want_ops = expect.get("ops", [])
    leaks = [c for c in accepted if c not in want_ops]
    for cell in leaks:
        problems.append(f"lọt: op {cell} qua cổng")
    for cell in want_ops:
        if cell not in accepted:
            problems.append(f"lệch: op {cell} bị chặn ({rejected.get(cell)})")
    for cell, reason in expect.get("rejected", {}).items():
        if rejected.get(cell) != reason:
            problems.append(f"lệch: {cell} mong {reason}, được {rejected.get(cell)}")
    if (code := expect.get("rejected_all")) and (not rejected or set(rejected.values()) != {code}):
        problems.append(f"lệch: mong mọi op bị {code}, được {sorted(set(rejected.values()))}")
    if plan.status != expect["status"]:
        problems.append(f"lệch: status {plan.status}, mong {expect['status']}")
    if (bad := expect.get("message_has_no")) and bad in plan.model_message:
        problems.append(f"lọt: câu model còn {bad}")
    for bad in expect.get("summary_has_no", []):  # K2: chữ ô ghi chú không được thành nhãn tóm tắt
        if any(bad in line.text for line in plan.summary):
            problems.append(f"lọt: chữ ô trong câu tóm tắt: {bad}")
    if plan.status == "edit" and len(plan.summary) != len(plan.ops):
        problems.append("lệch: thiếu dòng tóm tắt")
    return {"id": case["id"], "group": case["group"], "leaks": sum(p.startswith("lọt") for p in problems),
            "problems": problems, "sent": True, "accepted": accepted, "rejected": rejected,
            "summary": [s.text for s in plan.summary]}


def main() -> int:
    utf8()
    spec = json.loads(CASES.read_text(encoding="utf-8"))
    results = [run_case(spec, case) for case in spec["cases"]]
    by_group = Counter(r["group"] for r in results)
    failed = [r for r in results if r["problems"]]
    leaks = sum(1 for r in results for p in r["problems"] if p.startswith("lọt"))
    print(f"Eval tầng cổng khung chat: {len(results)} ca, model giả, không gọi mạng")
    for group, n in sorted(by_group.items()):
        bad = sum(1 for r in failed if r["group"] == group)
        print(f"  {group:15} {n - bad}/{n} đạt")
    for r in failed:
        for p in r["problems"]:
            print(f"  {r['id']}: {p}")
    print(f"Op hoặc nội dung lọt: {leaks}. Ca lệch: {len(failed)}.")
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps({"cases": len(results), "leaks": leaks, "failed": len(failed), "results": results},
                              ensure_ascii=False, indent=1), encoding="utf-8")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
