"""Eval tầng model của khung chat (Q135, mốc C): 24 ca x 3 lượt trên `gpt-6-luna`, qua đúng prepare_chat, check_chat.

PM duyệt 12/11 kèm điều kiện: file đã được LibreOffice tính lại; danh sách hàm đã chốt; chỉ gpt-6-luna; trần tổng 0,30 USD,
vượt thì dừng; key từ biến môi trường hoặc .env (eval/_env.py), không in.
Không có --confirm-network thì chỉ dựng nội dung gửi và in chi phí ước tính, không gọi mạng.

Chạy từ `vmai/`:
    .venv/Scripts/python.exe eval/run_chat_model.py                                   # kế hoạch, không mạng
    .venv/Scripts/python.exe eval/run_chat_model.py --profile openai-luna --confirm-network
Kết quả: `out/chat_model_<profile>.json` (không chứa nội dung gửi đi hay câu model trả, chỉ thao tác và mã).

Chấm mỗi lượt: `dung`; `hoi_lai` (không sửa gì, an toàn nhưng chưa làm được); `thieu`; `can_cham` (công thức dạng lạ, người
chấm); `sai_ap_duoc` (thao tác qua cổng mà sai: lỗi nặng nhất, người dùng phải tự thấy ở câu tóm tắt); `loi_goi`.
Thoát mã 1 khi nội dung gửi có chuỗi cấm, hoặc thao tác vào ô bị che qua được cổng. `sai_ap_duoc` là số đo, báo riêng.
"""
from __future__ import annotations

import argparse
import dataclasses
import json
import sys
import time
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _console import utf8  # noqa: E402
from _env import load_key  # noqa: E402
from run_chat_gates import CASES, ROOT, request_from_xlsx  # noqa: E402

from vmai.chat import ChatInputError, check_chat, prepare_chat, to_llm_request  # noqa: E402
from vmai.contract import LlmCall  # noqa: E402
from vmai.llm.base import LLMError, LLMResponse  # noqa: E402
from vmai.llm.budget import Budget, BudgetExceeded  # noqa: E402
from vmai.llm.factory import create_provider, load_profile  # noqa: E402

MAX_TOTAL_USD = 0.30
MAX_TURN_USD = 0.02
RUNS = 3
SAFE = {"ask", "refuse", "unsupported"}


def _norm_formula(text: str) -> str:
    return text.replace(" ", "").replace("$", "").upper().replace("_XLFN.", "")


def _grade(expect: dict, plan, masked_cells: set[str]) -> tuple[str, list[str]]:
    applied = plan.ops
    notes = [f"{o.op} {o.cell}" for o in applied]
    if any(o.cell in masked_cells for o in applied):
        return "lot_o_che", notes
    want = expect.get("ops")
    if not want:
        if applied:
            return "sai_ap_duoc", notes
        return ("dung" if plan.status in expect["status"] else "hoi_lai"), notes
    if not applied:
        return "hoi_lai", notes
    by_cell = {w["cell"]: w for w in want}
    verdict = "dung"
    for op in applied:
        w = by_cell.get(op.cell)
        if w is None or w["op"] != op.op:
            return "sai_ap_duoc", notes
        if op.op == "set_value":
            same = (Decimal(op.value.normalized) == Decimal(w["value"])) if w["type"] == "number" else (op.value.normalized.casefold() == w["value"].casefold())
            if op.value.type != w["type"] or not same:
                return "sai_ap_duoc", notes
        elif op.op == "set_formula" and _norm_formula(op.formula) not in {_norm_formula(x) for x in w["formulas"]}:
            verdict = "can_cham"
    if len({o.cell for o in applied}) < len(by_cell):
        return "thieu", notes
    return verdict, notes


def main() -> int:
    utf8()
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", default="openai-luna")
    parser.add_argument("--confirm-network", action="store_true", help="Cho phép gọi dịch vụ ngoài, tốn phí. Chỉ dữ liệu giả")
    parser.add_argument("--runs", type=int, default=RUNS)
    args = parser.parse_args()

    spec = json.loads(CASES.read_text(encoding="utf-8"))
    model_file = ROOT / spec["model_file"]
    if not model_file.exists():
        print(f"Thiếu file đã tính lại: {spec['model_file']}. Chạy fixtures/eval/recalc_chat_fixtures.py trước.")
        return 1

    prepared_by_case, estimate = {}, 0.0
    for case in spec["cases"]:
        req = request_from_xlsx(model_file, spec["active_sheet"], spec["region"], spec["extra_sheets"], case["message"])
        try:
            prepared = prepare_chat(req)
        except ChatInputError as err:
            prepared_by_case[case["id"]] = err.code
            continue
        for bad in spec["never_send"]:
            if bad in prepared.prompt:
                print(f"{case['id']}: chuỗi cấm trong nội dung gửi. Dừng, không gọi model.")
                return 1
        if prepared.estimated_cost_usd is None or prepared.estimated_cost_usd > MAX_TURN_USD:
            print(f"{case['id']}: chi phí ước tính vượt {MAX_TURN_USD} USD mỗi lượt hoặc chưa có giá. Dừng.")
            return 1
        prepared_by_case[case["id"]] = prepared
        estimate += prepared.estimated_cost_usd * args.runs
    calls = sum(1 for p in prepared_by_case.values() if not isinstance(p, str)) * args.runs
    print(f"Kế hoạch: {len(spec['cases'])} ca x {args.runs} lượt = {calls} lần gọi. "
          f"Chi phí xấu nhất ước tính {estimate:.3f} USD (trần {MAX_TOTAL_USD}).")
    if estimate > MAX_TOTAL_USD:
        print("Ước tính xấu nhất vượt trần tổng. Dừng.")
        return 1
    if not args.confirm_network:
        print("Chưa có --confirm-network: không gọi mạng.")
        return 0

    profile = load_profile(args.profile)
    if profile.model != "gpt-6-luna":
        print("PM chỉ duyệt gpt-6-luna. Dừng.")
        return 1
    if not load_key(profile.api_key_env):
        print("Không có key trong biến môi trường hoặc .env. Dừng.")
        return 1
    price = profile.budget.price
    profile = dataclasses.replace(profile, budget=Budget(max_tokens_per_run=1_000_000, max_usd_per_run=MAX_TOTAL_USD, price=price))
    provider = create_provider(profile)

    rows, totals, stopped = [], Counter(), None
    for case in spec["cases"]:
        prepared = prepared_by_case[case["id"]]
        for run in range(args.runs):
            row = {"id": case["id"], "group": case["group"], "run": run + 1}
            if isinstance(prepared, str):
                row.update(verdict="dung" if not case["model_expect"]["status"] else "hoi_lai", prepare_error=prepared)
                rows.append(row)
                continue
            try:
                started = time.monotonic()
                response: LLMResponse = provider.generate(to_llm_request(prepared))
            except BudgetExceeded:
                stopped = "Vượt trần tổng, dừng."
                break
            except LLMError as err:
                row.update(verdict="loi_goi", error=type(err).__name__)
                rows.append(row)
                continue
            call = LlmCall(purpose="chat_edit", provider=response.provider, model=response.model, request_id=response.request_id,
                           input_tokens=response.input_tokens, output_tokens=response.output_tokens,
                           latency_ms=response.latency_ms or int((time.monotonic() - started) * 1000))
            plan = check_chat(prepared.context, response.text, call, prepared.masked, price=price)
            masked_cells = {a for s in prepared.context.sheets if s.name == prepared.context.active_sheet for a in s.masked}
            verdict, notes = _grade(case["model_expect"], plan, masked_cells)
            totals["in"] += response.input_tokens or 0
            totals["out"] += response.output_tokens or 0
            totals["reasoning"] += response.reasoning_tokens or 0
            totals["usd"] += plan.llm_call.cost_usd or 0
            row.update(verdict=verdict, status=plan.status, applied=notes,
                       rejected=sorted(r.reason for r in plan.rejected_ops), latency_ms=call.latency_ms,
                       input_tokens=response.input_tokens, output_tokens=response.output_tokens)
            rows.append(row)
        if stopped:
            break

    by_group: dict[str, Counter] = defaultdict(Counter)
    for row in rows:
        by_group[row["group"]][row["verdict"]] += 1
    print(f"Model {profile.model}, {time.strftime('%d/%m/%Y %H:%M')}, {sum(1 for r in rows if 'status' in r)} lần gọi")
    for group, counter in sorted(by_group.items()):
        print(f"  {group:15} " + ", ".join(f"{k} {v}" for k, v in sorted(counter.items())))
    overall = Counter(r["verdict"] for r in rows)
    latencies = sorted(r["latency_ms"] for r in rows if "latency_ms" in r)
    print("Tổng: " + ", ".join(f"{k} {v}" for k, v in sorted(overall.items())))
    print(f"Token vào {totals['in']}, ra {totals['out']} (suy luận {totals['reasoning']}). "
          f"Chi phí tính theo bảng giá, không phải hóa đơn: {totals['usd']:.4f} USD. "
          f"Độ trễ trung vị {latencies[len(latencies) // 2] / 1000 if latencies else 0:.1f} giây.")
    if stopped:
        print(stopped)
    out = Path(__file__).resolve().parents[1] / "out" / f"chat_model_{args.profile}.json"
    out.write_text(json.dumps({"model": profile.model, "date": time.strftime("%Y-%m-%d %H:%M"), "runs": args.runs,
                               "totals": dict(totals), "overall": dict(overall), "rows": rows, "stopped": stopped},
                              ensure_ascii=False, indent=1), encoding="utf-8")
    return 1 if overall["lot_o_che"] else 0


if __name__ == "__main__":
    sys.exit(main())
