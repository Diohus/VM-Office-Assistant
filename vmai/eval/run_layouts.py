"""Đo infer_spec trên fixtures/eval/generated/layouts (Q53). Không mạng, không LLM.

Bố cục do chính người viết infer_spec tự nghĩ. Đoán đúng ở đây KHÔNG chứng minh được với mẫu thật.
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path

from vmai import MissingFields, SourceFileError, TemplateSpecError, plan_headers, propose, read_form, read_source
from vmai.infer import infer_spec, profile_from_guess

ROOT = Path(__file__).resolve().parents[2]
LAYOUTS = ROOT / "fixtures" / "eval" / "generated" / "layouts"
OUT = ROOT / "vmai" / "out"


def _same(a: str, b: str) -> bool:
    try:
        return Decimal(a) == Decimal(b)
    except InvalidOperation:
        return a == b


def run_case(directory: Path) -> dict:
    answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
    template_path, source_path = directory / "template.xlsx", directory / "source.xlsx"
    source = read_source(source_path)
    template_facts = read_source(template_path)
    guess = infer_spec(template_facts, source)
    # Trường hợp chỉ hỏi về cách ghi M1 và L1: thử lại khi người dùng đồng ý.
    after_answer = None
    if [q.code for q in guess.questions] == ["LEVEL_NAME_DIFFERS"]:
        after_answer = infer_spec(template_facts, source, frozenset({"LEVEL_NAME_DIFFERS"}))
    out = {
        "case": directory.name,
        "expect": answer["expect"],
        "status": guess.status,
        "questions": [q.code for q in guess.questions],
        "gold": len(answer["cells"]),
        "right": 0,
        "wrong": 0,
        "missed": 0,
        "colors": {},
        "wrong_green": 0,
        "confirmed_wrong_green": None,
        "forbidden_leak": [],
        "error": None,
    }
    if after_answer is not None:
        out["one_question"] = True
        guess = after_answer
    if guess.spec is None:
        return classify(out)
    try:
        form = read_form(template_path, guess.spec)
        plan = plan_headers(form, source)
        batch = propose(form, source, plan)
    except (SourceFileError, TemplateSpecError, MissingFields) as exc:
        out["error"] = exc.code
        return classify(out)

    proposed = {}
    colors: Counter[str] = Counter()
    for change in batch.changes:
        a1 = change.target.a1_at_write
        color = change.confidence.final_color.value
        colors[color] += 1
        proposed[a1] = change.value.normalized
        good = a1 in answer["cells"] and _same(change.value.normalized, answer["cells"][a1])
        out["right" if good else "wrong"] += 1
        if color == "green" and not good:
            out["wrong_green"] += 1
    out["colors"] = dict(colors)
    out["missed"] = len(set(answer["cells"]) - set(proposed))
    dumped = batch.model_dump_json()
    out["forbidden_leak"] = [f for f in answer.get("forbidden", []) if f in dumped]

    # Nếu người dùng bấm xác nhận một đoán sai thì sao? Đây là đường số sai có thể thành xanh.
    if not guess.questions:
        profile = profile_from_guess(guess, form.fingerprint, plan.headers, confirmed_by="tester.gia", confirmed_at="2026-11-03T00:00:00+07:00", profile_id="gia")
        confirmed = propose(form, source, plan_headers(form, source, profile))
        bad = 0
        greens = 0
        for change in confirmed.changes:
            if change.confidence.final_color.value == "green":
                greens += 1
                a1 = change.target.a1_at_write
                if a1 not in answer["cells"] or not _same(change.value.normalized, answer["cells"][a1]):
                    bad += 1
        out["confirmed_wrong_green"] = bad
        out["confirmed_green"] = greens
    return classify(out)


def classify(out: dict) -> dict:
    perfect = out["gold"] > 0 and out["right"] == out["gold"] and not out["wrong"] and not out["missed"]
    if out.get("one_question"):
        # Người dùng bị hỏi một câu. Bẫy L24 thì đồng ý là sai, nên đây là chỗ con người phải đọc kỹ.
        out["outcome"] = "hỏi 1 câu rồi đúng" if perfect else ("hỏi lại, đúng" if out["expect"] == "ask" else "hỏi 1 câu rồi vẫn sai")
    elif out["status"] == "ask":
        out["outcome"] = "hỏi lại, đúng" if out["expect"] in ("ask", "ok_or_ask") else "hỏi lại, thừa"
    elif out["expect"] == "ask":
        out["outcome"] = "đoán sai không báo"  # đáng lẽ phải hỏi
    elif perfect:
        out["outcome"] = "đoán đúng"
    else:
        out["outcome"] = "đoán sai không báo"
    return out


def failures(results: list[dict]) -> list[str]:
    """Ca làm eval thoát mã 1. Cùng điều kiện với test_infer: xanh trước khi xác nhận, sai bị xanh sau khi người dùng
    xác nhận đoán, đoán sai không báo, lộ giá trị cấm. Chỉ quyết định mã thoát, không đổi màu."""
    bad = []
    for r in results:
        if r["colors"].get("green", 0):
            bad.append(f"{r['case']}: xanh trước khi xác nhận")
        if r["wrong_green"]:
            bad.append(f"{r['case']}: {r['wrong_green']} sai bị xanh")
        if r["confirmed_wrong_green"]:
            bad.append(f"{r['case']}: {r['confirmed_wrong_green']} sai bị xanh nếu người dùng xác nhận")
        if r["outcome"] in ("đoán sai không báo", "hỏi 1 câu rồi vẫn sai"):
            bad.append(f"{r['case']}: {r['outcome']}")
        if r["forbidden_leak"]:
            bad.append(f"{r['case']}: lộ {r['forbidden_leak']}")
    return bad


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--save-as", default="eval_layouts")
    args = parser.parse_args()
    results = [run_case(d) for d in sorted(LAYOUTS.iterdir()) if d.is_dir()]
    print(f"{'bố cục':28}{'mong':>10}{'kết quả':>22}{'đúng':>6}{'sai':>5}{'sót':>5}  màu; câu hỏi")
    for r in results:
        print(f"{r['case']:28}{r['expect']:>10}{r['outcome']:>22}{r['right']:>6}{r['wrong']:>5}{r['missed']:>5}  {r['colors']} {r['questions']} {r['error'] or ''}")
    outcomes = Counter(r["outcome"] for r in results)
    summary = {
        "layouts": len(results),
        "outcomes": dict(outcomes),
        "green_before_confirm": sum(r["colors"].get("green", 0) for r in results),
        "wrong_green_before_confirm": sum(r["wrong_green"] for r in results),
        "wrong_green_if_user_confirms": sum(r["confirmed_wrong_green"] or 0 for r in results),
        "forbidden_leaks": [r["case"] for r in results if r["forbidden_leak"]],
    }
    print("\nTổng:", json.dumps(summary, ensure_ascii=False))
    print("Bố cục do người viết infer_spec tự nghĩ. CHƯA kiểm chứng với mẫu thật.")
    OUT.mkdir(exist_ok=True)
    (OUT / f"{args.save_as}.json").write_text(json.dumps({"summary": summary, "cases": results}, ensure_ascii=False, indent=1), encoding="utf-8")
    bad = failures(results)
    print("\nCó vấn đề:", bad or "không")
    return 1 if bad else 0


if __name__ == "__main__":
    import _console

    _console.utf8()
    sys.exit(main())
