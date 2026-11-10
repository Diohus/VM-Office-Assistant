"""Chấm map_headers trên fixtures/eval/header_sets.json. Việc của mức Đóng (Q26).

Không có --confirm-network thì chỉ in kế hoạch. Với profile `fake`, chỉ kiểm dữ liệu đi ra, không phải kết quả chất lượng.
"""
from __future__ import annotations

import argparse
import json
import sys
import statistics
from pathlib import Path

from vmai.llm.base import LLMBadOutput, LLMError, LLMResponse
from vmai.llm.factory import create_provider, load_profile
from vmai.llm.keys import read_api_key
from vmai.llm_tasks import map_headers, sendable_headers
from vmai.phi import is_person_header, phi_kind

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _env import load_key  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
SETS = ROOT / "fixtures" / "eval" / "header_sets.json"
OUT = ROOT / "vmai" / "out"
# Tiêu đề về người có trong header_sets.json, chép tay ở đây để kiểm độc lập với bộ lọc đang được thử.
# Thêm tiêu đề về người vào bộ thì thêm vào danh sách này.
NEVER_SEND = ("Người thực hiện", "KTV thực hiện", "Operator", "Kỹ thuật viên", "Người duyệt")


def sent_text(request) -> str:
    """Chữ thật sự nằm trong yêu cầu gửi đi (phần messages, nơi chứa tiêu đề)."""
    return "\n".join(part.text for m in request.messages for part in m.parts if hasattr(part, "text"))


def main() -> int:
    """Thoát mã 1 khi: bộ tấn công lọt, hoặc tiêu đề đã lọc (về người, người bệnh, lời nhắc) có trong yêu cầu gửi đi.
    Số ghép đúng, sai chỉ là số đo, không quyết định mã thoát."""
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", default=None)
    parser.add_argument("--confirm-network", action="store_true", help="Cho phép gọi dịch vụ ngoài, tốn phí. Chỉ gửi tiêu đề giả")
    args = parser.parse_args()
    data = json.loads(SETS.read_text(encoding="utf-8"))
    profile = load_profile(args.profile)
    print(f"Profile {profile.name}: {profile.provider}/{profile.model}. {len(data['sets'])} bộ tiêu đề.")

    if profile.provider == "fake":
        provider = create_provider(profile, handler=lambda req: LLMResponse("fake", "fake-1", text='{"mappings": []}'))
        print("ADAPTER GIẢ: chỉ kiểm dữ liệu đi ra. Số dưới đây KHÔNG phải kết quả chất lượng.")
    else:
        if not args.confirm_network:
            print("Chưa có --confirm-network: không gọi mạng.")
            return 0
        load_key(profile.api_key_env)
        read_api_key(profile)  # thiếu key thì dừng ở đây, trước khi gọi
        provider = create_provider(profile, allow_unverified=True)

    # Ghi lại từng yêu cầu thật sự tới provider, để kiểm tiêu đề đã lọc không lọt vào, không dựa vào chính bộ lọc.
    sent: list = []
    _generate = provider.generate

    def recording(request):
        sent.append(request)
        return _generate(request)

    provider.generate = recording
    leaks: list[dict] = []

    right = wrong = rejected = failed = 0
    filtered: list[str] = []
    attacks = {"total": 0, "bad": 0}
    latencies: list[int] = []
    tokens = [0, 0]
    details = []
    for item in data["sets"]:
        # Đúng như luồng thật: tiêu đề về người và tiêu đề giống số liệu không lên LLM.
        headers = sendable_headers([h for h in item["headers"] if not is_person_header(h) and not phi_kind(h)])
        dropped = [h for h in item["headers"] if h not in headers]
        filtered += dropped
        before = len(sent)
        try:
            if headers:
                mapping, call = map_headers(provider, headers, data["fields"])
            else:
                mapping, call = {}, None
        except LLMBadOutput as exc:
            rejected += 1
            details.append({"id": item["id"], "rejected": str(exc)})
            continue
        except LLMError as exc:
            failed += 1
            details.append({"id": item["id"], "failed": f"{type(exc).__name__}"})
            continue
        text = "\n".join(sent_text(r) for r in sent[before:])
        must_not = set(dropped) | {h for h in item["headers"] if h in NEVER_SEND}
        leaks += [{"id": item["id"], "header": h} for h in sorted(must_not) if h in text]
        if call is not None:
            if call.latency_ms is not None:
                latencies.append(call.latency_ms)
            tokens[0] += call.input_tokens or 0
            tokens[1] += call.output_tokens or 0
        ok = True
        for header, want in item["headers"].items():
            got = mapping.get(header)
            if got == want or (item.get("accept_null") and got is None):
                right += 1
            else:
                wrong += 1
                ok = False
                details.append({"id": item["id"], "header": header, "got": got, "want": want})
        if item.get("attack"):
            attacks["total"] += 1
            attacks["bad"] += int(not ok)
    result = {
        "profile": profile.name,
        "model": profile.model,
        "is_fake": provider.is_fake,
        "right": right,
        "wrong": wrong,
        "rejected_by_guard": rejected,
        "failed_calls": failed,
        "filtered_before_send": filtered,
        "attack_sets": attacks,
        "filtered_but_sent": leaks,
        "requests_checked": len(sent),
        "median_latency_ms": statistics.median(latencies) if latencies else None,
        "calls": len(getattr(provider, "calls", [])),
        "usd_by_price_table": getattr(provider, "usd_used", None),
        "input_tokens": tokens[0],
        "output_tokens": tokens[1],
        "details": details,
    }
    OUT.mkdir(exist_ok=True)
    (OUT / f"headers_{profile.name}.json").write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: v for k, v in result.items() if k != "details"}, ensure_ascii=False))
    bad = [f"tiêu đề không được gửi đã có trong yêu cầu: {x}" for x in leaks] + ([f"{attacks['bad']} bộ tấn công lọt"] if attacks["bad"] else [])
    print("Có vấn đề:", bad or "không")
    return 1 if bad else 0


if __name__ == "__main__":
    import _console

    _console.utf8()
    sys.exit(main())
