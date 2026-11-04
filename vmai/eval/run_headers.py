"""Chấm map_headers trên fixtures/eval/header_sets.json. Việc của mức Đóng (Q26).

Không có --confirm-network thì chỉ in kế hoạch. Với profile `fake`, chỉ kiểm dữ liệu đi ra, không phải kết quả chất lượng.
"""
from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path

from vmai.llm.base import LLMBadOutput, LLMError, LLMResponse
from vmai.llm.factory import create_provider, load_profile
from vmai.llm.keys import read_api_key
from vmai.llm_tasks import map_headers, sendable_headers
from vmai.phi import is_person_header, phi_kind

ROOT = Path(__file__).resolve().parents[2]
SETS = ROOT / "fixtures" / "eval" / "header_sets.json"
OUT = ROOT / "vmai" / "out"


def main() -> None:
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
            return
        read_api_key(profile)  # thiếu key thì dừng ở đây, trước khi gọi
        provider = create_provider(profile, allow_unverified=True)

    right = wrong = rejected = failed = 0
    filtered: list[str] = []
    attacks = {"total": 0, "bad": 0}
    latencies: list[int] = []
    tokens = [0, 0]
    details = []
    for item in data["sets"]:
        # Đúng như luồng thật: tiêu đề về người và tiêu đề giống số liệu không lên LLM.
        headers = sendable_headers([h for h in item["headers"] if not is_person_header(h) and not phi_kind(h)])
        filtered += [h for h in item["headers"] if h not in headers]
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
        "median_latency_ms": statistics.median(latencies) if latencies else None,
        "input_tokens": tokens[0],
        "output_tokens": tokens[1],
        "details": details,
    }
    OUT.mkdir(exist_ok=True)
    (OUT / f"headers_{profile.name}.json").write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: v for k, v in result.items() if k != "details"}, ensure_ascii=False))


if __name__ == "__main__":
    main()
