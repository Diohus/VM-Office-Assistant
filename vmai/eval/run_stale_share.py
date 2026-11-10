"""Đo đề xuất "nghi nhầm file cũ" của PO/BA (T-AI-13, chưa là quy tắc): tỷ lệ ô giống hệt tháng trước, cùng vị trí.

CHỈ ĐO, không gọi model, không đổi màu. Dữ liệu giả sinh bằng seed cố định, mỗi kiểu cột 2.000 cặp tháng liền nhau.
Báo nhầm: nguồn đúng của tháng này mà tỷ lệ giống tháng trước vượt ngưỡng. Bắt được: nạp lại file tháng trước, có thể đã sửa vài ô.

    .venv/Scripts/python.exe eval/run_stale_share.py
"""
from __future__ import annotations

import json
import math
import random
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "vmai" / "out"
PAIRS = 2000
CELLS = 30  # một cột, 30 ngày
THRESHOLDS = (0.3, 0.5, 0.7, 0.9)


def _poisson(rng: random.Random, lam: float) -> int:
    limit, k, p = math.exp(-lam), 0, 1.0
    while True:
        p *= rng.random()
        if p <= limit:
            return k
        k += 1


# Kiểu cột giả: tên -> hàm sinh một tháng (30 số). Ghi chú là lý do có trong bộ.
KINDS = {
    "qc_2_so_le": ("QC liên tục, 2 chữ số thập phân, như GIA-QC-01", lambda r: [round(r.gauss(5.5, 0.1), 2) for _ in range(CELLS)]),
    "qc_1_so_le_it_bien_dong": ("Kali 4,0 ± 0,05 làm tròn 1 chữ số: ít giá trị khác nhau", lambda r: [round(r.gauss(4.0, 0.05), 1) for _ in range(CELLS)]),
    "qc_lam_tron_so_nguyen": ("Natri 140 ± 1 làm tròn số nguyên", lambda r: [round(r.gauss(140, 1)) for _ in range(CELLS)]),
    "dem_it_lambda_1": ("Số đếm tiêu hao, trung bình 1/ngày: nhiều số 0, 1", lambda r: [_poisson(r, 1) for _ in range(CELLS)]),
    "dem_lambda_5": ("Số đếm, trung bình 5/ngày", lambda r: [_poisson(r, 5) for _ in range(CELLS)]),
    "dem_lambda_30": ("Số đếm, trung bình 30/ngày", lambda r: [_poisson(r, 30) for _ in range(CELLS)]),
    "co_dinh_dinh_muc": ("Cột gần như cố định (định mức, số ca): 90% ngày giữ nguyên", lambda r: [8 if r.random() < 0.9 else 7 for _ in range(CELLS)]),
}


def share_same(a: list, b: list) -> float:
    return sum(x == y for x, y in zip(a, b)) / len(a)


def main() -> None:
    rows = {}
    print("CHỈ ĐO, dữ liệu giả, chưa có quy tắc nào dùng số này. Không có số thật.\n")
    print(f"{'kiểu cột':28}{'tb giống':>9}{'p99':>7}{'max':>7}  " + "  ".join(f"báo nhầm>{t:.0%}" for t in THRESHOLDS))
    for name, (note, gen) in KINDS.items():
        rng = random.Random(zlib.crc32(name.encode()))
        shares = sorted(share_same(gen(rng), gen(rng)) for _ in range(PAIRS))
        false_alarm = {t: sum(s > t for s in shares) / PAIRS for t in THRESHOLDS}
        rows[name] = {"note": note, "mean": sum(shares) / PAIRS, "p99": shares[int(PAIRS * 0.99) - 1], "max": shares[-1], "false_alarm": false_alarm}
        print(f"{name:28}{rows[name]['mean']:>9.2f}{rows[name]['p99']:>7.2f}{rows[name]['max']:>7.2f}  " + "  ".join(f"{false_alarm[t]:>12.1%}" for t in THRESHOLDS))

    # Bắt được: nạp lại file tháng trước, người dùng đã sửa tay k ô trên 30.
    print("\nNạp lại file tháng trước, đã sửa tay vài ô (mọi kiểu cột như nhau, vì so cùng vị trí):")
    caught = {}
    for edited in (0, 3, 9, 15, 21):
        share = (CELLS - edited) / CELLS
        caught[edited] = {f"{t:.0%}": share > t for t in THRESHOLDS}
        print(f"  sửa {edited:>2}/30 ô: giống {share:.0%} -> " + ", ".join(f">{k}: {'bắt' if v else 'sót'}" for k, v in caught[edited].items()))
    # Cách khác: đếm DÒNG giống hệt (mọi cột của một ngày cùng bằng tháng trước), bỏ cột gần như cố định.
    # Biểu mẫu giả xấu nhất: 6 cột ít biến động (kiểu Kali làm tròn 1 chữ số) + 1 cột cố định bị bỏ ra.
    print("\nĐếm theo dòng: 6 cột ít biến động, bỏ cột cố định. Báo khi số dòng giống hệt vượt ngưỡng:")
    rng = random.Random(77)
    gen = KINDS["qc_1_so_le_it_bien_dong"][1]
    row_shares = []
    for _ in range(PAIRS):
        a, b = [gen(rng) for _ in range(6)], [gen(rng) for _ in range(6)]
        row_shares.append(sum(all(a[c][d] == b[c][d] for c in range(6)) for d in range(CELLS)) / CELLS)
    row_alarm = {f"{t:.0%}": sum(s > t for s in row_shares) / PAIRS for t in (0.1, 0.3, 0.5)}
    print(f"  nguồn đúng: tb {sum(row_shares) / PAIRS:.3f}, max {max(row_shares):.2f}; báo nhầm " + ", ".join(f">{k}: {v:.2%}" for k, v in row_alarm.items()))
    row_caught = {}
    for edited_rows in (0, 9, 15, 21, 27):
        share = (CELLS - edited_rows) / CELLS
        row_caught[edited_rows] = {k: share > float(k.rstrip("%")) / 100 for k in row_alarm}
        print(f"  nạp lại file cũ, sửa {edited_rows:>2}/30 dòng: dòng giống {share:.0%} -> " + ", ".join(f">{k}: {'bắt' if v else 'sót'}" for k, v in row_caught[edited_rows].items()))
    print("\nLưu ý: file tháng trước còn ngày của tháng trước thì đã bị chặn bởi kiểm tháng (SOURCE_WRONG_MONTH, X13), trước cả phép đo này.")
    print("CHỈ ĐO: lệnh này luôn thoát mã 0. Đọc số ở trên.")
    short = short_batches()
    per_row = per_row_option()
    OUT.mkdir(exist_ok=True)
    (OUT / "eval_stale_share.json").write_text(json.dumps({"pairs": PAIRS, "cells": CELLS, "kinds": rows, "reloaded": caught, "row_alarm": row_alarm, "row_caught": row_caught, "short_batches": short, "per_row_option": per_row}, ensure_ascii=False, indent=1), encoding="utf-8")


def per_row_option() -> dict:
    """Đề xuất của PO/BA sau Q120 (CHƯA là quy tắc): dòng nào trùng hệt tháng trước thì chính dòng đó tối đa vàng, không ngưỡng.

    V1: chỉ khi biểu mẫu có ≥ 3 cột thay đổi (như quy tắc hiện tại). V2: cả khi chỉ 1-2 cột thay đổi.
    Báo nhầm = tỷ lệ dòng của nguồn ĐÚNG bị hạ vàng. Bắt = số dòng cũ còn xanh sau khi nạp lại file cũ đã sửa.
    """
    from collections import Counter

    forms = {
        "6_cot_qc_2_so_le": (6, lambda r, n: [round(r.gauss(5.5, 0.1), 2) for _ in range(n)]),
        "6_cot_it_bien_dong": (6, lambda r, n: [round(r.gauss(4.0, 0.05), 1) for _ in range(n)]),
        "3_cot_2_gia_tri": (3, lambda r, n: [r.choice((0, 0, 1)) for _ in range(n)]),
        "4_cot_2_gia_tri": (4, lambda r, n: [r.choice((0, 0, 1)) for _ in range(n)]),
        "2_cot_it_bien_dong": (2, lambda r, n: [round(r.gauss(4.0, 0.05), 1) for _ in range(n)]),
        "1_cot_dem_lambda_1": (1, lambda r, n: [_poisson(r, 1) for _ in range(n)]),
    }
    out: dict = {}
    print("\nHạ vàng từng dòng trùng (Q121). Nguồn ĐÚNG: % dòng bị hạ vàng nhầm, V1 (chỉ khi ≥ 3 cột thay đổi) | V2 (mọi số cột, Q121 chọn):")
    print(f"  {'biểu mẫu giả':22}" + "".join(f"{f'{n} ngày':>18}" for n in (3, 30)))
    for name, (ncols, gen) in forms.items():
        rng = random.Random(zlib.crc32(("dong" + name).encode()))
        cells = []
        for days in (3, 30):
            v1 = v2 = 0
            for _ in range(PAIRS):
                this, last = [gen(rng, days) for _ in range(ncols)], [gen(rng, days) for _ in range(ncols)]
                moving = [c for c in range(ncols) if Counter(last[c]).most_common(1)[0][1] < 0.8 * days]
                same = sum(all(this[c][d] == last[c][d] for c in moving) for d in range(days)) if moving else 0
                v2 += same
                v1 += same if len(moving) > 2 else 0
            out.setdefault(name, {})[days] = {"v1": v1 / (PAIRS * days), "v2": v2 / (PAIRS * days)}
            cells.append(f"{v1 / (PAIRS * days):>8.1%} | {v2 / (PAIRS * days):<7.1%}")
        print(f"  {name:22}" + "".join(f"{c:>18}" for c in cells))
    print("  Nạp lại file cũ đã sửa k dòng: mọi dòng không sửa đều trùng, nên cả V1 và V2 hạ vàng hết các dòng đó: 0 ô xanh mang số cũ,")
    print("  trừ V1 khi biểu mẫu chỉ có 1-2 cột thay đổi (không kiểm, ô cũ vẫn xanh).")
    return out


def _rule(this: list[list], last: list[list], min_same_days: int) -> bool:
    """Chép đúng quy tắc trong propose._same_as_last_month (Q119), thêm tham số "ít nhất bao nhiêu ngày trùng" để so phương án."""
    from collections import Counter

    moving = [c for c in range(len(last)) if Counter(last[c]).most_common(1)[0][1] < 0.8 * len(last[c])]
    if len(moving) <= 2:
        return False
    days = len(this[0])
    same = sum(all(this[c][d] == last[c][d] for c in moving) for d in range(days))
    return same / days > 0.30 and same >= min_same_days


def short_batches() -> dict:
    """T-AI-15: báo nhầm ở lô ngắn với nguồn đúng, so quy tắc hiện tại (≥1 ngày trùng) với phương án cần ≥2 ngày trùng."""
    forms = {
        "6_cot_qc_2_so_le": (6, lambda r, n: [round(r.gauss(5.5, 0.1), 2) for _ in range(n)]),
        "6_cot_it_bien_dong": (6, lambda r, n: [round(r.gauss(4.0, 0.05), 1) for _ in range(n)]),
        "3_cot_2_gia_tri": (3, lambda r, n: [r.choice((0, 0, 1)) for _ in range(n)]),
        "4_cot_2_gia_tri": (4, lambda r, n: [r.choice((0, 0, 1)) for _ in range(n)]),
    }
    out: dict = {}
    print("\nLô ngắn, nguồn ĐÚNG (báo nhầm), 2.000 cặp tháng mỗi ô bảng. Quy tắc hiện tại | phương án cần ≥ 2 ngày trùng:")
    print(f"  {'biểu mẫu giả':22}" + "".join(f"{f'{n} ngày':>18}" for n in (3, 5, 10, 30)))
    for name, (ncols, gen) in forms.items():
        rng = random.Random(zlib.crc32(name.encode()))
        cells = []
        for days in (3, 5, 10, 30):
            now = then = 0
            for _ in range(PAIRS):
                this, last = [gen(rng, days) for _ in range(ncols)], [gen(rng, days) for _ in range(ncols)]
                now += _rule(this, last, 1)
                then += _rule(this, last, 2)
            out.setdefault(name, {})[days] = {"current": now / PAIRS, "min_2_days": then / PAIRS}
            cells.append(f"{now / PAIRS:>8.1%} | {then / PAIRS:<7.1%}")
        print(f"  {name:22}" + "".join(f"{c:>18}" for c in cells))
    # Bắt được ở lô ngắn: nạp lại file cũ đúng số ngày, đã sửa k ngày (với cột thay đổi, mọi ngày không sửa đều trùng).
    print("\nLô ngắn, nạp lại file cũ đã sửa k ngày: bắt hay sót (hiện tại | cần ≥ 2 ngày):")
    caught = {}
    for days in (3, 5):
        for edited in range(days):
            same = days - edited
            now, then = same / days > 0.30, same / days > 0.30 and same >= 2
            caught[f"{days}_ngay_sua_{edited}"] = {"current": now, "min_2_days": then}
            print(f"  lô {days} ngày, sửa {edited}: {'bắt' if now else 'sót'} | {'bắt' if then else 'sót'}")
    out["reload_caught"] = caught
    return out


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
