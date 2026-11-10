"""Bộ ảnh Q90: chưa dùng để chỉnh. Seed khác, 4 bố cục khác bộ cũ, có tiêu đề hai tầng và bảng toàn số nguyên.

Tạo, từ `spikes/ai`: .venv/Scripts/python.exe ../../fixtures/eval/generate_images_q90.py

[MẪU GIẢ] Mọi số, mã, tên đều giả, do seed sinh ra. Ảnh là mô phỏng bằng code, chưa phải ảnh chụp thật.
Đáp án lấy từ dữ liệu sinh ra. `manifest.json` ghi mã băm đáp án, tạo trước khi chạy đọc ảnh, để kiểm đáp án không đổi sau đó.
Người viết bộ này cũng là người viết phần đọc ảnh, nên đây chưa phải bộ độc lập hoàn toàn.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import random
import shutil
from pathlib import Path

import openpyxl
from generate import degrade, screen_photo, teams_copy
from openpyxl.utils import get_column_letter as L
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "generated" / "images_q90"
FONT = Path("C:/Windows/Fonts/arial.ttf")
SEEDS = (601, 602)
VARIANTS = ("clean", "phone", "skew5", "rot90", "rot180", "dark", "shadow", "lowres", "screen", "teams")
DISCLAIMER = "[MẪU GIẢ] Dữ liệu giả của đội. Ảnh mô phỏng, chưa phải ảnh chụp thật."
TITLE = "[MẪU GIẢ] BIỂU MẪU GIẢ ĐỊNH"
DAYS = [dt.date(2026, 10, 1) + dt.timedelta(days=i) for i in range(31)]


def _dmy(day: dt.date) -> str:
    return day.strftime("%d/%m/%Y")


def _decimal(value: float, places: int, comma: bool) -> str:
    text = f"{value:.{places}f}"
    return text.replace(".", ",") if comma else text


def draw_table(rows: list[list[str]], xs: list[int], *, size=(2200, 1584), font_size=30, row_h=40, grid=True, top=100, tiers: list[tuple[str, int, int]] | None = None, footer=()) -> Image.Image:
    """`tiers`: tầng tiêu đề trên, gồm (chữ, cột đầu, cột cuối), vẽ gộp ngang. `xs` có thêm mép phải của bảng."""
    image = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(str(FONT), font_size)
    draw.text((xs[0], 30), f"{TITLE} THÁNG 10/2026", font=ImageFont.truetype(str(FONT), font_size + 8), fill="black")
    y = top
    if tiers:
        for text, first, last in tiers:
            width = draw.textlength(text, font=font)
            draw.text(((xs[first] + xs[last + 1] - width) / 2, y + 6), text, font=font, fill="black")
            if grid:
                draw.rectangle((xs[first], y, xs[last + 1], y + row_h), outline="black", width=2)
        y += row_h
    for r, row in enumerate(rows):
        cy = y + row_h * r
        for c, text in enumerate(row):
            draw.text((xs[c] + 8, cy + 6), text, font=font, fill="black")
        if grid:
            draw.line((xs[0], cy + row_h, xs[-1], cy + row_h), fill="black", width=1)
    bottom = y + row_h * len(rows)
    if grid:
        draw.line((xs[0], top, xs[-1], top), fill="black", width=2)
        for x in xs:
            draw.line((x, top if tiers is None else y, x, bottom), fill="black", width=2)
    for i, line in enumerate(footer):
        draw.text((xs[0], bottom + 16 + i * (row_h + 4)), line, font=font, fill="black")
    return image


def _template(rows: list[list], merges=(), first_data_row=4):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = TITLE
    for r, row in enumerate(rows, start=3):
        for c, value in enumerate(row, start=1):
            if value is not None:
                ws.cell(r, c, value).number_format = "dd/mm/yyyy" if isinstance(value, dt.date) else "General"
    for span in merges:
        ws.merge_cells(span)
    wb.properties.title = "[MẪU GIẢ]"
    return wb


def supplies_integers(seed: int):
    """Bảng toàn số nguyên, khóa là mã chữ, không có ngày. Seed chẵn in đơn giá có dấu chấm ngăn nghìn."""
    rng = random.Random(seed)
    items = [(f"VT{i:03d}", f"Vật tư giả {i}", rng.choice(["Cái", "Hộp", "Lọ", "Đôi"]), rng.randint(1, 480), rng.randint(2, 95) * 1000) for i in range(1, 25)]
    thousands = seed % 2 == 0
    rows = [["Mã VT", "Tên vật tư", "ĐVT", "Số lượng", "Đơn giá"]]
    rows += [[code, name, unit, str(qty), f"{price:,}".replace(",", ".") if thousands else str(price)] for code, name, unit, qty, price in items]
    image = draw_table(rows, [60, 300, 900, 1120, 1420, 1800])
    template = _template([["Mã VT", "Tên vật tư", "ĐVT", "Số lượng", "Đơn giá"], *[[code, name, unit] for code, name, unit, _, _ in items]])
    cells = []
    for r, (_, _, _, qty, price) in enumerate(items, start=4):
        cells += [{"a1": f"D{r}", "value": str(qty)}, {"a1": f"E{r}", "value": str(price)}]
    return image, template, cells, {"integer_only": True, "thousands_dot": thousands}


def two_tier(seed: int):
    """Tiêu đề hai tầng: tên xét nghiệm gộp trên, mức ở dưới."""
    rng = random.Random(seed)
    tests = [("Glucose", 5.5, 0.15, 16.5, 0.4), ("Cholesterol", 3.0, 0.08, 6.5, 0.15), ("Creatinin", 80.0, 3.0, 400.0, 10.0)]
    comma = seed % 2 == 1
    days = DAYS[:25]
    values = {d: [round(rng.gauss(mean, sd * 0.7), 2) for _, a, b, c, e in tests for mean, sd in ((a, b), (c, e))] for d in days}
    rows = [["Ngày", *["Mức 1", "Mức 2"] * 3]]
    rows += [[_dmy(d), *[_decimal(v, 2, comma) for v in values[d]]] for d in days]
    xs = [60, 320, 560, 800, 1040, 1280, 1520, 1760]
    tiers = [("Glucose", 1, 2), ("Cholesterol", 3, 4), ("Creatinin", 5, 6)]
    image = draw_table(rows, xs, tiers=tiers)
    template = _template(
        [["Ngày", "Glucose", None, "Cholesterol", None, "Creatinin", None], [None, *["Mức 1", "Mức 2"] * 3], *[[d] for d in days]],
        merges=("A3:A4", "B3:C3", "D3:E3", "F3:G3"),
    )
    cells = [{"a1": f"{L(c)}{r}", "value": f"{v:.2f}"} for r, d in enumerate(days, start=5) for c, v in enumerate(values[d], start=2)]
    return image, template, cells, {"two_tier": True}


def no_grid_reordered(seed: int):
    """Không kẻ bảng, có cột STT, biểu mẫu xếp cột khác thứ tự trên ảnh."""
    rng = random.Random(seed)
    days = DAYS[:20]
    values = {d: {"Natri": round(rng.gauss(140, 2), 1), "Kali": round(rng.gauss(4.2, 0.2), 1), "Clo": round(rng.gauss(102, 2), 1)} for d in days}
    comma = seed % 2 == 0
    rows = [["STT", "Ngày XN", "Natri", "Kali", "Clo"]]
    rows += [[str(i), _dmy(d), *[_decimal(values[d][k], 1, comma) for k in ("Natri", "Kali", "Clo")]] for i, d in enumerate(days, start=1)]
    image = draw_table(rows, [80, 260, 640, 960, 1280, 1600], size=(1800, 1200), grid=False, row_h=46)
    template = _template([["Ngày XN", "Kali", "Natri", "Clo"], *[[d] for d in days]])
    cells = [{"a1": f"{column}{r}", "value": f"{values[d][key]:.1f}"} for r, d in enumerate(days, start=4) for column, key in (("B", "Kali"), ("C", "Natri"), ("D", "Clo"))]
    return image, template, cells, {"no_grid": True, "reordered": True}


def portrait_units_footer(seed: int):
    """Khổ dọc, chữ nhỏ, tiêu đề ảnh có đơn vị còn biểu mẫu thì không, cuối bảng có dòng trung bình và ghi chú."""
    rng = random.Random(seed)
    days = DAYS[:31]
    values = {d: (rng.randint(120, 160), round(rng.gauss(7.0, 0.5), 1), rng.randint(180, 320)) for d in days}
    comma = seed % 2 == 1
    rows = [["Ngày", "HGB (g/L)", "WBC (G/L)", "PLT (G/L)"]]
    rows += [[_dmy(d), str(a), _decimal(b, 1, comma), str(c)] for d, (a, b, c) in values.items()]
    average = [sum(v[i] for v in values.values()) / len(values) for i in range(3)]
    rows.append(["Trung bình", f"{average[0]:.0f}", _decimal(average[1], 1, comma), f"{average[2]:.0f}"])
    image = draw_table(rows, [60, 340, 620, 900, 1180], size=(1240, 1754), font_size=24, row_h=34, footer=("Ghi chú: dữ liệu giả định",))
    template = _template([["Ngày", "HGB", "WBC", "PLT"], *[[d] for d in days]])
    cells = []
    for r, d in enumerate(days, start=4):
        a, b, c = values[d]
        cells += [{"a1": f"B{r}", "value": str(a)}, {"a1": f"C{r}", "value": f"{b:.1f}"}, {"a1": f"D{r}", "value": str(c)}]
    return image, template, cells, {"portrait": True, "units_in_image_only": True, "summary_row": True}


LAYOUTS = {"B1_vattu_songuyen": supplies_integers, "B2_hai_tang": two_tier, "B3_khong_ke": no_grid_reordered, "B4_doc_donvi": portrait_units_footer}


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    manifest = {}
    for layout, build in LAYOUTS.items():
        for seed in SEEDS:
            page, template, cells, traits = build(seed)
            for variant in VARIANTS:
                directory = OUT / f"{layout}_{variant}_s{seed}"
                directory.mkdir(parents=True)
                quality = 85
                if variant == "screen":
                    out = screen_photo(page, seed)
                elif variant == "teams":
                    out, quality = teams_copy(degrade(page, "phone", seed))
                else:
                    out = degrade(page, variant, seed)
                out.save(directory / "page.jpg", quality=quality)
                template.save(directory / "template.xlsx")
                body = json.dumps({"disclaimer": DISCLAIMER, "layout": layout, "variant": variant, "seed": seed, "simulated": True, "traits": traits, "cells": cells}, ensure_ascii=False, indent=1)
                (directory / "answer.json").write_text(body, encoding="utf-8")
                manifest[directory.name] = hashlib.sha256(body.encode("utf-8")).hexdigest()
    (OUT / "manifest.json").write_text(json.dumps(manifest, indent=0), encoding="utf-8")
    print(len(manifest), "ảnh;", "mã băm cả bộ:", hashlib.sha256(json.dumps(manifest, sort_keys=True).encode()).hexdigest())


if __name__ == "__main__":
    main()
