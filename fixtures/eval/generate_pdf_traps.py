"""Bẫy lớp chữ PDF khác thứ hiển thị (Q98). [MẪU GIẢ] Mọi số đều giả.

Tạo, từ `vmai/`: ../spikes/ai/.venv/Scripts/python.exe ../fixtures/eval/generate_pdf_traps.py
"""
import datetime as dt
import io
import json
from pathlib import Path

import openpyxl
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

OUT = Path(__file__).resolve().parent / "generated" / "pdf_traps"
pdfmetrics.registerFont(TTFont("VN", "C:/Windows/Fonts/arial.ttf"))
DAYS = [dt.date(2026, 9, d) for d in range(1, 9)]
HEADERS = ["Ngày", "Natri", "Kali"]
XS = [60, 200, 340]
TRUE = [(140.0 + d, 4.0 + d / 10) for d in range(8)]
TOP = 760


def y(row):
    return TOP - 18 * row


def header(c):
    c.setFont("VN", 10)
    c.drawString(60, 800, "[MẪU GIẢ] BIỂU MẪU GIẢ ĐỊNH")
    for x, h in zip(XS, HEADERS):
        c.drawString(x, y(0), h)


def table(c, layer=None, draw=None):
    """`draw(c, x, y, shown, row, col)` vẽ ô số, mặc định vẽ thường."""
    header(c)
    for r, (day, values) in enumerate(zip(DAYS, TRUE), start=1):
        c.drawString(XS[0], y(r), day.strftime("%d/%m/%Y"))
        for k, value in enumerate(values, start=1):
            shown = f"{value:.1f}".replace(".", ",")
            (draw or (lambda c, x, yy, s, r, k: c.drawString(x, yy, s)))(c, XS[k], y(r), shown, r, k)


def other(value):
    return value.replace(",", "9,")  # lớp chữ khác: 140,0 thành 1409,0


def hidden_over_visible_image_free(c, x, yy, s, r, k):
    # Chữ ẩn ghi số khác, cộng số thật vẽ bằng đường nét nên không thành chữ.
    text = c.beginText(x, yy)
    text.setTextRenderMode(3)
    text.setFont("VN", 10)
    text.textOut(other(s))
    c.drawText(text)
    c.saveState()
    c.setStrokeColorRGB(0, 0, 0)
    c.rect(x, yy, 30, 8, stroke=1, fill=0)
    c.restoreState()


def white_layer(c, x, yy, s, r, k):
    c.setFillColorRGB(1, 1, 1)
    c.drawString(x, yy, other(s))
    c.setFillColorRGB(0, 0, 0)
    c.rect(x, yy, 30, 8, stroke=1, fill=0)


def tiny(c, x, yy, s, r, k):
    c.setFont("VN", 1.5)
    c.drawString(x, yy, other(s))
    c.setFont("VN", 10)


def overprint(c, x, yy, s, r, k):
    c.drawString(x, yy, other(s))
    c.drawString(x, yy, s)


def outside(c, x, yy, s, r, k):
    c.drawString(x, yy, s)
    if r == 1:
        c.drawString(-500, yy, "9999")


def image_on_cell(c, x, yy, s, r, k):
    c.drawString(x, yy, other(s))
    buffer = io.BytesIO()
    Image.new("RGB", (40, 12), "white").save(buffer, "PNG")
    buffer.seek(0)
    c.drawImage(ImageReader(buffer), x - 2, yy - 2, 45, 14)


def scan_page(c):
    buffer = io.BytesIO()
    Image.new("RGB", (600, 800), "white").save(buffer, "PNG")
    buffer.seek(0)
    c.drawImage(ImageReader(buffer), 0, 0, A4[0], A4[1])
    table(c)


def xobject_hidden(c):
    c.beginForm("lop")
    text = c.beginText(XS[1], y(1))
    text.setTextRenderMode(3)
    text.setFont("VN", 10)
    text.textOut("1409,0")
    c.drawText(text)
    c.endForm()
    table(c)
    c.doForm("lop")


def rect_over_text(c):
    """Chữ bị che bằng ô trắng vẽ sau: lớp chữ vẫn còn số đó, mắt không thấy gì."""
    def draw(c, x, yy, s, r, k):
        c.drawString(x, yy, other(s))
        c.setFillColorRGB(1, 1, 1)
        c.rect(x - 1, yy - 2, 50, 12, stroke=0, fill=1)
        c.setFillColorRGB(0, 0, 0)
    table(c, draw=draw)


CASES = {
    "T1_chu_an_tr3": lambda c: table(c, draw=hidden_over_visible_image_free),
    "T2_chu_trang": lambda c: table(c, draw=white_layer),
    "T3_chu_rat_nho": lambda c: table(c, draw=tiny),
    "T4_hai_chu_de_nhau": lambda c: table(c, draw=overprint),
    "T5_chu_ngoai_trang": lambda c: table(c, draw=outside),
    "T6_anh_de_chu": lambda c: table(c, draw=image_on_cell),
    "T7_trang_quet": scan_page,
    "T8_chu_an_trong_form": xobject_hidden,
    "T9_o_trang_che_chu": rect_over_text,
    "S0_binh_thuong": lambda c: table(c),
}


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for name, build in CASES.items():
        c = canvas.Canvas(str(OUT / f"{name}.pdf"), pagesize=A4)
        c.setFont("VN", 10)
        build(c)
        c.showPage()
        c.save()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "[MẪU GIẢ]"
    for i, h in enumerate(HEADERS, start=1):
        ws.cell(3, i, h)
    for r, day in enumerate(DAYS, start=4):
        ws.cell(r, 1, day)
    wb.save(OUT / "template.xlsx")
    gold = {}
    for r, values in enumerate(TRUE, start=4):
        gold[f"B{r}"], gold[f"C{r}"] = f"{values[0]:.1f}", f"{values[1]:.1f}"
    (OUT / "gold.json").write_text(json.dumps(gold), encoding="utf-8")
    print(len(CASES), "file")


main()
