"""Tài liệu hãng giả để thử Word đầu mục (Q54). [MẪU GIẢ] Hãng, máy và thông số đều giả.

Chạy bằng môi trường có reportlab: spikes/ai/.venv/Scripts/python.exe fixtures/eval/generate_word.py
Đáp án lấy từ dữ liệu sinh ra. Bố cục tự nghĩ, không chứng minh được với tài liệu hãng thật.
"""
from __future__ import annotations

import json
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.pdfencrypt import StandardEncryption
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

OUT = Path(__file__).resolve().parent / "generated" / "word"
FONT_DIR = Path("C:/Windows/Fonts")
pdfmetrics.registerFont(TTFont("Body", str(FONT_DIR / "arial.ttf")))
pdfmetrics.registerFont(TTFont("BodyBold", str(FONT_DIR / "arialbd.ttf")))
WIDTH, HEIGHT = A4
PARA = (
    "Máy phân tích giả định cần được đặt trên mặt phẳng, cách tường ít nhất 25 cm và tránh ánh nắng trực tiếp. "
    "Nhiệt độ phòng duy trì trong khoảng 18 đến 25 độ C, độ ẩm dưới 80 phần trăm. Điện áp nguồn 220 V, tần số 50 Hz."
)
PARA2 = (
    "Trước khi chạy mẫu, kiểm tra mức hóa chất và thùng thải. Thể tích hút mẫu 5 mL, thời gian ủ 30 phút ở 37 độ C. "
    "Không dùng hóa chất đã quá hạn. The analyzer must be calibrated every 90 days or after replacing the lamp."
)
STEPS = ["1. Mở nắp máy", "2. Đặt khay mẫu vào vị trí", "3. Đóng nắp và bấm nút chạy"]

SECTIONS = [
    (1, "Giới thiệu", [PARA]),
    (2, "Phạm vi sử dụng", [PARA2]),
    (2, "An toàn", [PARA, "STEPS"]),
    (1, "Lắp đặt", [PARA2]),
    (2, "Yêu cầu vị trí", [PARA, "TABLE"]),
    (3, "Yêu cầu điện", [PARA2]),
    (2, "Kết nối", [PARA]),
    (1, "Vận hành", [PARA2, "STEPS"]),
    (2, "Hiệu chuẩn", [PARA]),
    (2, "Bảo trì", [PARA2]),
]
# Lặp để tài liệu đủ dài như tài liệu hãng, đầu trang và chân trang lặp qua nhiều trang.
SECTIONS = SECTIONS + [(level, f"{title} phần {n}", parts) for n in (2, 3, 4) for level, title, parts in SECTIONS]
TABLE = [["Thông số", "Giá trị", "Đơn vị"], ["Điện áp", "220", "V"], ["Tần số", "50", "Hz"], ["Công suất", "350", "W"]]
PHI_LINE = "Họ tên: Nguyễn Giả Một, Mã BN: GIA0001"


@dataclass
class Style:
    numbered: bool = True  # tiêu đề có sẵn số
    size_by_level: tuple[int, int, int] = (16, 13, 11)
    bold: bool = True
    header: bool = True
    toc: bool = False
    two_column: bool = False
    extra_body: list[str] = field(default_factory=list)


def wrap(text: str, font: str, size: float, width: float) -> list[str]:
    lines, current = [], ""
    for word in text.split():
        trial = f"{current} {word}".strip()
        if stringWidth(trial, font, size) <= width:
            current = trial
        else:
            lines.append(current)
            current = word
    return [*lines, current] if current else lines


class Writer:
    def __init__(self, c: canvas.Canvas, style: Style):
        self.c, self.style = c, style
        self.page = 1
        self.column = 0
        self.left = 50
        self.width = (WIDTH - 130) / 2 if style.two_column else WIDTH - 100
        self.y = HEIGHT - 70
        self.rendered: list[str] = []  # đúng từng dòng đã in, làm đáp án cho OCR
        self.frame()

    def frame(self):
        if self.style.header:
            self.c.setFont("Body", 8)
            self.c.drawString(50, HEIGHT - 30, "[MẪU GIẢ] Hãng Giả Định - Hướng dẫn sử dụng Máy Giả A")
            self.c.drawString(50, 25, f"Trang {self.page}")

    def newline(self, height):
        self.y -= height
        if self.y < 60:
            if self.style.two_column and self.column == 0:
                self.column, self.left, self.y = 1, 50 + self.width + 30, HEIGHT - 70
                return
            self.c.showPage()
            self.page += 1
            self.column, self.left, self.y = 0, 50, HEIGHT - 70
            self.frame()

    def text(self, text, font="Body", size=10):
        for line in wrap(text, font, size, self.width):
            self.c.setFont(font, size)
            self.c.drawString(self.left, self.y, line)
            self.rendered.append(line)
            self.newline(size * 1.35)

    def table(self):
        rows, cw, rh = TABLE, self.width / 3, 16
        top = self.y
        for r, row in enumerate(rows):
            for i, cell in enumerate(row):
                self.c.setFont("Body", 9)
                self.c.drawString(self.left + i * cw + 3, top - r * rh - 11, cell)
            self.rendered.append(" ".join(row))
        for r in range(len(rows) + 1):
            self.c.line(self.left, top - r * rh, self.left + 3 * cw, top - r * rh)
        for i in range(4):
            self.c.line(self.left + i * cw, top, self.left + i * cw, top - len(rows) * rh)
        self.newline(len(rows) * rh + 10)


def number_of(counters):
    return ".".join(str(c) for c in counters if c)


def write_doc(path: Path, style: Style, encrypt=None) -> dict:
    c = canvas.Canvas(str(path), pagesize=A4, encrypt=encrypt)
    c.setTitle("[MẪU GIẢ] Hướng dẫn sử dụng giả")
    w = Writer(c, style)
    headings, body, counters = [], [], [0, 0, 0]
    w.text("[MẪU GIẢ] HƯỚNG DẪN SỬ DỤNG MÁY GIẢ A", "Body", 10)
    body.append("[MẪU GIẢ] HƯỚNG DẪN SỬ DỤNG MÁY GIẢ A")
    if style.toc:
        w.text("Mục lục")
        body.append("Mục lục")
        for i, (level, title, _) in enumerate(SECTIONS[:5], start=1):
            line = f"{title} ........................ {i + 1}"
            w.text(line)
            body.append(line)
    for extra in style.extra_body:
        w.text(extra)
        body.append(extra)
    for level, title, parts in SECTIONS:
        counters[level - 1] += 1
        for i in range(level, 3):
            counters[i] = 0
        text = f"{number_of(counters[:level])} {title}" if style.numbered else title
        size = style.size_by_level[level - 1]
        w.newline(6)
        w.text(text, "BodyBold" if style.bold else "Body", size)
        headings.append({"text": text, "level": level})
        for part in parts:
            if part == "TABLE":
                if not style.two_column:
                    w.table()
            elif part == "STEPS":
                for step in STEPS:
                    w.text(step)
                    body.append(step)
            else:
                w.text(part)
                body.append(part)
    c.save()
    return {"headings": headings, "body": body, "has_table": not style.two_column, "pages": w.page, "rendered_lines": w.rendered}


def main() -> None:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    cases = {
        "D01_so_va_dam": (Style(), None, "Có số, chữ đậm, cỡ chữ theo cấp"),
        "D02_khong_so": (Style(numbered=False), None, "Không đánh số, chỉ có cỡ chữ và chữ đậm"),
        "D03_chi_co_so": (Style(size_by_level=(10, 10, 10), bold=False), None, "Có số, không có khác biệt về chữ. Số cấp 1 như `1 Giới thiệu` giống danh sách"),
        "D04_khong_gi_ca": (Style(numbered=False, size_by_level=(10, 10, 10), bold=False), None, "Không số, không khác biệt: bộ đoán không có căn cứ gì"),
        "D05_muc_luc": (Style(toc=True), None, "Có mục lục đầu tài liệu, không được tính là đầu mục"),
        "D06_hai_cot": (Style(two_column=True), None, "Hai cột: thứ tự đọc dễ sai"),
        "D07_chan_copy": (Style(), StandardEncryption("", canPrint=1, canCopy=0, canModify=0), "Chặn copy, không có mật khẩu mở"),
        "D08_co_mat_khau": (Style(), StandardEncryption("mat-khau-gia", canPrint=1), "Có mật khẩu mở"),
        "D09_co_dau_hieu_nguoi_benh": (Style(extra_body=[PHI_LINE]), None, "Có dòng giống thông tin người bệnh: không được gửi lên model"),
    }
    for name, (style, encrypt, note) in cases.items():
        directory = OUT / name
        directory.mkdir()
        gold = write_doc(directory / "source.pdf", style, encrypt)
        (directory / "answer.json").write_text(json.dumps({"case": name, "note": note, "simulated": True, **gold}, ensure_ascii=False, indent=1), encoding="utf-8")

    # Trang quét: render bản có lớp chữ thành ảnh rồi làm xấu nhẹ. Đáp án giữ từ bản gốc.
    import numpy as np
    import pypdfium2
    from PIL import Image

    source = pypdfium2.PdfDocument(str(OUT / "D01_so_va_dam" / "source.pdf"))
    directory = OUT / "D10_quet"
    directory.mkdir()
    rng = np.random.default_rng(10)
    c = canvas.Canvas(str(directory / "source.pdf"), pagesize=A4)
    for i in range(len(source)):
        image = source[i].render(scale=200 / 72).to_pil().convert("RGB")
        noisy = np.clip(np.array(image, dtype=np.float32) * 0.96 + rng.normal(0, 5, (image.height, image.width, 3)), 0, 255).astype(np.uint8)
        path = directory / f"page{i + 1}.jpg"
        Image.fromarray(noisy).save(path, quality=80)
        c.drawImage(str(path), 0, 0, WIDTH, HEIGHT)
        c.showPage()
    c.save()
    source.close()
    gold = json.loads((OUT / "D01_so_va_dam" / "answer.json").read_text(encoding="utf-8"))
    (directory / "answer.json").write_text(json.dumps({**gold, "case": "D10_quet", "note": "Trang quét từ D01, ảnh mô phỏng", "scanned": True}, ensure_ascii=False, indent=1), encoding="utf-8")
    print("Tài liệu:", len(list(OUT.iterdir())))


if __name__ == "__main__":
    main()
