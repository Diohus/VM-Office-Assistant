"""Một lệnh sinh mọi fixture giả mà test của `vmai` cần (T-AI-17, Q128). [MẪU GIẢ] Seed cố định.

Chạy từ gốc repo: vmai/.venv/Scripts/python.exe fixtures/eval/make_fixtures.py
`vmai/tests/conftest.py` tự gọi lệnh này khi thiếu fixture. Cần `vmai[ocr,dev]` (opencv, pillow, reportlab)
và font trong C:/Windows/Fonts, nên hiện chỉ sinh được trên Windows.

Token OCR trong `ocr_tokens/` là kết quả hai OCR nội bộ chạy trên ảnh giả do chính các script này sinh. Ảnh sinh ra
giống hệt từng byte giữa các lần chạy, nên token vẫn khớp. Chép vào đây để test không cần cài OCR.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
GENERATED = HERE / "generated"
TOKENS = HERE / "ocr_tokens"

# Thứ tự có nghĩa: generate_images_q90 dùng hàm của generate. resave_x12 cần Excel thật, không chạy ở đây.
SCRIPTS = ["generate.py", "generate_layouts.py", "generate_pdf_traps.py", "generate_word.py"]

# Đường dẫn test cần. Thiếu một đường sau khi sinh là lỗi.
REQUIRED = [
    "xlsx/X01_m0/source.xlsx",
    "xlsx/X01_m1/source.xlsx",
    "images/clean_s301/page.jpg",
    "pdf/P01/source.pdf",
    "layouts/L01_hai_dong_tieu_de/source.xlsx",
    "layouts/L34_ngang_may_2_ten_khac/source.xlsx",
    "pdf_traps/T9_o_trang_che_chu.pdf",
    "word/D10_quet/ocr_tokens.json",
]


def missing() -> list[str]:
    return [p for p in REQUIRED if not (GENERATED / p).exists()]


def make() -> None:
    env = {**os.environ, "PYTHONIOENCODING": "utf8", "PYTHONPATH": os.pathsep.join([str(ROOT / "vmai" / "src"), os.environ.get("PYTHONPATH", "")])}
    for name in SCRIPTS:
        done = subprocess.run([sys.executable, str(HERE / name)], cwd=ROOT, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if done.returncode != 0:
            raise SystemExit(f"{name} lỗi (mã {done.returncode}):\n{done.stderr[-2000:]}")
        print(f"{name}: xong")
    shutil.copyfile(TOKENS / "D10_quet.json", GENERATED / "word" / "D10_quet" / "ocr_tokens.json")
    left = missing()
    if left:
        raise SystemExit("Sinh xong mà vẫn thiếu: " + ", ".join(left))


if __name__ == "__main__":
    make()
    print("Đủ fixture cho test.")
