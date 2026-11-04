"""Chạy hai OCR nội bộ trên các trang quét của D10, lưu token để vmai đọc. Không gửi gì ra ngoài.

Dùng môi trường của spike vì có sẵn OCR: spikes/ai/.venv/Scripts/python.exe fixtures/eval/ocr_word_pages.py
"""
from __future__ import annotations

import json
import sys
from dataclasses import asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "spikes" / "ai"))

from vmai.readers.image_prep import load_image  # noqa: E402
from vmai.readers.ocr import create_engine  # noqa: E402

directory = ROOT / "fixtures" / "eval" / "generated" / "word" / "D10_quet"
pages = sorted(directory.glob("page*.jpg"), key=lambda p: int(p.stem[4:]))
out = {}
for name in ("tesseract", "rapidocr"):
    engine = create_engine(name)
    out[name] = {}
    for path in pages:
        out[name][path.stem[4:]] = [asdict(t) for t in engine.read(load_image(path))]
    print(name, engine.reader_id, sum(len(v) for v in out[name].values()), "token")
(directory / "ocr_tokens.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
