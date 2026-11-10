from __future__ import annotations

import os
import pathlib

DATABASE_URL = os.environ.get("DATABASE_URL", "")
STORE_DIR = pathlib.Path(os.environ.get("STORE_DIR", "/store"))
SPOOL_DIR = pathlib.Path(os.environ.get("SPOOL_DIR", "/spool"))
CONVERTER_SPOOL_DIR = pathlib.Path(os.environ.get("CONVERTER_SPOOL_DIR", "/lospool"))
CONVERTER_TIMEOUT_SECONDS = int(os.environ.get("CONVERTER_TIMEOUT_SECONDS", "130"))
CONVERTER_ALIVE_SECONDS = 10
READER_TIMEOUT_SECONDS = int(os.environ.get("READER_TIMEOUT_SECONDS", "110"))
# Lớn hơn giới hạn của cổng đầu vào một chút: để reader báo số MB thật, API chỉ chặn file vượt xa.
MAX_UPLOAD_BYTES = 26 * 1024 * 1024
MAX_READER_OUTPUT_BYTES = 512 * 1024 * 1024
