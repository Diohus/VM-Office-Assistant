"""Tiến trình con xử lý đúng một file rồi thoát."""
from __future__ import annotations

import json
import os
import pathlib
import sys

from .job import process


def main() -> int:
    source, job, out = map(pathlib.Path, sys.argv[1:4])
    filename = job.read_text(encoding="utf8")
    result = process(source.read_bytes(), filename, source)
    tmp = out.with_name(out.name + ".tmp")
    tmp.write_text(json.dumps(result, ensure_ascii=False), encoding="utf8")
    os.replace(tmp, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
