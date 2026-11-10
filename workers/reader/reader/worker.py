"""Vòng lặp của container reader: không mạng, không khóa, gốc chỉ đọc. Trao đổi với API qua thư mục spool.

Mỗi file chạy trong một tiến trình con có giới hạn thời gian và bộ nhớ, để file độc không làm chết worker.
Quy ước: API ghi <id>.xlsx, rồi <id>.job (chứa tên file) sau cùng. Worker ghi <id>.out.json bằng cách đổi tên nguyên tử.
"""
from __future__ import annotations

import json
import os
import pathlib
import resource
import subprocess
import sys
import tempfile
import time

SPOOL = pathlib.Path(os.environ.get("SPOOL_DIR", "/spool"))
POLL_SECONDS = 0.1
JOB_TIMEOUT_SECONDS = int(os.environ.get("JOB_TIMEOUT_SECONDS", "90"))
MEMORY_LIMIT_BYTES = int(os.environ.get("JOB_MEMORY_BYTES", str(3 * 1024 ** 3)))
MAX_OUTPUT_BYTES = 1024 ** 3
STALE_SECONDS = 600
CHILD_ARGS = ["-m", "reader.child"]
SWEEP_EVERY_SECONDS = 60


def _limits() -> None:
    resource.setrlimit(resource.RLIMIT_AS, (MEMORY_LIMIT_BYTES, MEMORY_LIMIT_BYTES))
    resource.setrlimit(resource.RLIMIT_CPU, (JOB_TIMEOUT_SECONDS, JOB_TIMEOUT_SECONDS + 5))
    resource.setrlimit(resource.RLIMIT_FSIZE, (MAX_OUTPUT_BYTES, MAX_OUTPUT_BYTES))
    resource.setrlimit(resource.RLIMIT_NOFILE, (64, 64))


def _write_atomic(path: pathlib.Path, text: str) -> None:
    fd, tmp = tempfile.mkstemp(dir=SPOOL, suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf8") as fh:
        fh.write(text)
    os.replace(tmp, path)


def _failed(reason: str) -> str:
    return json.dumps({"ok": False, "code": "FILE_NOT_READABLE", "message": "App không xử lý xong file này trong thời gian cho phép. File của bạn không bị đổi. Hãy thử lại, hoặc hỏi người quản lý app nếu vẫn không được.", "detail": {"reason": reason}})


def _run_job(job: pathlib.Path) -> None:
    job_id = job.name.removesuffix(".job")
    out = SPOOL / f"{job_id}.out.json"
    try:
        proc = subprocess.run([sys.executable, *CHILD_ARGS, str(SPOOL / f"{job_id}.xlsx"), str(job), str(out)],
                              env={"PATH": "/usr/bin:/bin", "PYTHONPATH": os.environ.get("PYTHONPATH", ""), "HOME": "/tmp"},
                              preexec_fn=_limits, timeout=JOB_TIMEOUT_SECONDS, capture_output=True)
        # Không xét việc file kết quả còn hay không: API có thể đã lấy và xóa nó ngay khi tiến trình con vừa ghi.
        if proc.returncode != 0:
            print(f"job {job_id} rc={proc.returncode}", file=sys.stderr, flush=True)
            _write_atomic(out, _failed("crash"))
    except subprocess.TimeoutExpired:
        print(f"job {job_id} timeout", file=sys.stderr, flush=True)
        _write_atomic(out, _failed("timeout"))
    finally:
        job.unlink(missing_ok=True)
        (SPOOL / f"{job_id}.xlsx").unlink(missing_ok=True)


def _sweep() -> None:
    """Dọn file sót khi API đã bỏ cuộc: nội dung tải lên không được nằm lại trong spool."""
    limit = time.time() - STALE_SECONDS
    for path in SPOOL.iterdir():
        try:
            if path.stat().st_mtime < limit:
                path.unlink()
        except OSError:
            pass


def main() -> None:
    print("reader ready", flush=True)
    last_sweep = time.monotonic()
    while True:
        jobs = sorted(SPOOL.glob("*.job"))
        for job in jobs:
            _run_job(job)
        if time.monotonic() - last_sweep > SWEEP_EVERY_SECONDS:
            _sweep()
            last_sweep = time.monotonic()
        if not jobs:
            time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
