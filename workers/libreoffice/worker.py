"""Worker LibreOffice: không mạng, không khóa, gốc chỉ đọc. Nhận việc qua thư mục spool, chỉ dùng thư viện chuẩn.

Việc: `xls_to_xlsx` chuyển file đời cũ; `pdf` xuất PDF và tính lại để so với lưới (Q48c). File xlsx do LibreOffice tạo
chỉ dùng để đọc số hoặc làm bản mở của file đời cũ, không bao giờ ghi số của nó vào biểu mẫu đang có.
"""
from __future__ import annotations

import io
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile

SPOOL = pathlib.Path(os.environ.get("SPOOL_DIR", "/spool"))
TIMEOUT_SECONDS = int(os.environ.get("JOB_TIMEOUT_SECONDS", "120"))
ALIVE = SPOOL / ".alive"
ALIVE_EVERY_SECONDS = 2
STALE_SECONDS = 600
# LibreOffice ghi tên vùng rỗng, ví dụ _xlfn_IFERROR, làm Excel không mở được file (đo ở M0).
_EMPTY_NAME = re.compile(rb"<definedName\b[^>]*>\s*</definedName>|<definedName\b[^>]*/>")
_EMPTY_NAMES = re.compile(rb"<definedNames>\s*</definedNames>")
_ITEMS = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
{}</oor:items>
"""
_LOCALE_ITEM = """<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>{}</value></prop></item>
"""
# Tính số theo dấu chấm thập phân, như lưới và công thức trong file: ở vi-VN, VALUE("12.5") ra #VALUE! và TEXT đổi dấu.
# Dùng en-GB, không dùng en-US: en-US làm sheet không khai khổ giấy thành Letter thay vì A4 (đo 05/11).
_COMPUTE = _ITEMS.format(_LOCALE_ITEM.format("en-GB"))
# In PDF từ file đã tính, không tính lại lần nữa, theo dấu thập phân của app (Q122, frontend/src/numberFormat.js):
# dấu chấm in theo en-GB (giữ khổ A4), dấu phẩy in theo vi-VN.
_NO_RECALC = """<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>1</value></prop></item>
"""
_PRINT = {".": _ITEMS.format(_LOCALE_ITEM.format("en-GB") + _NO_RECALC), ",": _ITEMS.format(_LOCALE_ITEM.format("vi-VN") + _NO_RECALC)}


def _soffice(source: pathlib.Path, fmt: str, outdir: pathlib.Path, profile: pathlib.Path) -> pathlib.Path:
    subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--norestore", "--convert-to", fmt, "--outdir", str(outdir), str(source)],
                   check=True, timeout=TIMEOUT_SECONDS, capture_output=True, env={"PATH": "/usr/bin:/bin", "HOME": "/tmp"})
    ext = fmt.split(":")[0]
    out = outdir / f"{source.stem}.{ext}"
    if not out.exists():
        raise FileNotFoundError(ext)
    return out


def _clean_names(data: bytes) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in zi.namelist():
            part = zi.read(name)
            if name == "xl/workbook.xml":
                part = _EMPTY_NAMES.sub(b"", _EMPTY_NAME.sub(b"", part))
            zo.writestr(name, part)
    return out.getvalue()


def _publish(path: pathlib.Path, data: bytes) -> None:
    fd, tmp = tempfile.mkstemp(dir=SPOOL, suffix=".tmp")
    with os.fdopen(fd, "wb") as fh:
        fh.write(data)
    os.replace(tmp, path)


def _profile(work: pathlib.Path, name: str, settings: str) -> pathlib.Path:
    profile = work / name
    (profile / "user").mkdir(parents=True)
    (profile / "user" / "registrymodifications.xcu").write_text(settings, encoding="utf8")
    return profile


def _run(job: dict, job_id: str) -> dict:
    with tempfile.TemporaryDirectory(dir="/tmp") as tmp:
        work = pathlib.Path(tmp)
        compute = _profile(work, "tinh", _COMPUTE)
        outdir = work / "out"
        outdir.mkdir()
        started = time.monotonic()
        if job["kind"] == "xls_to_xlsx":
            source = work / "vao.xls"
            shutil.copyfile(SPOOL / f"{job_id}.in", source)
            out = _soffice(source, "xlsx:Calc MS Excel 2007 XML", outdir, compute)
            _publish(SPOOL / f"{job_id}.out.xlsx", _clean_names(out.read_bytes()))
        elif job["kind"] == "pdf":
            source = work / "vao.xlsx"
            shutil.copyfile(SPOOL / f"{job_id}.in", source)
            values_dir = work / "so"
            values_dir.mkdir()
            values = _soffice(source, "xlsx:Calc MS Excel 2007 XML", values_dir, compute)
            # PDF in từ chính file đã tính, nên số trong PDF là số được đem so với lưới.
            pdf = _soffice(values, "pdf", outdir, _profile(work, "in", _PRINT[job.get("decimal", ".")]))
            _publish(SPOOL / f"{job_id}.out.xlsx", values.read_bytes())
            _publish(SPOOL / f"{job_id}.out.pdf", pdf.read_bytes())
        else:
            raise ValueError(job["kind"])
        return {"ok": True, "seconds": round(time.monotonic() - started, 2)}


def _handle(job_file: pathlib.Path) -> None:
    job_id = job_file.name.removesuffix(".lojob")
    result_file = SPOOL / f"{job_id}.out.json"
    try:
        result = _run(json.loads(job_file.read_text(encoding="utf8")), job_id)
    except subprocess.TimeoutExpired:
        result = {"ok": False, "reason": "timeout"}
    except Exception as exc:  # noqa: BLE001 - file lạ làm hỏng LibreOffice thì báo lỗi, worker phải sống tiếp
        print(f"job {job_id}: {type(exc).__name__}", file=sys.stderr, flush=True)
        result = {"ok": False, "reason": type(exc).__name__}
    _publish(result_file, json.dumps(result).encode("utf8"))
    job_file.unlink(missing_ok=True)
    (SPOOL / f"{job_id}.in").unlink(missing_ok=True)


def start_heartbeat(stop: threading.Event | None = None) -> threading.Thread:
    """Cập nhật .alive đều đặn từ một luồng riêng, kể cả khi đang chạy một việc dài (T-DEV-10).

    Trước đây chỉ cập nhật giữa các việc, nên việc PDF quá 10 giây (bảng lớn) bị API báo nhầm là worker không chạy.
    Worker chết thật thì luồng này chết theo, API vẫn báo đúng. Việc treo thì API có hạn chờ riêng (CONVERTER_TIMEOUT_SECONDS).
    """
    stop = stop or threading.Event()

    def beat() -> None:
        while not stop.is_set():
            try:
                ALIVE.touch()
            except OSError:
                pass
            stop.wait(ALIVE_EVERY_SECONDS)

    thread = threading.Thread(target=beat, name="alive", daemon=True)
    thread.start()
    return thread


def _sweep() -> None:
    limit = time.time() - STALE_SECONDS
    for path in SPOOL.iterdir():
        try:
            if path != ALIVE and path.stat().st_mtime < limit:
                path.unlink()
        except OSError:
            pass


def main() -> None:
    print("converter ready", flush=True)
    start_heartbeat()
    last = 0.0
    while True:
        now = time.monotonic()
        if now - last > ALIVE_EVERY_SECONDS:
            _sweep()          # .alive do luồng start_heartbeat cập nhật, kể cả lúc đang chạy việc
            last = now
        jobs = sorted(SPOOL.glob("*.lojob"))
        for job in jobs:
            _handle(job)
        if not jobs:
            time.sleep(0.1)


if __name__ == "__main__":
    main()
