"""T-DEV-10: việc PDF dài không bị báo nhầm là worker không chạy; quá hạn thì báo đúng lý do.

Chạy được không cần container: nạp thẳng worker và converter_client với thư mục spool tạm.
Ca worker chết thật (tắt container) nằm ở test_pdf_legacy.py::test_tat_worker_bao_ro_va_excel_van_xuat_duoc.
"""
from __future__ import annotations

import importlib.util
import json
import pathlib
import sys
import threading
import time

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[2]


def load_worker(spool: pathlib.Path, mp):
    mp.setenv("SPOOL_DIR", str(spool))
    spec = importlib.util.spec_from_file_location(f"lo_worker_{time.monotonic_ns()}", ROOT / "workers" / "libreoffice" / "worker.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def load_client(spool: pathlib.Path, timeout: int, mp):
    mp.setenv("CONVERTER_SPOOL_DIR", str(spool))
    mp.setenv("CONVERTER_TIMEOUT_SECONDS", str(timeout))
    sys.path.insert(0, str(ROOT / "backend"))
    try:
        for name in [n for n in sys.modules if n == "app" or n.startswith("app.")]:
            del sys.modules[name]
        return importlib.import_module("app.converter_client")
    finally:
        sys.path.remove(str(ROOT / "backend"))


def test_viec_dai_hon_10_giay_van_con_song(tmp_path, monkeypatch):
    worker = load_worker(tmp_path, monkeypatch)
    stop = threading.Event()
    worker.start_heartbeat(stop)
    # Việc giả 12 giây, dài hơn ngưỡng 10 giây của API.
    worker._run = lambda job, job_id: (time.sleep(12), {"ok": True, "seconds": 12})[1]
    (tmp_path / "j1.in").write_bytes(b"x")
    (tmp_path / "j1.lojob").write_text(json.dumps({"kind": "pdf"}), encoding="utf8")
    job = threading.Thread(target=worker._handle, args=(tmp_path / "j1.lojob",))
    job.start()
    ages = []
    while job.is_alive():
        ages.append(time.time() - (tmp_path / ".alive").stat().st_mtime)
        time.sleep(0.5)
    stop.set()
    assert max(ages) < 10, f".alive cũ tới {max(ages):.1f} giây trong lúc chạy việc"
    assert json.loads((tmp_path / "j1.out.json").read_text(encoding="utf8"))["ok"] is True


def test_qua_han_thi_bao_qua_lau_khong_bao_khong_chay(tmp_path, monkeypatch):
    client = load_client(tmp_path, 3, monkeypatch)
    stop = threading.Event()

    def alive_forever():
        while not stop.is_set():
            (tmp_path / ".alive").touch()
            stop.wait(1)

    (tmp_path / ".alive").touch()
    threading.Thread(target=alive_forever, daemon=True).start()
    try:
        with pytest.raises(client.ConverterTimeout):
            client.pdf(b"x")
    finally:
        stop.set()


def test_worker_tu_bao_het_gio_thi_la_qua_lau(tmp_path, monkeypatch):
    client = load_client(tmp_path, 20, monkeypatch)
    (tmp_path / ".alive").touch()

    def answer():
        for _ in range(100):
            jobs = list(tmp_path.glob("*.lojob"))
            if jobs:
                job_id = jobs[0].name.removesuffix(".lojob")
                (tmp_path / f"{job_id}.out.json").write_text(json.dumps({"ok": False, "reason": "timeout"}), encoding="utf8")
                return
            time.sleep(0.05)

    threading.Thread(target=answer, daemon=True).start()
    with pytest.raises(client.ConverterTimeout):
        client.pdf(b"x")


def test_khong_co_nhip_tim_thi_bao_khong_chay(tmp_path, monkeypatch):
    client = load_client(tmp_path, 20, monkeypatch)
    with pytest.raises(client.ConverterUnavailable):
        client.pdf(b"x")
