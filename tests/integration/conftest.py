"""Kiểm thử tầng 1: chạy với toàn bộ app thật (compose) và adapter AI giả. Không gọi dịch vụ ngoài."""
from __future__ import annotations

import http.server
import json
import os
import pathlib
import threading
import urllib.error
import urllib.parse
import urllib.request

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[2]
FIX = ROOT / "fixtures" / "templates"
GENERATED = ROOT / "fixtures" / "generated"
BASE = os.environ.get("VM_BASE", "http://127.0.0.1:8089")
PROJECT = os.environ.get("VM_PROJECT", "vm-office-test")


def request(method: str, path: str, data: bytes | None = None, timeout: int = 130, headers: dict | None = None):
    req = urllib.request.Request(BASE + path, data=data, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as res:
            return res.status, dict(res.headers), res.read()
    except urllib.error.HTTPError as err:
        return err.code, dict(err.headers), err.read()


def upload(data: bytes, filename: str):
    status, _, body = request("POST", f"/api/documents?filename={urllib.parse.quote(filename)}", data)
    return status, json.loads(body)


@pytest.fixture(scope="session", autouse=True)
def stack_is_up():
    status, _, _ = request("GET", "/api/health", timeout=20)
    if status != 200:
        pytest.exit("App chưa chạy. Dùng: sh tests/integration/run.sh", returncode=2)


class _Canary(http.server.BaseHTTPRequestHandler):
    hits: list[str] = []

    def do_GET(self):  # noqa: N802 - tên do thư viện chuẩn đặt
        type(self).hits.append(self.path)
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"1")

    do_POST = do_GET

    def log_message(self, *_):
        pass


@pytest.fixture(scope="session")
def canary():
    """Máy chủ bẫy trên máy này. Nếu trình duyệt hay gì khác làm theo địa chỉ trong file, nó sẽ ghi lại."""
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), _Canary)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    yield {"url": f"http://127.0.0.1:{server.server_port}", "hits": _Canary.hits}
    server.shutdown()
