"""Dịch vụ llm (Q135): việc duy nhất là gọi model với payload api đã dựng, trả chữ thô và số token.

An toàn:
- Key chỉ đọc từ file secret /run/secrets/llm_env (dòng OPENAI_API_KEY=...), lúc gọi, không đặt vào biến môi trường.
- Không ghi payload, câu trả lời, key hay header vào log. Lỗi trả về api chỉ là mã, không kèm chữ lỗi gốc của nhà cung cấp.
- Ra Internet chỉ qua egress (CONNECT api.openai.com:443). TLS xác minh đầy đủ (Q85).
- Chỉ thư viện chuẩn của Python: không có gói ngoài phải rà.
Model cố định gpt-6-luna (Q85, Q135). Đổi model phải có quyết định PM.
"""
from __future__ import annotations

import hashlib
import http.client
import http.server
import json
import logging
import os
import pathlib
import ssl
import time
import urllib.error
import urllib.request

MODEL = "gpt-6-luna"
SECRET_FILE = pathlib.Path(os.environ.get("LLM_SECRET_FILE", "/run/secrets/llm_env"))
KEY_NAME = "OPENAI_API_KEY"
BASE_URL = os.environ.get("LLM_BASE_URL", "https://api.openai.com/v1").rstrip("/")
PROXY = os.environ.get("LLM_PROXY", "http://egress:3128")
CA_FILE = os.environ.get("LLM_CA_FILE", "/opt/llm/ca.pem")
# PM 14/11: ghi cả cột (28 công thức) đo 13-23 giây, 25 giây quá sát. Thứ tự hạn chờ: backend/tests/test_timeouts.py.
UPSTREAM_TIMEOUT_SECONDS = float(os.environ.get("LLM_UPSTREAM_TIMEOUT_SECONDS", "45"))
MAX_REQUEST_BYTES = 1024 * 1024
MAX_RESPONSE_BYTES = 2 * 1024 * 1024
MAX_PROMPT_CHARS = 200_000
MAX_OUTPUT_TOKENS = 8000

log = logging.getLogger("llm")


class CallFailed(Exception):
    """Lỗi có mã. Không bao giờ mang chữ lỗi gốc, header hay nội dung."""

    def __init__(self, code: str, not_sent: bool = False):
        super().__init__(code)
        self.code = code
        # PM 14/11: True khi chắc chắn chưa gửi byte nào tới nhà cung cấp (chưa mở xong kết nối qua egress). api ghi tiền 0.
        self.not_sent = not_sent


def read_key(path: pathlib.Path | None = None) -> str:
    """Đọc đúng một dòng KEY_NAME trong file secret. Không in, không trả về chỗ khác ngoài lệnh gọi."""
    try:
        text = (path or SECRET_FILE).read_text(encoding="utf-8-sig")
    except OSError:
        raise CallFailed("CHAT_MODEL_NO_KEY") from None
    for line in text.splitlines():
        name, sep, value = line.strip().partition("=")
        if sep and name.strip() == KEY_NAME:
            value = value.strip().strip('"').strip("'")
            if value:
                return value
    raise CallFailed("CHAT_MODEL_NO_KEY")


def build_payload(req: dict) -> dict:
    """Payload Responses API, giống vmai.llm.openai_adapter.build_payload: store False, structured output strict, không tool."""
    return {
        "model": MODEL,
        "instructions": req["system"],
        "input": [{"role": "user", "content": [{"type": "input_text", "text": req["prompt"]}]}],
        "max_output_tokens": req["max_output_tokens"],
        "store": False,
        "text": {"format": {"type": "json_schema", "name": req["schema_name"], "schema": req["response_schema"], "strict": True}},
    }


def validate(req: object) -> dict:
    if not isinstance(req, dict):
        raise CallFailed("CHAT_BAD_REQUEST")
    ok = (
        isinstance(req.get("system"), str) and 0 < len(req["system"]) <= 20_000
        and isinstance(req.get("prompt"), str) and 0 < len(req["prompt"]) <= MAX_PROMPT_CHARS
        and isinstance(req.get("schema_name"), str) and 0 < len(req["schema_name"]) <= 64
        and isinstance(req.get("response_schema"), dict)
        and isinstance(req.get("max_output_tokens"), int) and not isinstance(req["max_output_tokens"], bool)
        and 1 <= req["max_output_tokens"] <= MAX_OUTPUT_TOKENS
    )
    if not ok:
        raise CallFailed("CHAT_BAD_REQUEST")
    return req


def parse_response(body: dict) -> dict:
    """Lấy chữ trả lời và số token. Câu từ chối của model thành mã, không chép chữ."""
    status = body.get("status")
    if status == "incomplete":
        raise CallFailed("CHAT_MODEL_TRUNCATED")
    if status != "completed":
        raise CallFailed("CHAT_MODEL_FAILED")
    texts = []
    for item in body.get("output") or []:
        if not isinstance(item, dict) or item.get("type") != "message":
            continue
        for block in item.get("content") or []:
            if not isinstance(block, dict):
                continue
            if block.get("type") == "refusal":
                raise CallFailed("CHAT_MODEL_REFUSED")
            if block.get("type") == "output_text" and isinstance(block.get("text"), str):
                texts.append(block["text"])
    usage = body.get("usage") if isinstance(body.get("usage"), dict) else {}

    def count(v: object) -> int | None:
        return v if isinstance(v, int) and not isinstance(v, bool) and v >= 0 else None

    details = usage.get("output_tokens_details") if isinstance(usage.get("output_tokens_details"), dict) else {}
    return {
        "ok": True,
        "text": "".join(texts),
        "model": body.get("model") if isinstance(body.get("model"), str) else MODEL,
        "request_id": body.get("id") if isinstance(body.get("id"), str) else None,
        "input_tokens": count(usage.get("input_tokens")),
        "output_tokens": count(usage.get("output_tokens")),
        "reasoning_tokens": count(details.get("reasoning_tokens")),
    }


class _TrackedHTTPSHandler(urllib.request.HTTPSHandler):
    """Như HTTPSHandler, nhưng đánh dấu request.vm_connected = True khi kết nối (TCP, đường hầm qua egress, TLS) đã mở xong.
    Lỗi xảy ra khi cờ này chưa bật thì chắc chắn chưa gửi byte nào của yêu cầu tới nhà cung cấp."""

    def https_open(self, req):
        def connection(*args, **kwargs):
            conn = http.client.HTTPSConnection(*args, **kwargs)
            connect = conn.connect

            def tracked() -> None:
                connect()
                req.vm_connected = True

            conn.connect = tracked
            return conn
        return self.do_open(connection, req, context=self._context)


def _opener() -> urllib.request.OpenerDirector:
    context = ssl.create_default_context(cafile=CA_FILE if os.path.exists(CA_FILE) else None)
    handlers: list = [_TrackedHTTPSHandler(context=context)]
    # Chỉ https đi qua egress. http chỉ dùng cho máy giả trong bộ test (LLM_BASE_URL do compose test đặt).
    handlers.append(urllib.request.ProxyHandler({"https": PROXY} if PROXY else {}))
    return urllib.request.build_opener(*handlers)


def _not_sent(request: urllib.request.Request) -> bool:
    """Chỉ đường https qua _TrackedHTTPSHandler mới biết chắc. http (máy giả của bộ test) thì coi như có thể đã gửi."""
    return request.type == "https" and not getattr(request, "vm_connected", False)


def call(req: dict, opener: urllib.request.OpenerDirector | None = None) -> dict:
    req = validate(req)
    data = json.dumps(build_payload(req)).encode()
    key = read_key()
    request = urllib.request.Request(f"{BASE_URL}/responses", data=data, method="POST",
                                     headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"})
    del key
    started = time.perf_counter()
    try:
        with (opener or _opener()).open(request, timeout=UPSTREAM_TIMEOUT_SECONDS) as resp:
            raw = resp.read(MAX_RESPONSE_BYTES + 1)
    except urllib.error.HTTPError as exc:
        # Không đọc thân lỗi: nhà cung cấp có thể lặp lại header hay nội dung yêu cầu.
        exc.close()
        raise CallFailed({401: "CHAT_MODEL_KEY_REJECTED", 403: "CHAT_MODEL_KEY_REJECTED", 429: "CHAT_MODEL_BUSY"}.get(exc.code, "CHAT_MODEL_UNAVAILABLE")) from None
    except TimeoutError:
        raise CallFailed("CHAT_MODEL_TIMEOUT", not_sent=_not_sent(request)) from None
    except (urllib.error.URLError, OSError) as exc:
        timeout = isinstance(getattr(exc, "reason", None), TimeoutError)
        raise CallFailed("CHAT_MODEL_TIMEOUT" if timeout else "CHAT_MODEL_UNAVAILABLE", not_sent=_not_sent(request)) from None
    finally:
        request.headers.clear()
        request.unredirected_hdrs.clear()
    if len(raw) > MAX_RESPONSE_BYTES:
        raise CallFailed("CHAT_MODEL_FAILED")
    try:
        body = json.loads(raw)
    except ValueError:
        raise CallFailed("CHAT_MODEL_FAILED") from None
    if not isinstance(body, dict):
        raise CallFailed("CHAT_MODEL_FAILED")
    out = parse_response(body)
    out["latency_ms"] = int((time.perf_counter() - started) * 1000)
    return out


class Handler(http.server.BaseHTTPRequestHandler):
    server_version = "llm"
    sys_version = ""

    def log_message(self, *args) -> None:  # không ghi dòng truy cập mặc định (có thể kèm đường dẫn, địa chỉ)
        pass

    def _send(self, status: int, payload: dict) -> None:
        data = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        if self.path == "/health":
            self._send(200, {"ok": True})
        else:
            self._send(404, {"ok": False, "code": "NOT_FOUND"})

    def do_POST(self) -> None:
        if self.path != "/call":
            self._send(404, {"ok": False, "code": "NOT_FOUND"})
            return
        started = time.perf_counter()
        code, not_sent = "OK", False
        try:
            length = int(self.headers.get("Content-Length") or 0)
            if not 0 < length <= MAX_REQUEST_BYTES:
                raise CallFailed("CHAT_BAD_REQUEST")
            try:
                req = json.loads(self.rfile.read(length))
            except ValueError:
                raise CallFailed("CHAT_BAD_REQUEST") from None
            out = call(req)
            out["prompt_sha256"] = hashlib.sha256(req["prompt"].encode()).hexdigest()
            self._send(200, out)
            # Mã băm prompt (12 ký tự) để đối chiếu với màn xem trước, không phải nội dung.
            log.info("call ok prompt=%s in=%s out=%s ms=%s", out["prompt_sha256"][:12], out["input_tokens"], out["output_tokens"], out["latency_ms"])
            return
        except CallFailed as exc:
            code, not_sent = exc.code, exc.not_sent
        except Exception as exc:  # noqa: BLE001 - lỗi lạ: chỉ ghi tên loại lỗi, không ghi chữ lỗi
            code = "CHAT_MODEL_UNAVAILABLE"
            log.warning("call error type=%s", type(exc).__name__)
        self._send(200 if code != "CHAT_BAD_REQUEST" else 400, {"ok": False, "code": code, "not_sent": not_sent})
        log.info("call failed code=%s not_sent=%s ms=%s", code, not_sent, int((time.perf_counter() - started) * 1000))


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(message)s")
    server = http.server.ThreadingHTTPServer(("0.0.0.0", 8100), Handler)
    server.daemon_threads = True
    log.info("llm ready model=%s", MODEL)
    server.serve_forever()


if __name__ == "__main__":
    main()
