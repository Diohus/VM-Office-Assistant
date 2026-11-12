"""Máy giả thay OpenAI, chỉ dùng trong bộ test (compose.llm-test.yaml). Không có mạng ra ngoài.

Cách trả lời chọn theo dấu trong prompt, để test đi qua cả những đường lỗi xấu nhất:
- FAKE:ok          trả lời hợp lệ như Responses API, có số token.
- FAKE:echo_401    trả 401, thân lỗi lặp lại nguyên header Authorization (nhà cung cấp tồi nhất).
- FAKE:echo_500    trả 500, thân lỗi lặp lại header và prompt.
- FAKE:garbage     trả 200 nhưng thân không phải JSON, có lặp header.
- FAKE:refusal     model từ chối.
- FAKE:slow        chờ quá hạn của llm.
- FAKE:edit        model đề xuất B10 = 5.62 (có trong câu) và H10 = 7 (không có trong câu: cổng phải bỏ).
- FAKE:costly      trả lời hợp lệ nhưng báo nhiều token (để thử tiền theo bảng giá).
Mặc định: model hỏi lại (status ask).
GET /seen trả SHA-256 của header Authorization lần gần nhất, để test biết key thật sự đã đi tới đây (không trả chính header).
"""
from __future__ import annotations

import hashlib
import http.server
import json
import time

LAST = {"auth_sha256": None, "calls": 0}


def ok_body(text: str) -> dict:
    return {
        "id": "resp_fake_1", "status": "completed", "model": "gpt-6-luna",
        "output": [{"type": "message", "content": [{"type": "output_text", "text": text}]}],
        "usage": {"input_tokens": 1200, "output_tokens": 300, "output_tokens_details": {"reasoning_tokens": 100}},
    }


class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *args) -> None:
        pass

    def _send(self, status: int, data: bytes, kind: str = "application/json") -> None:
        self.send_response(status)
        self.send_header("Content-Type", kind)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        self._send(200, json.dumps(LAST).encode())

    def do_POST(self) -> None:
        auth = self.headers.get("Authorization", "")
        LAST["auth_sha256"] = hashlib.sha256(auth.encode()).hexdigest()
        LAST["calls"] += 1
        body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)))
        prompt = body["input"][0]["content"][0]["text"]
        if "FAKE:echo_401" in prompt:
            self._send(401, json.dumps({"error": {"message": f"Incorrect API key provided: {auth}"}}).encode())
        elif "FAKE:echo_500" in prompt:
            self._send(500, json.dumps({"error": {"message": f"boom {auth} {prompt}"}}).encode())
        elif "FAKE:garbage" in prompt:
            self._send(200, f"<html>{auth}</html>".encode(), "text/html")
        elif "FAKE:refusal" in prompt:
            self._send(200, json.dumps({"id": "r", "status": "completed", "model": "gpt-6-luna",
                                        "output": [{"type": "message", "content": [{"type": "refusal", "refusal": f"no {auth}"}]}]}).encode())
        elif "FAKE:edit" in prompt:
            ops = [{"op": "set_value", "cell": "B10", "type": "number", "value": "5.62", "formula": None},
                   {"op": "set_value", "cell": "H10", "type": "number", "value": "7", "formula": None}]
            self._send(200, json.dumps(ok_body(json.dumps({"status": "edit", "unsupported": None, "message": "Đã sửa.", "ops": ops}))).encode())
        elif "FAKE:costly" in prompt:
            body = ok_body(json.dumps({"status": "ask", "unsupported": None, "message": "Ô nào?", "ops": []}))
            body["usage"] = {"input_tokens": 50000, "output_tokens": 6000}
            self._send(200, json.dumps(body).encode())
        elif "FAKE:slow" in prompt:
            time.sleep(40)
            self._send(200, json.dumps(ok_body("{}")).encode())
        else:
            self._send(200, json.dumps(ok_body(json.dumps({"status": "ask", "unsupported": None, "message": "Anh/chị muốn sửa ô nào?", "ops": []}))).encode())


if __name__ == "__main__":
    server = http.server.ThreadingHTTPServer(("0.0.0.0", 8080), Handler)
    server.daemon_threads = True
    server.serve_forever()
