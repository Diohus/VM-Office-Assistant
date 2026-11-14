"""Máy giả thay OpenAI, chỉ dùng trong bộ test (compose.llm-test.yaml). Không có mạng ra ngoài.

Cách trả lời chọn theo dấu trong prompt, để test đi qua cả những đường lỗi xấu nhất:
- FAKE:ok          trả lời hợp lệ như Responses API, có số token.
- FAKE:echo_401    trả 401, thân lỗi lặp lại nguyên header Authorization (nhà cung cấp tồi nhất).
- FAKE:echo_500    trả 500, thân lỗi lặp lại header và prompt.
- FAKE:garbage     trả 200 nhưng thân không phải JSON, có lặp header.
- FAKE:refusal     model từ chối.
- FAKE:slow        chờ quá hạn của llm.
- FAKE:edit        model đề xuất B10 = 5.62 (có trong câu) và H10 = 7 (không có trong câu: cổng phải bỏ).
- FAKE:all_blocked model nói "Đã ghi xong." nhưng chỉ đề xuất ghi vào ô người (H10): cổng chặn hết (Q135 (e2)).
- FAKE:costly      trả lời hợp lệ nhưng báo nhiều token (để thử tiền theo bảng giá).
- FAKE:col_i       model ghi chữ "chạy lại" vào I8 đến I11 (Q135 (h2): ô đang có chữ thì không được ghi đè).
- FAKE:many        model ghi B8 đến B37 = 4.01 đến 4.30 (các số này có trong câu gõ): tóm tắt 30 ô (O1, 1.5).
- FAKE:long_ask    model hỏi lại bằng một câu rất dài (O1, 1.5).
- FAKE:fill_file   model báo việc này cần Điền từ file (unsupported fill_from_file, N4).
- FAKE:read_only   model ghi B10 = 5.62 và A10 = 5.62 (A10 chỉ được đọc khi vùng chọn là B8:C37: N2, CHAT_CELL_READ_ONLY).
- FAKE:outside     model ghi công thức "chạy lại" vào I8 (dòng dữ liệu) và I38 đến I44 (dòng thống kê ngoài thân bảng: K4).
- FAKE:already:<số> model ghi B10 = 5.62 và B12 = <số> (B12 đang đúng bằng số đó: N7, ô đã đúng sẵn).
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
        elif "FAKE:col_i" in prompt:
            ops = [{"op": "set_value", "cell": c, "type": "text", "value": "chạy lại", "formula": None} for c in ("I8", "I9", "I10", "I11")]
            self._send(200, json.dumps(ok_body(json.dumps({"status": "edit", "unsupported": None, "message": "Đã ghi.", "ops": ops}))).encode())
        elif "FAKE:many" in prompt:
            ops = [{"op": "set_value", "cell": f"B{8 + i}", "type": "number", "value": f"4.{i + 1:02d}", "formula": None} for i in range(30)]
            self._send(200, json.dumps(ok_body(json.dumps({"status": "edit", "unsupported": None, "message": "Đã ghi.", "ops": ops}))).encode())
        elif "FAKE:already:" in prompt:
            same = prompt.split("FAKE:already:", 1)[1].split()[0].strip('".,')
            ops = [{"op": "set_value", "cell": "B10", "type": "number", "value": "5.62", "formula": None},
                   {"op": "set_value", "cell": "B12", "type": "number", "value": same, "formula": None}]
            self._send(200, json.dumps(ok_body(json.dumps({"status": "edit", "unsupported": None, "message": "Đã ghi.", "ops": ops}))).encode())
        elif "FAKE:outside" in prompt:
            ops = [{"op": "set_formula", "cell": f"I{r}", "type": None, "value": None, "formula": f'=IF(F{r}>44,"chạy lại","")'} for r in (8, 38, 39, 40, 41, 42, 43, 44)]
            self._send(200, json.dumps(ok_body(json.dumps({"status": "edit", "unsupported": None, "message": "Đã ghi.", "ops": ops}))).encode())
        elif "FAKE:read_only" in prompt:
            ops = [{"op": "set_value", "cell": c, "type": "number", "value": "5.62", "formula": None} for c in ("B10", "A10")]
            self._send(200, json.dumps(ok_body(json.dumps({"status": "edit", "unsupported": None, "message": "Đã ghi.", "ops": ops}))).encode())
        elif "FAKE:fill_file" in prompt:
            self._send(200, json.dumps(ok_body(json.dumps({"status": "unsupported", "unsupported": "fill_from_file", "message": "Dùng Điền từ file.", "ops": []}))).encode())
        elif "FAKE:long_ask" in prompt:
            text = "Anh/chị cho biết rõ hơn: " + " ".join(f"ý {i} là ngày nào, mức nào, cột nào?" for i in range(1, 40))
            self._send(200, json.dumps(ok_body(json.dumps({"status": "ask", "unsupported": None, "message": text, "ops": []}))).encode())
        elif "FAKE:all_blocked" in prompt:
            ops = [{"op": "set_value", "cell": "H10", "type": "number", "value": "5.62", "formula": None}]
            self._send(200, json.dumps(ok_body(json.dumps({"status": "edit", "unsupported": None, "message": "Đã ghi xong.", "ops": ops}))).encode())
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
