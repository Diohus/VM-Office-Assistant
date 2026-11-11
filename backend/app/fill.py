"""Lát 1 nối AI (Q133): kiểm lô do reader trả về, hiểu quyết định của người dùng, tính ô cần ghi.

Nguyên tắc: giá trị ghi vào biểu mẫu chỉ lấy từ lô đã lưu. Trình duyệt chỉ gửi change_id và lựa chọn, không gửi số
(acceptance_lat1 4.6). Ngoại lệ duy nhất là ô đỏ người dùng tự gõ (D4): hiểu số theo đúng quy tắc như gõ trong lưới.
Không có hàm nào ở đây ghi log nội dung lô hay tên file.
"""
from __future__ import annotations

import re

SCHEMA_VERSION = "0.4.1"
SCHEMA_VERSIONS = ("0.4", "0.4.1")  # 0.4.1 thêm Notice.items, không bắt buộc (PM 11/11)
MAX_ITEMS = 200
MAX_ITEM_CHARS = 120
SPEC_ID = "GIA-QC-01"
# SHA-256 của fixtures/templates/GIA-QC-01_mau.xlsx. Lát 1 chỉ điền biểu mẫu này (PM, 11/11). Test so với file thật.
QC_TEMPLATE_SHA256 = "77df92dd69beb4a4d63b633ab7ab288e5a95decf524315e16f682b4c892eec8f"
MAX_CHANGES = 5000
COLORS_ALLOWED = {"yellow", "red"}


class BatchInvalid(Exception):
    """Lô reader trả về không đạt điều kiện lát 1. Không lưu, không cho người dùng xem."""


class DecisionInvalid(Exception):
    """Quyết định của người dùng không hợp lệ. `message` là câu tiếng Việt, không chứa số đã gửi."""

    def __init__(self, message: str, change_id: str | None = None):
        super().__init__(message)
        self.message = message
        self.change_id = change_id


# ---- Hiểu số gõ tay: bản Python của frontend/src/grid/typedNumber.js. Test dùng chung một bộ ca. ----
def parse_typed_number(raw) -> dict:
    """{kind: 'number', value} | {kind: 'ambiguous'} | {kind: 'unclear'} | {kind: 'other'}. Không đoán khi nhập nhằng."""
    text = str(raw if raw is not None else "").strip()
    if not re.fullmatch(r"-?[\d.,]*\d[\d.,]*", text):
        return {"kind": "other"}
    body = text[1:] if text.startswith("-") else text
    sign = -1 if text.startswith("-") else 1
    has_comma, has_dot = "," in body, "." in body
    if not has_comma and not has_dot:
        return {"kind": "number", "value": sign * float(body)}
    if has_comma and has_dot:
        decimal = "," if body.rfind(",") > body.rfind(".") else "."
        group = "." if decimal == "," else ","
        parts = body.split(decimal)
        if len(parts) != 2:
            return {"kind": "unclear"}
        whole, fraction = parts
        if not re.fullmatch(r"\d{1,3}(" + re.escape(group) + r"\d{3})+", whole) or not re.fullmatch(r"\d+", fraction):
            return {"kind": "unclear"}
        return {"kind": "number", "value": sign * float(whole.replace(group, "") + "." + fraction)}
    sep = "," if has_comma else "."
    parts = body.split(sep)
    if any(p == "" for p in parts):
        return {"kind": "unclear"}
    if len(parts) > 2:
        if re.fullmatch(r"\d{1,3}", parts[0]) and all(re.fullmatch(r"\d{3}", p) for p in parts[1:]):
            return {"kind": "number", "value": sign * float("".join(parts))}
        return {"kind": "unclear"}
    if re.fullmatch(r"[1-9]\d{0,2}", parts[0]) and len(parts[1]) == 3:
        return {"kind": "ambiguous"}
    return {"kind": "number", "value": sign * float(parts[0] + "." + parts[1])}


# ---- Kiểm lô ----
_A1 = re.compile(r"([A-Z]{1,3})([1-9][0-9]{0,6})")


def _cell_rc(a1: str) -> tuple[int, int]:
    """"B8" -> (7, 1). Không dùng openpyxl: API không có thư viện đọc file nào."""
    m = _A1.fullmatch(a1 or "")
    if not m:
        raise ValueError(a1)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, col - 1


def _input_cell(grid: dict, hidden: set[str], sheet: str, a1: str) -> bool:
    """Ô có thể nhận số: sheet có thật và không ẩn, trong vùng bảng, không phải ô công thức, không bị khóa, không nằm trong ô gộp (trừ ô đầu)."""
    sh = next((s for s in grid.get("sheets", []) if s.get("name") == sheet), None)
    if sh is None or sheet in hidden or sh.get("hidden"):
        return False
    try:
        r, c = _cell_rc(a1)
    except ValueError:
        return False
    if r < 0 or c < 0 or r >= sh.get("rows", 0) or c >= sh.get("cols", 0):
        return False
    cell = next((x for x in sh.get("cells", []) if x.get("r") == r and x.get("c") == c), None)
    if cell is not None and cell.get("f"):
        return False
    if sh.get("protected") and not (cell or {}).get("unlocked"):
        return False
    for m in sh.get("merges", []):
        r0, c0, r1, c1 = (m.get(k) for k in ("r0", "c0", "r1", "c1")) if isinstance(m, dict) else (m + [None] * 4)[:4]
        if None in (r0, c0, r1, c1):
            continue
        if r0 <= r <= r1 and c0 <= c <= c1 and (r, c) != (r0, c0):
            return False
    return True


def check_batch(batch: dict, grid: dict, hidden_sheets: list[str]) -> tuple[dict, dict]:
    """Kiểm lô theo điều kiện lát 1. Trả (targets, decisions ban đầu). Sai thì BatchInvalid (cả lô bị từ chối)."""
    if batch.get("schema_version") not in SCHEMA_VERSIONS:
        raise BatchInvalid("schema_version")
    notices = batch.get("notices") or []
    if not isinstance(notices, list):
        raise BatchInvalid("notices")
    for n in notices:
        items = (n or {}).get("items", [])
        if not isinstance(items, list) or len(items) > MAX_ITEMS or not all(isinstance(x, str) and len(x) <= MAX_ITEM_CHARS for x in items):
            raise BatchInvalid("items")
    if (batch.get("template") or {}).get("spec_id") != SPEC_ID:
        raise BatchInvalid("spec_id")
    changes = batch.get("changes")
    if not isinstance(changes, list) or len(changes) > MAX_CHANGES:
        raise BatchInvalid("changes")
    hidden = set(hidden_sheets or [])
    targets, decisions, seen = {}, {}, set()
    for ch in changes:
        cid = ch.get("change_id")
        target = ch.get("target") or {}
        sheet, a1 = target.get("sheet"), target.get("a1_at_write")
        color = (ch.get("confidence") or {}).get("final_color")
        if not isinstance(cid, str) or cid in targets or ch.get("op") != "set_value":
            raise BatchInvalid("change")
        if color not in COLORS_ALLOWED:  # lát 1 không có ô xanh (Q133); vmai đã hạ, đây là lớp hai
            raise BatchInvalid("color")
        if not isinstance(sheet, str) or not isinstance(a1, str) or not _input_cell(grid, hidden, sheet, a1):
            raise BatchInvalid("target")
        key = f"{sheet}!{a1}"
        if key in seen:  # mỗi ô nhiều nhất một Change (hợp đồng 0.4)
            raise BatchInvalid("duplicate")
        seen.add(key)
        value = ch.get("value")
        if value is None and color != "red":
            raise BatchInvalid("value")
        if value is not None and value.get("type") != "number":  # lát 1 chỉ điền số
            raise BatchInvalid("type")
        targets[cid] = key
        decisions[cid] = {"state": "pending"}
    return targets, decisions


# ---- Quyết định của người dùng ----
def _number(normalized) -> float:
    try:
        return float(str(normalized))
    except ValueError as exc:
        raise BatchInvalid("normalized") from exc


def _cell_value(grid: dict, key: str):
    sheet, a1 = key.split("!", 1)
    sh = next((s for s in grid.get("sheets", []) if s.get("name") == sheet), None)
    if sh is None:
        return None
    r, c = _cell_rc(a1)
    cell = next((x for x in sh.get("cells", []) if x.get("r") == r and x.get("c") == c), None)
    return None if cell is None else cell.get("v")


def _same(previous: dict, current) -> bool:
    """Ô đích còn như lúc tạo lô không (3.8a). previous rỗng thì ô phải còn trống."""
    prev = (previous or {}).get("value")
    if prev is None:
        return current is None or current == ""
    try:
        return abs(float(prev.get("normalized")) - float(current)) < 1e-9
    except (TypeError, ValueError):
        return str(prev.get("raw_text")) == str(current)


ACTIONS = {"accept", "reject", "red_accept", "red_choice", "red_typed", "red_blank"}


def decide(batch: dict, stored: dict, requests: list[dict], current_grid: dict) -> tuple[dict, list[dict], dict, dict]:
    """Áp các quyết định lên bản sao `stored`. Trả (decisions mới, edits để ghi, change_id cần duyệt lại, đếm theo cách).

    - accept: chỉ ô vàng không ghi đè ô có số (ô có số thì vẫn accept được nhưng phải gửi riêng từng ô; giao diện lo việc không nhận theo dòng).
    - red_accept: ô đỏ có số đề xuất, người dùng chấp nhận sau khi xem nguồn (D4).
    - red_choice: ô đỏ nhiều lần đọc, `index` là lần đọc; số lấy từ evidence[index].normalized.
    - red_typed: ô đỏ, người dùng gõ `text`; hiểu như gõ trong lưới, nhập nhằng thì từ chối.
    - red_blank: ô đỏ để trống. reject: bỏ đề xuất.
    Ô đích đã đổi so với lúc tạo lô thì không ghi, trả về {change_id: số hiện tại} để duyệt lại cũ → mới (3.8a, D6).
    Người dùng vẫn muốn ghi đè thì gửi lại `seen` là số hiện tại đã thấy; chỉ ghi khi ô vẫn đúng số đó. Số ghi vẫn lấy từ lô.
    """
    by_id = {ch["change_id"]: ch for ch in batch["changes"]}
    decisions = {k: dict(v) for k, v in stored.items()}
    edits, review_again = [], {}
    counts: dict[str, int] = {}
    seen = set()
    for req in requests:
        cid, action = req.get("change_id"), req.get("action")
        if cid not in by_id or action not in ACTIONS or cid in seen:
            raise DecisionInvalid("Có quyết định không thuộc lô này hoặc bị lặp nên app chưa ghi gì. Bản đang lưu không bị đổi. Hãy tải lại trang rồi duyệt lại.")
        seen.add(cid)
        if decisions[cid].get("state") != "pending":
            raise DecisionInvalid("Có ô đã được xử lý trước đó nên app chưa ghi gì. Bản đang lưu không bị đổi. Hãy tải lại trang rồi duyệt lại.", cid)
        ch = by_id[cid]
        color = ch["confidence"]["final_color"]
        target = ch["target"]
        if action == "reject":
            decisions[cid] = {"state": "rejected"}
            counts["rejected"] = counts.get("rejected", 0) + 1
            continue
        if action == "red_blank":
            if color != "red":
                raise DecisionInvalid("Chỉ ô đỏ mới để trống được. App chưa ghi gì.", cid)
            decisions[cid] = {"state": "blank"}
            counts["red_blank"] = counts.get("red_blank", 0) + 1
            continue
        if action == "accept":
            if color != "yellow":
                raise DecisionInvalid("Ô đỏ cần được xử lý riêng (xem nguồn, chọn số, gõ số hoặc để trống) nên app chưa ghi gì. Bản đang lưu không bị đổi.", cid)
            value, how = _number(ch["value"]["normalized"]), "proposal"
        elif action == "red_accept":
            if color != "red" or ch.get("value") is None:
                raise DecisionInvalid("Ô này không có số đề xuất để chấp nhận. App chưa ghi gì.", cid)
            value, how = _number(ch["value"]["normalized"]), "red_accept"
        elif action == "red_choice":
            index = req.get("index")
            evidence = ch.get("evidence") or []
            if color != "red" or not isinstance(index, int) or not 0 <= index < len(evidence) or evidence[index].get("normalized") is None:
                raise DecisionInvalid("Lựa chọn không thuộc ô này. App chưa ghi gì.", cid)
            value, how = _number(evidence[index]["normalized"]), f"red_choice:{index}"
        else:  # red_typed
            if color != "red":
                raise DecisionInvalid("Chỉ ô đỏ mới gõ số khác ở đây được. Ô khác hãy sửa trên lưới. App chưa ghi gì.", cid)
            parsed = parse_typed_number(req.get("text"))
            if parsed["kind"] == "ambiguous":
                raise DecisionInvalid("Số vừa gõ có thể hiểu hai cách (số nghìn hoặc số lẻ) nên app chưa ghi. Bản đang lưu không bị đổi. Hãy gõ không có dấu nghìn, dùng dấu chấm cho số lẻ.", cid)
            if parsed["kind"] != "number":
                raise DecisionInvalid("Số vừa gõ không phải một số rõ ràng nên app chưa ghi. Bản đang lưu không bị đổi. Hãy gõ lại, dùng dấu chấm cho số lẻ.", cid)
            value, how = parsed["value"], "red_typed"
        key = f"{target['sheet']}!{target['a1_at_write']}"
        current = _cell_value(current_grid, key)
        if not _same(ch.get("previous"), current):
            seen = req.get("seen")
            if seen is None or str(seen) != ("" if current is None else str(current)):
                review_again[cid] = current
                continue
        edits.append({"sheet": target["sheet"], "cell": target["a1_at_write"], "value": value})
        decisions[cid] = {"state": "written", "how": how, "version": None}
        counts[how] = counts.get(how, 0) + 1
    return decisions, edits, review_again, counts


def summary(source_name: str, counts: dict, written: int) -> str:
    """Tóm tắt bản trong lịch sử (D7, acceptance_lat1 3.7, 4.3): "Nhận N ô từ file A; ô đỏ: 1 chọn lần đọc 2, 1 tự gõ, 2 để trống".
    Không có ô đỏ thì không có phần sau dấu chấm phẩy."""
    text = f"Nhận {written} ô từ file {source_name}"
    red = []
    if counts.get("red_accept"):
        red.append(f"{counts['red_accept']} nhận số đề xuất")
    for key in sorted(k for k in counts if k.startswith("red_choice:")):
        red.append(f"{counts[key]} chọn lần đọc {int(key.split(':')[1]) + 1}")
    if counts.get("red_typed"):
        red.append(f"{counts['red_typed']} tự gõ")
    if counts.get("red_blank"):
        red.append(f"{counts['red_blank']} để trống")
    return text + (f"; ô đỏ: {', '.join(red)}" if red else "")
