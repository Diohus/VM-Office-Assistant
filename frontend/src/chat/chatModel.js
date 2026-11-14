// Khung chat lát 2 (Q135): phần không đụng tới lưới, để test bằng node. Hợp đồng: docs/contracts/chat_edit_plan.md.
// Câu chữ chỉ lấy từ chat_texts trong vmai/src/vmai/data/doc_codes.json (AI Engineer viết, PO/BA duyệt).

export const MAX_REGION_CELLS = 3000;
export const AI_MARK = '#E9DCF7'; // PM chốt 12/11: nền tím nhạt, ngoài thang xanh/vàng/đỏ của Điền từ file
// Câu tạm khi chat_texts chưa có mã: chỉ dùng trong lúc làm. Test bảo đảm không mã nào rơi vào đây trước nghiệm thu.
export const TEMP_PREFIX = '[câu tạm]';
const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 31); // số ngày 1 là 01/01/1900 theo Excel

// Câu theo mã, điền chỗ trống {tên}. Không có mã thì trả câu tạm có đánh dấu, không im lặng.
export function chatText(texts, code, vars = {}) {
  const raw = texts?.chat_texts?.[code];
  if (typeof raw !== 'string') return `${TEMP_PREFIX} ${code}`;
  return raw.replace(/\{([^{}]+)\}/g, (m, key) => (key in vars ? String(vars[key]) : m));
}

// Số ngày Excel (hệ 1900) sang ISO. Excel coi 29/02/1900 là ngày có thật (số 60): không có ngày tương ứng thì trả null.
export function serialToIso(serial) {
  if (typeof serial !== 'number' || !Number.isFinite(serial) || serial < 1) return null;
  let days = Math.floor(serial);
  if (days === 60) return null;
  if (days > 60) days -= 1;
  return new Date(EPOCH + days * DAY_MS).toISOString().slice(0, 10);
}

export function isoToSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso ?? '');
  if (!m) return null;
  const t = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(t);
  if (d.getUTCMonth() !== Number(m[2]) - 1 || d.getUTCDate() !== Number(m[3])) return null;
  const days = Math.round((t - EPOCH) / DAY_MS);
  if (days < 1) return null;
  return days >= 60 ? days + 1 : days;
}

// Mẫu định dạng là ngày: có d, y, hoặc m mà không có giờ, giây. Bỏ phần trong ngoặc kép, [màu]/[điều kiện], ký tự thoát.
export function isDatePattern(pattern) {
  if (typeof pattern !== 'string' || !pattern || /^general$/i.test(pattern)) return false;
  const bare = pattern.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '').replace(/_.|\*./g, '');
  if (/[dy]/i.test(bare)) return true;
  return /m/i.test(bare) && !/[hs]/i.test(bare);
}

// Sheet mà công thức tham chiếu: 'Tên có cách'!A1 hoặc Ten!A1. Chỉ trả tên có trong danh sách sheet của sổ.
export function referencedSheets(formula, sheetNames) {
  if (typeof formula !== 'string') return [];
  const found = new Set();
  const text = formula.replace(/"(?:[^"]|"")*"/g, '');
  for (const m of text.matchAll(/'((?:[^']|'')+)'!|([A-Za-z_À-ỹ][\w.À-ỹ]*)!/g)) {
    const name = m[1] ? m[1].replace(/''/g, "'") : m[2];
    const hit = sheetNames.find((s) => s.toLowerCase() === name.toLowerCase());
    if (hit) found.add(hit);
  }
  return [...found];
}

// Một ô của bản chụp gửi đi (ChatCell). raw: { v, f, pattern, display }. Ô trống không gửi: trả null.
export function chatCell(a1, raw) {
  const f = raw.f ? (raw.f.startsWith('=') ? raw.f : `=${raw.f}`) : null;
  const v = raw.v;
  const display = raw.display ?? null;
  if ((v === null || v === undefined || v === '') && !f) return null;
  if (typeof v === 'boolean') return { a1, t: 'boolean', v: v ? 'true' : 'false', f, display };
  if (typeof v === 'number' && Number.isFinite(v)) {
    const iso = isDatePattern(raw.pattern) ? serialToIso(v) : null;
    return iso ? { a1, t: 'date', v: iso, f, display } : { a1, t: 'number', v: String(v), f, display };
  }
  // O2: công thức chưa có kết quả (bộ tính còn chạy) gửi chữ rỗng kèm công thức. Gửi null thì vmai từ chối cả lượt (CHAT_BAD_INPUT).
  if (v === null || v === undefined) return { a1, t: 'text', v: '', f, display };
  return { a1, t: 'text', v: String(v), f, display };
}

// Ô hiện tại có còn như lúc gửi không. Không có trong bản chụp nghĩa là lúc gửi ô trống.
export function sameAsSent(sent, now) {
  if (!sent) return now === null;
  if (!now) return false;
  return sent.t === now.t && sent.v === now.v && (sent.f ?? null) === (now.f ?? null);
}

// Khóa để biết có phải hiện lại màn xem trước không (2.5, 2.6): đổi sheet, vùng, hoặc ô bị che thì hiện lại.
export function previewKey(request, preview) {
  // 0.2: phần đọc kèm đổi thì cũng hiện lại xem trước (nội dung gửi đi đổi).
  const regions = request.sheets.map((s) => `${s.name}!${s.region}${s.context_region ? `+${s.context_region}` : ''}`).join('|');
  const masked = (preview.masked ?? []).map((g) => `${g.sheet}:${g.reason}:${(g.ranges ?? []).join(',')}`).join('|');
  return `${request.active_sheet}#${regions}#${masked}`;
}

// Dòng "Không gửi: ..." theo nhãn và vùng che, không có giá trị ô. Không có gì bị che thì câu riêng.
export function notSentLine(texts, masked) {
  if (!masked?.length) return chatText(texts, 'CHAT_UI_PREVIEW_NOT_SENT_NONE');
  const items = masked.map((g) => chatText(texts, 'CHAT_UI_PREVIEW_NOT_SENT_ITEM', { 'cột': `${g.label ? `${g.label} ` : ''}${(g.ranges ?? []).join(', ')}`, sheet: g.sheet }));
  return chatText(texts, 'CHAT_UI_PREVIEW_NOT_SENT', { 'danh sách': items.join('; ') });
}

// Giá trị ghi vào lưới cho một op. Trả { kind: 'value'|'formula'|'clear', value, text }.
export function opWrite(op) {
  if (op.op === 'clear_cell') return { kind: 'clear' };
  if (op.op === 'set_formula') return { kind: 'formula', formula: op.formula };
  const val = op.value ?? {};
  if (val.type === 'number') return { kind: 'value', value: Number(val.normalized) };
  if (val.type === 'date') return { kind: 'date', value: isoToSerial(val.normalized) };
  // as_text: chữ bắt đầu bằng = + - @ phải ở dạng chữ, không thành công thức.
  return { kind: 'text', value: String(val.normalized ?? val.raw_text ?? ''), asText: Boolean(val.as_text) };
}

// K3 (5.4): dòng lịch sử bản "Có {N} ô do AI sửa qua chat: {danh sách ô}." từ meta ai_cells ("sheet!A1"). Cùng một sheet thì
// chỉ ghi ô (B10, B11); nhiều sheet thì ghi kèm tên sheet. Ô xếp theo cột rồi hàng như trên lưới. Không có ô nào thì null.
export function historyLine(texts, aiCells) {
  if (!aiCells?.length) return null;
  const parts = aiCells.map((key) => {
    const cut = key.lastIndexOf('!');
    return { sheet: key.slice(0, cut), cell: key.slice(cut + 1) };
  });
  const order = (c) => { const m = /^([A-Z]+)(\d+)$/i.exec(c) ?? [null, c, '0']; return [m[1].length, m[1].toUpperCase(), Number(m[2])]; };
  parts.sort((a, b) => {
    if (a.sheet !== b.sheet) return a.sheet < b.sheet ? -1 : 1;
    const x = order(a.cell), y = order(b.cell);
    return x[0] - y[0] || (x[1] < y[1] ? -1 : x[1] > y[1] ? 1 : 0) || x[2] - y[2];
  });
  const oneSheet = new Set(parts.map((p) => p.sheet)).size === 1;
  const list = parts.map((p) => (oneSheet ? p.cell : `${p.sheet}!${p.cell}`)).join(', ');
  return chatText(texts, 'CHAT_UI_HISTORY', { N: parts.length, 'danh sách ô': list });
}

// Mọi mã câu khung chat có thể cần, để test đối chiếu với chat_texts.
export const UI_KEYS = [
  'CHAT_UI_LEGEND', 'CHAT_UI_INPUT_HINT', 'CHAT_UI_PREVIEW_TITLE', 'CHAT_UI_PREVIEW_NOT_SENT', 'CHAT_UI_PREVIEW_SEND', 'CHAT_UI_PREVIEW_CANCEL',
  'CHAT_UI_WAITING', 'CHAT_UI_STOP', 'CHAT_UI_STOPPED', 'CHAT_UI_EDITED', 'CHAT_UI_EDITED_HINT', 'CHAT_UI_EDITED_NONE', 'CHAT_UI_PARTIAL',
  'CHAT_UI_CELL_CHANGED', 'CHAT_UI_CELL_LOCKED', 'CHAT_CELL_HAS_TEXT', 'CHAT_UI_STALE', 'CHAT_UI_EDITED_ROW', 'CHAT_UI_EDITED_ROW_BARE',
  'CHAT_UI_FORMULA', 'CHAT_UI_EMPTY', 'CHAT_UI_PREVIEW_SHEET', 'CHAT_UI_PREVIEW_NOT_SENT_NONE', 'CHAT_UI_PREVIEW_NOT_SENT_ITEM', 'CHAT_UI_ERROR_CODE', 'CHAT_UI_MODEL_SAID', 'CHAT_UI_UNDO_TURN', 'CHAT_UI_UNDO_TURN_DISABLED', 'CHAT_UI_SAVE_CONFIRM',
  'CHAT_UI_SAVE_BUTTON', 'CHAT_UI_REVIEW_BUTTON', 'CHAT_UI_TRIAL_LABEL', 'CHAT_UI_HISTORY', 'CHAT_UI_PREVIEW_CONTEXT', 'CHAT_UI_TURN_UNDONE', 'CHAT_UI_OUTSIDE_BODY', 'CHAT_UI_PREPARING',
  'CHAT_ASK', 'CHAT_REFUSED', 'CHAT_UNSUPPORTED', 'CHAT_UNSUPPORTED_FILL_FROM_FILE', 'CHAT_UNSUPPORTED_FORMAT', 'CHAT_UNSUPPORTED_INSERT', 'CHAT_UNSUPPORTED_OTHER',
];
const UNSUPPORTED = { fill_from_file: 'CHAT_UNSUPPORTED_FILL_FROM_FILE', format: 'CHAT_UNSUPPORTED_FORMAT', insert: 'CHAT_UNSUPPORTED_INSERT', other: 'CHAT_UNSUPPORTED_OTHER' };

// Câu chính cho kết quả một lượt (trước khi áp). Áp xong thì giao diện thêm số ô đã sửa.
export function planHeadline(texts, plan) {
  if (plan.status === 'ask') return chatText(texts, 'CHAT_ASK');
  if (plan.status === 'refuse') return chatText(texts, 'CHAT_REFUSED');
  if (plan.status === 'unsupported') return chatText(texts, UNSUPPORTED[plan.unsupported] ?? 'CHAT_UNSUPPORTED');
  if (plan.status === 'rejected') return chatText(texts, plan.reason ?? 'CHAT_BAD_OUTPUT');
  return null;
}

export function rejectedLines(texts, plan) {
  return (plan.rejected_ops ?? []).map((r) => ({ cell: r.cell, code: r.reason, text: chatText(texts, r.reason, { 'ô': r.cell }) }));
}

// Lỗi cài đặt: kèm mã để người cài app tra được (AI Engineer 12/11).
export const SETUP_ERRORS = new Set(['CHAT_NO_PRICE', 'CHAT_MODEL_NO_KEY', 'CHAT_MODEL_KEY_REJECTED']);

function viDate(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso ?? '');
  return m ? `${m[3]}/${m[2]}/${m[1]}` : String(iso ?? '');
}

// Chữ của một ô trong bản chụp đã gửi (chữ thật trên lưới, Q135 (h)). fmt: viết số theo dấu thập phân của app.
export function cellText(texts, cell, fmt) {
  if (!cell || cell.v === null || cell.v === undefined || cell.v === '') {
    return cell?.f ? chatText(texts, 'CHAT_UI_FORMULA', { 'công thức': cell.f }) : chatText(texts, 'CHAT_UI_EMPTY');
  }
  if (cell.t === 'number') return fmt(Number(cell.v));
  if (cell.t === 'date') return viDate(cell.v);
  return String(cell.v);
}

// Một dòng tóm tắt do app viết (C4): ô (nhãn cột, nhãn dòng): cũ → mới.
export function summaryText(texts, line, op, sentCell, fmt) {
  const label = [line?.col_label, line?.row_label].filter((x) => x && String(x).trim()).join(', ');
  let next;
  if (op.op === 'clear_cell') next = chatText(texts, 'CHAT_UI_EMPTY');
  else if (op.op === 'set_formula') next = chatText(texts, 'CHAT_UI_FORMULA', { 'công thức': op.formula });
  else if (op.value?.type === 'number') next = fmt(Number(op.value.normalized));
  else if (op.value?.type === 'date') next = viDate(op.value.normalized);
  else next = String(op.value?.normalized ?? '');
  const vars = { 'ô': op.cell, 'nhãn': label, 'cũ': cellText(texts, sentCell, fmt), 'mới': next };
  return chatText(texts, label ? 'CHAT_UI_EDITED_ROW' : 'CHAT_UI_EDITED_ROW_BARE', vars);
}

// Q135 (h2, PM 13/11): AI không ghi đè hay xóa MỌI ô chữ đang có nội dung, dù chữ đó đã gửi cho AI hay chưa
// (ví dụ ghi chú ở cột I). Ô trống, ô số vẫn ghi được. Ô công thức đã bị cổng của vmai chặn.
export function overwritesText(op, sentCell) {
  return Boolean(sentCell && !sentCell.f && sentCell.t === 'text' && sentCell.v !== null && sentCell.v !== undefined && sentCell.v !== '');
}
