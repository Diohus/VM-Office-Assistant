// Lát 1 nối AI (Q133): dựng nội dung khung duyệt từ lô đã lưu, và biến lựa chọn của người dùng thành quyết định gửi lên.
// Không dùng React, để test bằng node. Câu chữ (texts) là vmai/src/vmai/data/doc_codes.json, truyền vào; không chép tay câu nào.
// Trình duyệt không bao giờ gửi số đề xuất: chỉ change_id và lựa chọn (acceptance_lat1 4.6). Số tự gõ ở ô đỏ là số của người dùng (D4).
import { parseTypedNumber } from '../grid/typedNumber.js';

export const LAT1_HEADER_KEY = 'Màn duyệt, bản thử đầu (lát 1)';

// Tên các {chỗ trống} trong câu.
export function slotNames(text) {
  return [...String(text ?? '').matchAll(/\{([^}]+)\}/g)].map((m) => m[1]);
}

// Điền {chỗ trống}. Chỗ không có giá trị thì bỏ, không để lộ dấu ngoặc; nơi gọi phải tránh trường hợp này (xem slotText).
export function fillSlots(text, values = {}) {
  let out = String(text ?? '');
  for (const [k, v] of Object.entries(values)) if (v != null && v !== '') out = out.split(`{${k}}`).join(v);
  return out.replace(/\s*:?\s*\{[^}]*\}/g, '').trim();
}

// Ngày ISO kiểu Việt: 2026-09-08 → 08/09 (ngắn) hoặc 08/09/2026; tháng 2026-10 → 10/2026. Chữ khác giữ nguyên.
export function viDate(text, { year = false } = {}) {
  const s = String(text);
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return year ? `${m[3]}/${m[2]}/${m[1]}` : `${m[3]}/${m[2]}`;
  m = s.match(/^(\d{4})-(\d{2})$/);
  if (m) return `${m[2]}/${m[1]}`;
  return s;
}

// Giá trị cho chỗ trống của một lời báo (Notice 0.4.1: items là ngày, tháng hoặc tên cột; count là số dòng).
function noticeValues(n, base) {
  const list = (n.items ?? []).map((x) => viDate(x)).join(', ');
  return { ...base, 'tên': n.detail, n: n.count == null ? null : String(n.count), 'ngày': list, 'cột': list, 'tháng nguồn': list };
}

// Câu theo mã, đủ mọi chỗ trống. Thiếu giá trị cho chỗ trống nào (lô cũ không có items) thì dùng câu đầy đủ trong detail,
// để không mất ngày, tên cột... (PO/BA K2, N2, N3).
function slotText(text, values, detail) {
  if (!text) return '';
  const missing = slotNames(text).some((k) => values[k] == null || values[k] === '');
  return missing && detail ? String(detail) : fillSlots(text, values);
}

const pick = (a, b, code) => (a?.[code] !== undefined ? a[code] : b?.[code]);

export function reasonText(texts, code, values, detail) {
  return slotText(pick(texts.texts_lat1, texts.reason_texts, code) ?? '', values, detail);
}

export function actionText(texts, code, values) {
  return fillSlots(pick(texts.actions_lat1, texts.actions, code) ?? '', values);
}

const isOverwrite = (c) => c.previous?.value != null || c.previous?.formula != null;
const colLabel = (c) => (c.target.col_path ?? []).join(' / ');
const rowLabel = (c) => (c.target.row_key ?? []).map((p) => viDate(p.value, { year: true })).join(' · ') || c.target.a1_at_write;

// Nội dung khung duyệt. Chỉ ô còn chờ; ô đã ghi, đã bỏ, đã để trống không hiện lại.
export function buildView(fill, texts) {
  const name = fill.source_name;
  const values = { 'tên file': name };
  const pending = fill.batch.changes.filter((c) => fill.decisions?.[c.change_id]?.state === 'pending');
  const cell = (c) => ({
    id: c.change_id,
    a1: c.target.a1_at_write,
    col: colLabel(c),
    color: c.confidence.final_color,
    value: c.value?.normalized ?? null,
    source: (c.evidence ?? []).map((e) => e.read_text),
    readings: (c.evidence ?? []).map((e, i) => ({ index: i, text: e.read_text, normalized: e.normalized })).filter((e) => e.normalized != null),
    overwrite: isOverwrite(c),
    previous: c.previous?.value?.normalized ?? (c.previous?.formula ?? null),
    reasons: c.confidence.reasons.map((r) => ({ code: r.code, text: reasonText(texts, r.code, noticeValues(r, values), r.detail) })).filter((r) => r.text),
  });
  // Câu lý do chung cho mọi ô vàng còn chờ: hiện một lần ở đầu khung, ô vàng chỉ hiện câu riêng; ô đỏ giữ đủ câu (N1).
  const yellows = pending.map(cell).filter((c) => c.color === 'yellow');
  const common = yellows.length ? yellows[0].reasons.filter((r) => yellows.every((c) => c.reasons.some((x) => x.code === r.code && x.text === r.text))) : [];
  const commonCodes = new Set(common.map((r) => r.code));
  const rows = [];
  const byRow = new Map();
  for (const c of pending) {
    const key = rowLabel(c);
    if (!byRow.has(key)) {
      byRow.set(key, { key, cells: [] });
      rows.push(byRow.get(key));
    }
    const one = cell(c);
    byRow.get(key).cells.push(one.color === 'yellow' ? { ...one, reasons: one.reasons.filter((r) => !commonCodes.has(r.code)) } : one);
  }
  // Ô ghi đè: mục riêng, duyệt từng ô (D6, acceptance_lat1 5.1, 5.2).
  const overwrites = rows.flatMap((r) => r.cells.filter((x) => x.overwrite).map((x) => ({ ...x, row: r.key })));
  for (const r of rows) r.cells = r.cells.filter((x) => !x.overwrite);
  const mapping = new Map();
  for (const c of fill.batch.changes) {
    const key = colLabel(c);
    if (!mapping.has(key)) mapping.set(key, { target: key, source: c.mapping?.source_header ?? '', samples: [] });
    const m = mapping.get(key);
    if (m.samples.length < 3 && c.value) m.samples.push(c.value.raw_text);
  }
  const notices = (fill.batch.notices ?? []).map((n) => ({
    code: n.reason,
    count: n.count,
    text: reasonText(texts, n.reason, noticeValues(n, values), n.detail),
    action: actionText(texts, n.reason, noticeValues(n, values)),
  }));
  return {
    sourceName: name,
    empty: fill.batch.changes.length === 0,
    pendingCount: pending.length,
    header: fillSlots(texts.ui_texts?.[LAT1_HEADER_KEY] ?? '', values),
    counts: {
      yellow: pending.filter((c) => c.confidence.final_color === 'yellow').length,
      red: pending.filter((c) => c.confidence.final_color === 'red').length,
    },
    notices,
    common,
    mapping: [...mapping.values()],
    rows: rows.filter((r) => r.cells.length),
    overwrites,
  };
}

// Nhận cả dòng: chỉ ô vàng không ghi đè (D3, 5.2). Bấm lại khi cả dòng đã nhận thì bỏ nhận (3.10).
export function toggleRow(row, choices) {
  const yellow = row.cells.filter((x) => x.color === 'yellow' && !x.overwrite);
  const next = { ...choices };
  const allTaken = yellow.length > 0 && yellow.every((x) => next[x.id]?.action === 'accept');
  for (const x of yellow) {
    if (allTaken) delete next[x.id];
    else next[x.id] = { action: 'accept' };
  }
  return next;
}

export function rejectRow(row, choices) {
  const next = { ...choices };
  for (const x of row.cells.filter((c) => c.color === 'yellow' && !c.overwrite)) next[x.id] = { action: 'reject' };
  return next;
}

// Ô đỏ tự gõ: hiểu như gõ trong lưới; nhập nhằng hay không rõ ràng thì không nhận (acceptance_lat1 điểm 7 PO/BA).
export function typedChoice(text) {
  const parsed = parseTypedNumber(text);
  if (parsed.kind === 'number') return { ok: true, choice: { action: 'red_typed', text: String(text).trim() } };
  return { ok: false, kind: parsed.kind };
}

// Số ô sẽ được ghi khi bấm Ghi: mọi lựa chọn có ghi số (không tính Bỏ, Để trống).
export function writeCount(choices) {
  return Object.values(choices).filter((c) => ['accept', 'red_accept', 'red_choice', 'red_typed'].includes(c.action)).length;
}

// Quyết định gửi lên API. Không có trường số nào: chỉ change_id, hành động, chỉ số lần đọc, chữ người dùng gõ, số đã thấy (3.8a).
export function toDecisions(choices) {
  return Object.entries(choices).map(([changeId, c]) => {
    const d = { change_id: changeId, action: c.action };
    if (c.action === 'red_choice') d.index = c.index;
    if (c.action === 'red_typed') d.text = c.text;
    if (c.seen !== undefined) d.seen = c.seen;
    return d;
  });
}
