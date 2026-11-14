// Chạy: node --test chat_model.test.mjs. Không cần trình duyệt. Lát 2 (Q135): khung chat, phần không đụng lưới.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {
  chatText, serialToIso, isoToSerial, isDatePattern, referencedSheets, chatCell, sameAsSent, previewKey, notSentLine, opWrite,
  planHeadline, rejectedLines, UI_KEYS, TEMP_PREFIX, summaryText, overwritesText, cellText, historyLine,
} from '../../../frontend/src/chat/chatModel.js';

const ROOT = path.resolve('..', '..', '..');
const texts = JSON.parse(fs.readFileSync(path.join(ROOT, 'vmai', 'src', 'vmai', 'data', 'doc_codes.json'), 'utf8'));

// PM 12/11: nghiệm thu chỉ bắt đầu khi mọi mã có câu thật. Quét mã CHAT_* ở mọi nơi có thể trả về giao diện.
const SOURCES = ['vmai/src/vmai/chat.py', 'backend/app/main.py', 'backend/app/chat.py', 'backend/app/reader_client.py', 'backend/app/llm_client.py', 'workers/llm/llm_service.py'];
// Không phải mã trả về: biến môi trường, hằng phiên bản, khóa câu chung nội bộ của api (_CHAT_TEXT).
const NOT_CODES = new Set(['CHAT_MAX_USD_PER_DAY', 'CHAT_MAX_USD_PER_TURN', 'CHAT_TEST_NOW', 'CHAT_SCHEMA_VERSION', 'CHAT_MODEL']);

test('mọi mã CHAT_ và mọi câu giao diện đều có câu thật trong chat_texts', () => {
  const codes = new Set();
  for (const file of SOURCES) {
    for (const m of fs.readFileSync(path.join(ROOT, file), 'utf8').matchAll(/\bCHAT_[A-Z_]*[A-Z]\b/g)) if (!NOT_CODES.has(m[0])) codes.add(m[0]);
  }
  assert.ok(codes.size > 40, 'quét được quá ít mã, có thể đường dẫn sai');
  const missing = [...codes, ...UI_KEYS].filter((c) => typeof texts.chat_texts?.[c] !== 'string' || !texts.chat_texts[c].trim());
  assert.deepEqual(missing, []);
  for (const c of [...codes, ...UI_KEYS]) assert.ok(!chatText(texts, c).startsWith(TEMP_PREFIX), c);
});

test('câu điền chỗ trống, mã lạ thì ra câu tạm có đánh dấu', () => {
  assert.equal(chatText({ chat_texts: { X: 'Ô {ô} có {N} lần, {khác}' } }, 'X', { 'ô': 'B10', N: 2 }), 'Ô B10 có 2 lần, {khác}');
  assert.equal(chatText(texts, 'KHONG_CO_MA'), `${TEMP_PREFIX} KHONG_CO_MA`);
});

test('số ngày Excel hệ 1900, kể cả ngày 29/02/1900 không có thật', () => {
  assert.equal(serialToIso(1), '1900-01-01');
  assert.equal(serialToIso(59), '1900-02-28');
  assert.equal(serialToIso(60), null);
  assert.equal(serialToIso(61), '1900-03-01');
  assert.equal(serialToIso(46266), '2026-09-01');
  assert.equal(serialToIso(46266.75), '2026-09-01');
  for (const iso of ['1900-01-01', '1900-02-28', '1900-03-01', '2026-09-01', '2026-12-31']) assert.equal(serialToIso(isoToSerial(iso)), iso);
  assert.equal(isoToSerial('2026-02-30'), null);
  assert.equal(isoToSerial('01/09/2026'), null);
});

test('mẫu ngày', () => {
  for (const p of ['dd/mm/yyyy', 'd/m/yy', 'mmm-yy', '[$-42A]dd/mm/yyyy', 'yyyy-mm-dd']) assert.equal(isDatePattern(p), true, p);
  for (const p of ['General', '0.00', '#,##0', 'hh:mm', '0.0"d"', '[Red]0.00', '', null]) assert.equal(isDatePattern(p), false, String(p));
});

test('sheet được công thức tham chiếu', () => {
  const names = ['KQ_QC_Thang', 'GioiHan', 'Giới hạn 2'];
  assert.deepEqual(referencedSheets('=IF(B10>GioiHan!C3,"Cao","")', names), ['GioiHan']);
  assert.deepEqual(referencedSheets("=SUM('Giới hạn 2'!A1:A3)+gioihan!B2", names), ['Giới hạn 2', 'GioiHan']);
  assert.deepEqual(referencedSheets('="GioiHan!A1"', names), []);
  assert.deepEqual(referencedSheets('=SUM(B8:B37)', names), []);
  assert.deepEqual(referencedSheets('=Khac!A1', names), []);
});

test('ô gửi đi', () => {
  assert.equal(chatCell('B10', { v: null }), null);
  assert.deepEqual(chatCell('B10', { v: 5.62, display: '5,62' }), { a1: 'B10', t: 'number', v: '5.62', f: null, display: '5,62' });
  assert.deepEqual(chatCell('A8', { v: 46266, pattern: 'dd/mm/yyyy' }), { a1: 'A8', t: 'date', v: '2026-09-01', f: null, display: null });
  assert.deepEqual(chatCell('C1', { v: true }), { a1: 'C1', t: 'boolean', v: 'true', f: null, display: null });
  assert.deepEqual(chatCell('B40', { v: 2.5, f: 'STDEV(B8:B37)' }), { a1: 'B40', t: 'number', v: '2.5', f: '=STDEV(B8:B37)', display: null });
  assert.deepEqual(chatCell('B41', { v: '', f: '=IFERROR(1/0,"")' }), { a1: 'B41', t: 'text', v: '', f: '=IFERROR(1/0,"")', display: null });
  // O2: công thức chưa có kết quả không được gửi v null (vmai từ chối cả lượt với CHAT_BAD_INPUT).
  assert.deepEqual(chatCell('H40', { v: null, f: 'B10*2' }), { a1: 'H40', t: 'text', v: '', f: '=B10*2', display: null });
  assert.deepEqual(chatCell('I8', { v: 'chạy lại' }), { a1: 'I8', t: 'text', v: 'chạy lại', f: null, display: null });
});

test('ô còn như lúc gửi', () => {
  const sent = { a1: 'B10', t: 'number', v: '5.5', f: null };
  assert.equal(sameAsSent(sent, { ...sent }), true);
  assert.equal(sameAsSent(sent, { ...sent, v: '5.6' }), false);
  assert.equal(sameAsSent(sent, null), false);
  assert.equal(sameAsSent(null, null), true);
  assert.equal(sameAsSent(null, sent), false);
});

test('hiện lại xem trước khi vùng hoặc ô che đổi', () => {
  const req = { active_sheet: 'KQ', sheets: [{ name: 'KQ', region: 'A1:I47' }] };
  const prev = { masked: [{ sheet: 'KQ', reason: 'person', label: 'Người chạy', ranges: ['H8:H37'] }] };
  const k = previewKey(req, prev);
  assert.equal(previewKey(req, prev), k);
  assert.notEqual(previewKey({ ...req, sheets: [{ name: 'KQ', region: 'A1:C10' }] }, prev), k);
  assert.notEqual(previewKey({ ...req, sheets: [...req.sheets, { name: 'GioiHan', region: 'A1:M7' }] }, prev), k);
  assert.notEqual(previewKey(req, { masked: [] }), k);
  const T = texts.chat_texts;
  const item = T.CHAT_UI_PREVIEW_NOT_SENT_ITEM.replace('{cột}', 'Người chạy H8:H37').replace('{sheet}', 'KQ');
  assert.equal(notSentLine(texts, prev.masked), T.CHAT_UI_PREVIEW_NOT_SENT.replace('{danh sách}', item));
  assert.equal(notSentLine(texts, []), T.CHAT_UI_PREVIEW_NOT_SENT_NONE);
});

test('ghi op lên lưới', () => {
  assert.deepEqual(opWrite({ op: 'set_value', value: { type: 'number', normalized: '5.62' } }), { kind: 'value', value: 5.62 });
  assert.deepEqual(opWrite({ op: 'set_value', value: { type: 'date', normalized: '2026-09-01' } }), { kind: 'date', value: 46266 });
  assert.deepEqual(opWrite({ op: 'set_value', value: { type: 'text', normalized: '=1+1', as_text: true } }), { kind: 'text', value: '=1+1', asText: true });
  assert.deepEqual(opWrite({ op: 'set_formula', formula: '=AVERAGE(B8:B37)' }), { kind: 'formula', formula: '=AVERAGE(B8:B37)' });
  assert.deepEqual(opWrite({ op: 'clear_cell' }), { kind: 'clear' });
});

test('câu kết quả cả lượt và op bị bỏ', () => {
  assert.equal(planHeadline(texts, { status: 'ask' }), texts.chat_texts.CHAT_ASK);
  assert.equal(planHeadline(texts, { status: 'unsupported', unsupported: 'fill_from_file' }), texts.chat_texts.CHAT_UNSUPPORTED_FILL_FROM_FILE);
  assert.equal(planHeadline(texts, { status: 'rejected', reason: 'CHAT_TOO_MANY_OPS' }), texts.chat_texts.CHAT_TOO_MANY_OPS);
  assert.equal(planHeadline(texts, { status: 'edit' }), null);
  const lines = rejectedLines(texts, { rejected_ops: [{ index: 1, op: 'set_value', cell: 'H10', reason: 'CHAT_CELL_PERSON' }] });
  assert.equal(lines[0].text, texts.chat_texts.CHAT_CELL_PERSON.replace('{ô}', 'H10'));
});

test('dòng tóm tắt do app viết: chữ cũ thật, số theo dấu của app', () => {
  const T = texts.chat_texts;
  const comma = (v) => String(v).replace('.', ',');
  const line = { cell: 'B10', col_label: 'Glucose (mmol/L) · Mức 1', row_label: '03/09/2026' };
  const op = { op: 'set_value', cell: 'B10', value: { type: 'number', normalized: '5.62' } };
  const want = T.CHAT_UI_EDITED_ROW.replace('{ô}', 'B10').replace('{nhãn}', 'Glucose (mmol/L) · Mức 1, 03/09/2026').replace('{cũ}', '5,52').replace('{mới}', '5,62');
  assert.equal(summaryText(texts, line, op, { a1: 'B10', t: 'number', v: '5.52', f: null }, comma), want);
  const bare = summaryText(texts, { cell: 'I9', col_label: '', row_label: null }, { op: 'set_value', cell: 'I9', value: { type: 'text', normalized: 'chạy lại' } }, null, comma);
  assert.equal(bare, T.CHAT_UI_EDITED_ROW_BARE.replace('{ô}', 'I9').replace('{cũ}', T.CHAT_UI_EMPTY).replace('{mới}', 'chạy lại'));
  const f = summaryText(texts, null, { op: 'set_formula', cell: 'I8', formula: '=IF(B8>1,"x","")' }, { a1: 'I8', t: 'text', v: 'ghi chú cũ', f: null }, comma);
  assert.ok(f.includes('ghi chú cũ') && f.includes(T.CHAT_UI_FORMULA.replace('{công thức}', '=IF(B8>1,"x","")')), f);
  assert.ok(!f.includes(T.CHAT_UI_OLD_TEXT), 'Q135 (h): chữ cũ thật, không ghi "chữ cũ"');
  assert.equal(cellText(texts, { t: 'date', v: '2026-09-03' }, comma), '03/09/2026');
});

test('Q135 (h2): không ghi đè, không xóa mọi ô chữ đang có nội dung; ô trống, ô số vẫn ghi', () => {
  const op = { op: 'set_value', cell: 'I10' };
  assert.equal(overwritesText(op, { a1: 'I10', t: 'text', v: 'ghi chú dài chưa gửi', f: null }), true);
  assert.equal(overwritesText(op, { a1: 'I10', t: 'text', v: 'ghi chú ngắn', f: null }), true, 'chữ đã gửi cũng giữ');
  assert.equal(overwritesText({ op: 'clear_cell', cell: 'I10' }, { a1: 'I10', t: 'text', v: 'x', f: null }), true, 'xóa cũng không');
  assert.equal(overwritesText(op, null), false);
  assert.equal(overwritesText(op, { a1: 'I10', t: 'text', v: '', f: null }), false);
  assert.equal(overwritesText(op, { a1: 'B10', t: 'number', v: '5.5', f: null }), false);
  assert.equal(overwritesText(op, { a1: 'B40', t: 'text', v: '', f: '=IFERROR(STDEV(B8:B37),"")' }), false);
});

test('K3: lịch sử bản nêu ô nào do AI sửa', () => {
  assert.equal(historyLine(texts, []), null);
  assert.equal(historyLine(texts, undefined), null);
  const one = historyLine(texts, ['KQ_QC_Thang!B11', 'KQ_QC_Thang!I9', 'KQ_QC_Thang!B10', 'KQ_QC_Thang!AA2']);
  assert.equal(one, texts.chat_texts.CHAT_UI_HISTORY.replace('{N}', '4').replace('{danh sách ô}', 'B10, B11, I9, AA2'));
  const two = historyLine(texts, ['KQ_QC_Thang!B10', 'GioiHan!C3']);
  assert.ok(two.includes('GioiHan!C3, KQ_QC_Thang!B10'), two);
  assert.ok(!two.startsWith(TEMP_PREFIX));
});

test('0.2: phần đọc kèm đổi thì khóa xem trước đổi', () => {
  const pv = { masked: [] };
  const a = previewKey({ active_sheet: 'S', sheets: [{ name: 'S', region: 'B8:C37' }] }, pv);
  const b = previewKey({ active_sheet: 'S', sheets: [{ name: 'S', region: 'B8:C37', context_region: 'A1:C37' }] }, pv);
  assert.notEqual(a, b);
});
