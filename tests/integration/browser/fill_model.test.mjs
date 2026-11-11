// Chạy: node --test fill_model.test.mjs. Không cần trình duyệt. Lát 1 (Q133): nội dung khung duyệt và quyết định gửi lên.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { buildView, toggleRow, rejectRow, typedChoice, writeCount, toDecisions, fillSlots, slotNames, viDate, LAT1_HEADER_KEY } from '../../../frontend/src/fill/fillModel.js';

const texts = JSON.parse(fs.readFileSync(path.resolve('..', '..', '..', 'vmai', 'src', 'vmai', 'data', 'doc_codes.json'), 'utf8'));

const change = (id, a1, day, col, color, { value = '5.51', previous = null, evidence = null, reasons = ['READ_XLSX_DIRECT'] } = {}) => ({
  change_id: id,
  target: { sheet: 'KQ', a1_at_write: a1, row_key: [{ value: day }], col_path: col },
  previous: { value: previous === null ? null : { normalized: previous }, formula: null },
  value: value === null ? null : { normalized: value, raw_text: value },
  evidence: evidence ?? [{ read_text: value, normalized: null }],
  mapping: { source_header: 'Kết quả' },
  confidence: { final_color: color, reasons: reasons.map((code) => ({ code, detail: 'chi tiết' })) },
});

const batch = {
  source_name: 'GIA-NGUON.xlsx',
  batch: {
    changes: [
      change('c1', 'B8', '2026-09-01', ['Glucose', 'Mức 1'], 'yellow'),
      change('c2', 'C8', '2026-09-01', ['Glucose', 'Mức 2'], 'yellow'),
      change('c3', 'D8', '2026-09-01', ['Chol', 'Mức 1'], 'red', { value: null, evidence: [{ read_text: '5.64', normalized: '5.64' }, { read_text: '5.94', normalized: '5.94' }], reasons: ['SOURCE_DUPLICATE_RESULT'] }),
      change('c4', 'E8', '2026-09-01', ['Chol', 'Mức 2'], 'yellow', { previous: '6.00' }),
      change('c5', 'B9', '2026-09-02', ['Glucose', 'Mức 1'], 'yellow'),
    ],
    notices: [{ reason: 'MAP_TEST_UNKNOWN', count: 1, detail: 'Đường máu (giả)' }],
  },
  decisions: { c1: { state: 'pending' }, c2: { state: 'pending' }, c3: { state: 'pending' }, c4: { state: 'pending' }, c5: { state: 'written' } },
};

test('khung duyệt: câu chung, chỉ ô còn chờ, ô ghi đè nằm riêng', () => {
  const v = buildView(batch, texts);
  assert.ok(v.header.includes('GIA-NGUON.xlsx') && !v.header.includes('{'), 'câu chung lát 1 có tên file, không lộ chỗ trống');
  assert.equal(v.header, fillSlots(texts.ui_texts[LAT1_HEADER_KEY], { 'tên file': 'GIA-NGUON.xlsx' }));
  assert.deepEqual(v.counts, { yellow: 3, red: 1 });
  assert.equal(v.rows.length, 1, 'c5 đã ghi, không hiện lại');
  assert.deepEqual(v.rows[0].cells.map((c) => c.id), ['c1', 'c2', 'c3']);
  assert.deepEqual(v.overwrites.map((c) => [c.id, c.previous, c.value]), [['c4', '6.00', '5.51']]);
  const red = v.rows[0].cells.find((c) => c.id === 'c3');
  assert.deepEqual(red.readings.map((r) => r.text), ['5.64', '5.94']);
  assert.ok(v.notices[0].text.includes('Đường máu (giả)') && !v.notices[0].text.includes('{'));
});

test('nhận dòng chỉ lấy ô vàng không ghi đè; bấm lại thì bỏ nhận', () => {
  const v = buildView(batch, texts);
  let choices = toggleRow(v.rows[0], {});
  assert.deepEqual(Object.keys(choices).sort(), ['c1', 'c2'], 'không nhận ô đỏ c3, không nhận ô ghi đè c4');
  assert.equal(writeCount(choices), 2);
  choices = toggleRow(v.rows[0], choices);
  assert.deepEqual(choices, {});
  choices = rejectRow(v.rows[0], {});
  assert.equal(writeCount(choices), 0);
});

test('quyết định gửi lên không có số đề xuất', () => {
  const decisions = toDecisions({ c1: { action: 'accept' }, c3: { action: 'red_choice', index: 1 }, c4: { action: 'accept', seen: '6.00' } });
  assert.deepEqual(decisions, [
    { change_id: 'c1', action: 'accept' },
    { change_id: 'c3', action: 'red_choice', index: 1 },
    { change_id: 'c4', action: 'accept', seen: '6.00' },
  ]);
  for (const d of decisions) assert.ok(!('value' in d));
});

test('ô đỏ tự gõ: nhập nhằng và không rõ ràng thì không nhận', () => {
  assert.deepEqual(typedChoice('5.5'), { ok: true, choice: { action: 'red_typed', text: '5.5' } });
  assert.deepEqual(typedChoice('1,234'), { ok: false, kind: 'ambiguous' });
  assert.deepEqual(typedChoice('1,5,5'), { ok: false, kind: 'unclear' });
  assert.equal(typedChoice('abc').ok, false);
});

test('câu lý do chung gom theo ô vàng; ô đỏ giữ đủ câu (N1)', () => {
  const v = buildView(batch, texts);
  assert.deepEqual(v.common.map((r) => r.code), ['READ_XLSX_DIRECT'], 'ô đỏ c3 không chặn việc gom câu của ô vàng');
  const cells = v.rows[0].cells;
  assert.ok(cells.filter((c) => c.color === 'yellow').every((c) => !c.reasons.some((r) => r.code === 'READ_XLSX_DIRECT')));
  assert.deepEqual(cells.find((c) => c.id === 'c3').reasons.map((r) => r.code), ['SOURCE_DUPLICATE_RESULT'], 'ô đỏ hiện đủ câu của nó');
});

test('tiêu đề dòng viết ngày 01/09/2026 (N4)', () => {
  assert.equal(buildView(batch, texts).rows[0].key, '01/09/2026');
  assert.equal(viDate('2026-09-08'), '08/09');
  assert.equal(viDate('2026-10'), '10/2026');
  assert.equal(viDate('Glucose'), 'Glucose');
});

// K2, K3, N2, N3: mọi chỗ trống trong câu lát 1 được điền, không mất giá trị.
const withNotices = (notices, version = '0.4.1') => ({ ...batch, batch: { ...batch.batch, schema_version: version, notices } });
const noticeOf = (b, code) => buildView(b, texts).notices.find((n) => n.code === code);

test('K2: ngày thiếu điền đủ, viết 08/09', () => {
  const n = noticeOf(withNotices([{ reason: 'SOURCE_MISSING_DAYS', count: 3, detail: 'Nguồn thiếu: 2026-09-08, 2026-09-18, 2026-09-20.', items: ['2026-09-08', '2026-09-18', '2026-09-20'] }]), 'SOURCE_MISSING_DAYS');
  assert.ok(n.text.includes('08/09, 18/09, 20/09'), n.text);
  assert.ok(!n.text.includes('{') && !n.text.includes('2026-09'), n.text);
});

test('K3: cột người bệnh dùng câu lát 1, có tên cột, không có vế gửi ra ngoài', () => {
  const n = noticeOf(withNotices([{ reason: 'PHI_COLUMN_SKIPPED', count: 2, detail: 'x', items: ['Họ tên BN', 'Mã BN'] }]), 'PHI_COLUMN_SKIPPED');
  assert.ok(n.text.includes('Họ tên BN, Mã BN'), n.text);
  assert.ok(!n.text.includes('gửi ra ngoài') && !n.text.includes('{'), n.text);
});

test('N2, N3: câu sai tháng có số dòng và tháng; câu Người chạy có tên cột một lần', () => {
  const b = withNotices([
    { reason: 'SOURCE_WRONG_MONTH', count: 4, detail: 'x', items: ['2026-10'] },
    { reason: 'PERSON_COLUMN_LEFT_BLANK', count: 1, detail: 'x', items: ['Người chạy'] },
  ]);
  const wrong = noticeOf(b, 'SOURCE_WRONG_MONTH');
  assert.ok(wrong.text.includes('4') && wrong.text.includes('10/2026') && !wrong.text.includes('{'), wrong.text);
  const person = noticeOf(b, 'PERSON_COLUMN_LEFT_BLANK');
  assert.ok(person.text.includes('Người chạy') && person.text.split('Người chạy').length === 2 && !person.text.includes('{'), person.text);
});

test('mọi câu lát 1 của lời báo có chỗ trống điền được từ items, count, detail', () => {
  const known = new Set(['tên file', 'tên', 'n', 'ngày', 'cột', 'tháng nguồn']);
  for (const [code, text] of Object.entries(texts.texts_lat1 ?? {})) {
    for (const k of slotNames(text)) assert.ok(known.has(k), `${code}: chỗ trống {${k}} chưa có giá trị`);
  }
});

test('lô 0.4 cũ không có items: dùng câu đầy đủ trong detail, không mất ngày', () => {
  const n = noticeOf(withNotices([{ reason: 'SOURCE_MISSING_DAYS', count: 1, detail: 'Nguồn thiếu: 2026-09-08. Để trống, không tự suy số.' }], '0.4'), 'SOURCE_MISSING_DAYS');
  assert.ok(n.text.includes('2026-09-08'), n.text);
});
