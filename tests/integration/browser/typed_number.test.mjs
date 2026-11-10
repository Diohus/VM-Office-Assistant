// Chạy: node --test typed_number.test.mjs. Không cần trình duyệt.
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseTypedNumber as p } from '../../../frontend/src/grid/typedNumber.js';
import { interpretTable } from '../../../frontend/src/grid/pastedTable.js';

test('hiểu chắc chắn', () => {
  for (const [text, value] of [['5,51', 5.51], ['5.51', 5.51], ['-5,5', -5.5], ['0', 0], ['12', 12], ['0,125', 0.125], ['1234,567', 1234.567], ['1.234,5', 1234.5], ['1,234.5', 1234.5], ['1.234.567', 1234567], ['1,5', 1.5]]) {
    assert.deepEqual(p(text), { kind: 'number', value }, text);
  }
});

test('nhập nhằng', () => {
  for (const text of ['1.234', '1,234', '12,345', '999.000', '-1.234']) assert.equal(p(text).kind, 'ambiguous', text);
});

test('không phải số', () => {
  for (const text of ['abc', '', '=1+1', '01/09', '5,5a']) assert.equal(p(text).kind, 'other', text);
});

test('vùng dán chữ thuần: suy từ ô khác trong vùng, báo ô đã suy', () => {
  const cells = (list) => list.map((t) => [{ text: t }]);
  const a = interpretTable(cells(['1,234,567', '1,234', '5.47']));
  assert.deepEqual(a.rows.flat(), [1234567, 1234, 5.47]);
  assert.equal(a.guessed.length, 1);
  assert.equal(a.decimal, '.');
  const b = interpretTable(cells(['1.234.567', '1.234', '5,47']));
  assert.deepEqual(b.rows.flat(), [1234567, 1234, 5.47]);
  assert.equal(b.decimal, ',');
  const none = interpretTable(cells(['1,234', '1.234']));
  // Không có gì để suy: theo dấu của app, dấu chấm (Q122). Vẫn báo cả hai ô.
  assert.deepEqual(none.rows.flat(), [1234, 1.234]);
  assert.equal(none.hadEvidence, false);
  assert.equal(none.guessed.length, 2);
  assert.deepEqual(interpretTable(cells(['abc', '', '=1+1'])).rows.flat(), ['abc', null, '=1+1']);
});

test('ô Excel: mẫu định dạng quyết định, kể cả Excel dùng dấu phẩy thập phân', () => {
  const read = (list) => interpretTable(list.map(([text, format]) => [{ text, format, numeric: true }]), { fromExcel: true });
  const dotExcel = read([['1,234,567', '\\#\\,\\#\\#0'], ['1,234', '\\#\\,\\#\\#0'], ['1.234', '0\\.000'], ['5.47', null], ['5.47', 'Fixed'], ['1,234.5', '\\#\\,\\#\\#0\\.0'], ['-1,234', '\\#\\,\\#\\#0'], ['50%', '0%']]);
  assert.deepEqual(dotExcel.rows.flat(), [1234567, 1234, 1.234, 5.47, 5.47, 1234.5, -1234, 0.5]);
  assert.equal(dotExcel.guessed.length, 0);
  const commaExcel = read([['1.234.567', '\\#\\,\\#\\#0'], ['1.234', '\\#\\,\\#\\#0'], ['1,234', '0\\.000'], ['5,47', null], ['1.234,5', '\\#\\,\\#\\#0\\.0']]);
  assert.deepEqual(commaExcel.rows.flat(), [1234567, 1234, 1.234, 5.47, 1234.5]);
  const unread = read([['9/1/2026', 'm/d/yyyy'], ['$1,234', '\\$\\#\\,\\#\\#0']]);
  assert.equal(unread.unread.length, 2);
  assert.deepEqual(unread.rows.flat(), ['9/1/2026', '$1,234']);
  const text = interpretTable([[{ text: '1,234', format: null, numeric: false }]], { fromExcel: true });
  assert.deepEqual(text.rows.flat(), ['1,234'], 'Excel căn trái là chữ, giữ nguyên');
});

test('chữ số lẫn lộn thì không rõ ràng, không để thư viện đoán', () => {
  for (const text of ['1,2,3', '5,5.1', '5,5511.234', '1.2.3', '1,,2', ',5', '5,', '1.234.56']) {
    assert.equal(p(text).kind, 'unclear', text);
  }
});

// Q104 lỗi chặn 1, Q113: ô Excel chỉ chắc chắn khi chữ thường của clipboard chứng minh được. Không chứng minh được thì báo làm tròn.
const xl = (list) => list.map((item) => { const [text, format = null] = Array.isArray(item) ? item : [item]; return [{ text, format, numeric: true }]; });
const plainOf = (list) => list.map((t) => [{ text: t }]);
const roundedTexts = (read) => read.rounded.map((x) => x.text);

test('ô General cột hẹp: lấy số gốc từ chữ thường khi khớp số đang hiện', () => {
  const shown = ['1235', '-2.7', '-12', '1234', '0.25', '0'];
  const plain = ['1234.9', '-2.71828', '-12.345', '1234', '0.25', '0.000123456'];
  const read = interpretTable(xl(shown), { fromExcel: true, plain: plainOf(plain) });
  assert.deepEqual(read.rows.flat(), [1234.9, -2.71828, -12.345, 1234, 0.25, 0.000123456]);
  // 0.000123456 dài 11 ký tự: đúng ngưỡng Excel có thể đã cắt, nên vẫn báo dù số đã lấy là số gốc.
  assert.deepEqual(roundedTexts(read), ['0']);
});

test('ô General: thiếu chữ thường thì mọi ô số đều báo, kể cả số nguyên', () => {
  const read = interpretTable(xl(['1235', '0', '1234']), { fromExcel: true });
  assert.deepEqual(read.rows.flat(), [1235, 0, 1234]);
  assert.deepEqual(roundedTexts(read), ['1235', '0', '1234']);
});

test('ô General: số dài từ 10 ký tự thì báo, đo trên Excel máy dự án 10/11', () => {
  const shown = ['1234567890', '123456789.1', '-1234.56789', '12345678902', '123456789', '0.00000123'];
  const read = interpretTable(xl(shown), { fromExcel: true, plain: plainOf(shown) });
  assert.deepEqual(roundedTexts(read), ['1234567890', '123456789.1', '-1234.56789', '12345678902', '0.00000123']);
  assert.equal(read.rows.flat()[4], 123456789);
});

test('ô General: chữ thường không làm tròn ra số đang hiện thì không dùng', () => {
  const read = interpretTable(xl(['1235', '5.47']), { fromExcel: true, plain: plainOf(['999', '5.4812']) });
  assert.deepEqual(read.rows.flat(), [1235, 5.47]);
  assert.deepEqual(roundedTexts(read), ['1235', '5.47']);
});

test('chữ thường lệch số dòng, số cột, hoặc có ô gộp thì bỏ, mọi ô đều báo', () => {
  const rows = xl(['1235', '1234']);
  assert.equal(interpretTable(rows, { fromExcel: true, plain: plainOf(['1234.9']) }).rounded.length, 2, 'thiếu dòng');
  assert.equal(interpretTable(rows, { fromExcel: true, plain: [[{ text: '1234.9' }, { text: '' }], [{ text: '1234' }]] }).rounded.length, 2, 'thừa cột');
  const merged = xl(['1235', '1234']);
  merged.merged = true;
  assert.equal(interpretTable(merged, { fromExcel: true, plain: plainOf(['1234.9', '1234']) }).rounded.length, 2, 'ô gộp');
});

test('Excel dấu phẩy thập phân (clipboard giả, chưa thử trên Excel thật dấu phẩy)', () => {
  const read = interpretTable(xl(['1235', '5,47', '-2,7', '0']), { fromExcel: true, plain: plainOf(['1234,9', '5,4712', '-2,71828', '0,000123']) });
  assert.deepEqual(read.rows.flat(), [1234.9, 5.4712, -2.71828, 0.000123]);
  assert.deepEqual(roundedTexts(read), []);
  // Hai dấu khác nhau giữa số đang hiện và chữ thường: không biết chắc, báo.
  const mixed = interpretTable(xl(['5,47']), { fromExcel: true, plain: plainOf(['5.4712']) });
  assert.deepEqual(mixed.rows.flat(), [5.47]);
  assert.equal(mixed.rounded.length, 1);
  // 1.234 trong chữ thường: mẫu General không có dấu nghìn nên là 1,234, và phải làm tròn ra đúng số đang hiện mới dùng.
  const amb = interpretTable(xl(['1234']), { fromExcel: true, plain: plainOf(['1.234']) });
  assert.deepEqual(amb.rows.flat(), [1234]);
  assert.equal(amb.rounded.length, 1);
});

test('ô có định dạng riêng luôn báo làm tròn: phần trăm, dấu nghìn, số lẻ cố định, số âm', () => {
  const read = interpretTable(xl([['50%', '0%'], ['1,235', '\#\,\#\#0'], ['5.47', '0.00'], ['-1,234.57', '\#\,\#\#0.00']].map(([t, f]) => [t, f])), { fromExcel: true, plain: plainOf(['50%', '1,235', '5.47', '-1,234.57']) });
  assert.deepEqual(read.rows.flat(), [0.5, 1235, 5.47, -1234.57]);
  assert.equal(read.rounded.length, 4);
});

test('ô canh riêng mẫu General: vẫn là ô phải xem lại, và cũng theo quy tắc làm tròn', () => {
  const cells = [[{ text: '5.47', format: null, numeric: false, aligned: true }], [{ text: '1235', format: null, numeric: false, aligned: true }]];
  const read = interpretTable(cells, { fromExcel: true, plain: plainOf(['5.4712', '1234.9']) });
  assert.deepEqual(read.rows.flat(), [5.4712, 1234.9]);
  assert.equal(read.guessed.length, 2);
  assert.equal(read.rounded.length, 0);
  const noPlain = interpretTable(cells, { fromExcel: true });
  assert.equal(noPlain.rounded.length, 2);
});

// Q122: dấu thập phân của app là dấu chấm (frontend/src/numberFormat.js). Gõ dấu phẩy không bao giờ thành số nghìn im lặng.
import { DECIMAL, formatNumber, usesOtherDecimal } from '../../../frontend/src/numberFormat.js';

test('Q122: cấu hình mặc định là dấu chấm', () => {
  assert.equal(DECIMAL, '.');
  assert.equal(formatNumber(3.14159265358979), '3.14159265358979');
  assert.equal(formatNumber(0.000123456), '0.000123456');
  assert.equal(formatNumber(1234567890), '1234567890');
});

test('Q122: gõ 5,5; 0,5; 12,5 là số lẻ, có báo; 1,5,5 không ghi', () => {
  for (const [text, value] of [['5,5', 5.5], ['0,5', 0.5], ['12,5', 12.5], ['-2,75', -2.75], ['1.234,5', 1234.5]]) {
    assert.deepEqual(p(text), { kind: 'number', value }, text);
    assert.equal(usesOtherDecimal(text), true, `${text} phải được báo`);
  }
  for (const text of ['5.5', '12', '1,234,567', '1.234.567']) assert.equal(usesOtherDecimal(text), false, text);
  assert.equal(p('1,5,5').kind, 'unclear');
  assert.equal(p('1,234').kind, 'ambiguous');
  assert.equal(p('1.234').kind, 'ambiguous');
});

test('Q125: dán chữ thuần báo giống gõ tay: số lẻ dấu phẩy và chữ số lẫn lộn', () => {
  const cells = (list) => list.map((t) => [{ text: t }]);
  const read = interpretTable(cells(['5,5', '0,5', '12,75', '1,5,5', '7']));
  assert.deepEqual(read.rows.flat(), [5.5, 0.5, 12.75, '1,5,5', 7]);
  assert.deepEqual(read.swapped.map((x) => [x.text, x.value]), [['5,5', 5.5], ['0,5', 0.5], ['12,75', 12.75]]);
  assert.deepEqual(read.unclear.map((x) => x.text), ['1,5,5']);
});
