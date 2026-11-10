// Hiểu vùng dán từ ngoài app (US-1.2, Q80). Clipboard của Excel chỉ có chữ đang hiện, theo dấu của máy nguồn, kèm mẫu định dạng ô.
// Mẫu định dạng luôn viết theo một kiểu nên cho biết chắc dấu nào là dấu nghìn. Không có mẫu thì mới phải suy, và phải báo.
import { parseTypedNumber } from './typedNumber.js';
import { DECIMAL, usesOtherDecimal } from '../numberFormat.js';

// Trả về hình dạng của mẫu số, hoặc null nếu mẫu không phải số thường (ngày, giờ, chữ).
export function formatShape(format) {
  if (!format || /^general$/i.test(format)) return { general: true };
  if (/^standard$/i.test(format)) return { group: true, decimal: true };
  if (/^(fixed|percent)$/i.test(format)) return { group: false, decimal: true };
  const plain = format.split(';')[0].replace(/\\/g, '').replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  if (/[dmyhs@]/i.test(plain) || !/[0#]/.test(plain)) return null;
  return { group: /[#0],[#0]/.test(plain), decimal: /\.[0#]/.test(plain) };
}

function byShape(text, shape) {
  const percent = text.endsWith('%');
  const negative = /^\(.*\)$/.test(text) || text.startsWith('-');
  const body = text.replace(/[()%\s -]/g, '');
  if (!/^\d[\d.,]*$/.test(body)) return null;
  const seps = body.match(/[.,]/g) ?? [];
  let digits;
  if (shape.general || (shape.decimal && !shape.group)) {
    // Mẫu General không bao giờ hiện dấu nghìn, mẫu không có nhóm cũng vậy.
    if (seps.length > 1) return null;
    digits = body.replace(/[.,]/, '.');
  } else if (shape.group && shape.decimal) {
    const at = Math.max(body.lastIndexOf('.'), body.lastIndexOf(','));
    digits = at < 0 ? body : `${body.slice(0, at).replace(/[.,]/g, '')}.${body.slice(at + 1)}`;
  } else {
    digits = body.replace(/[.,]/g, '');
  }
  const value = Number(digits);
  if (!Number.isFinite(value)) return null;
  const signed = negative ? -value : value;
  return percent ? signed / 100 : signed;
}

// Chữ thường của mẫu General dài tối đa 11 ký tự, không tính dấu trừ. Số dài hơn bị làm tròn, có khi còn 10 ký tự.
// Đo bằng Excel trên máy dự án ngày 10/11, dấu chấm thập phân: 12345678901,5 ra 12345678902; 1234567890,12 ra 1234567890;
// 0,000123456789 ra 0.000123457; -1234,56789012 ra -1234.56789. Còn 0,00000123 (10 ký tự) thì giữ đúng, nhưng không phân biệt được nên vẫn báo.
const GENERAL_SAFE_CHARS = 10;

// Số đang hiện có bao nhiêu chữ số lẻ, để kiểm số trong chữ thường có làm tròn ra đúng số đang hiện không.
function shownDecimals(text) {
  const m = text.replace(/[()%\s-]/g, '').match(/[.,](\d+)$/);
  return m ? m[1].length : 0;
}

// Ô General: HTML chỉ có số đang hiện, Excel làm tròn cho vừa cột (đo 10/11: 1234,9 hiện 1235, 0,000123456 hiện 0).
// Chữ thường của clipboard có số theo mẫu General không giới hạn cột. Chỉ dùng khi nó làm tròn ra đúng số đang hiện.
// Mẫu General không bao giờ có dấu nghìn, nên một dấu duy nhất là dấu thập phân, theo vùng ngôn ngữ của Excel (5.4712 hoặc 5,4712).
// Chữ thường và số đang hiện cùng có dấu mà khác nhau thì không chắc là cùng một kiểu, nên coi là có thể đã làm tròn.
function generalValue(shown, text, plain) {
  const body = (plain ?? '').trim();
  const full = body ? byShape(body, { general: true }) : null;
  if (full === null) return { value: shown, rounded: true };
  const sepShown = text.match(/[.,]/)?.[0];
  const sepPlain = body.match(/[.,]/)?.[0];
  if (sepShown && sepPlain && sepShown !== sepPlain) return { value: shown, rounded: true };
  const k = shownDecimals(text);
  const back = Math.round(full * 10 ** k) / 10 ** k;
  if (Math.abs(back - shown) > 1e-9 * Math.max(1, Math.abs(shown))) return { value: shown, rounded: true };
  return { value: full, rounded: body.replace(/^-/, '').length >= GENERAL_SAFE_CHARS };
}

// Chữ thường chỉ dùng được khi khớp từng ô với bảng HTML: cùng số dòng, cùng số cột mỗi dòng, không có ô gộp.
function plainMatches(rows, plain) {
  if (!plain || rows.merged || plain.length !== rows.length) return false;
  return rows.every((row, r) => plain[r].length === row.length);
}

// Ô từ Excel. numeric: Excel tự căn phải, tức ô số. aligned: người dùng tự canh ô, nên không biết là số hay chữ.
// plain: chữ của cùng ô trong phần chữ thường của clipboard, nếu có.
// Trả về { value, rounded, guessed? }, { unread: true }, hoặc null nếu là chữ.
function fromExcelCell(cell, plain) {
  const text = cell.text.trim();
  const shape = formatShape(cell.format);
  const guessed = Boolean(cell.aligned && !cell.numeric);
  if (!guessed && !cell.numeric) return null;
  const shown = shape ? byShape(text, shape) : null;
  if (shown === null) return guessed ? null : { unread: true };
  // Mẫu có định dạng riêng: Excel chỉ cho đọc số đang hiện, nên luôn có thể đã làm tròn.
  const read = shape.general ? generalValue(shown, text, plain) : { value: shown, rounded: true };
  return guessed ? { ...read, guessed: true } : read;
}

// Dấu nào là dấu thập phân, suy từ các ô không nhập nhằng trong cùng vùng. null nếu không có hoặc mâu thuẫn.
function decimalFrom(texts) {
  let dot = false;
  let comma = false;
  for (const text of texts) {
    if (parseTypedNumber(text).kind !== 'number') continue;
    const body = String(text).trim().replace(/^-/, '');
    const lastDot = body.lastIndexOf('.');
    const lastComma = body.lastIndexOf(',');
    if (lastDot >= 0 && lastComma >= 0) {
      if (lastDot > lastComma) dot = true;
      else comma = true;
    } else if (body.split('.').length > 2) comma = true;
    else if (body.split(',').length > 2) dot = true;
    else if (lastDot >= 0) dot = true;
    else if (lastComma >= 0) comma = true;
  }
  return dot === comma ? null : dot ? '.' : ',';
}

// rows: mảng dòng, mỗi ô { text, format?, numeric? }. fromExcel: có mẫu định dạng đi kèm.
// Trả về { rows: giá trị đã hiểu, guessed: ô phải suy, unread: ô số không đọc được nên giữ dạng chữ,
// rounded: ô lấy theo số đang hiện nên có thể đã làm tròn, narrow: ô đang hiện dấu thăng, không có số nào để lấy, decimal, hadEvidence }.
// plain: các ô của phần chữ thường cùng clipboard (cellsFromText), cùng vị trí với rows. Chỉ dùng cho ô Excel mẫu General.
export function interpretTable(rows, { fromExcel = false, plain = null } = {}) {
  const evidence = fromExcel ? null : decimalFrom(rows.flat().map((c) => c.text));
  // Không khớp thì bỏ hẳn chữ thường: mọi ô số từ Excel đều bị coi là có thể đã làm tròn, không đoán.
  if (!plainMatches(rows, plain)) plain = null;
  // Không có bằng chứng thì theo dấu của app (numberFormat.js). Trường hợp này luôn được báo.
  const decimal = evidence ?? DECIMAL;
  const guessed = [];
  const unread = [];
  const rounded = [];
  const narrow = [];
  // Chữ thuần, giống khi gõ tay (Q125): số lẻ viết bằng dấu khác dấu của app thì báo đã hiểu thế nào; chữ số lẫn lộn thì báo.
  const swapped = [];
  const unclear = [];
  const out = rows.map((row, r) => row.map((cell, c) => {
    const text = cell.text ?? '';
    if (text.trim() === '') return null;
    if (/^#+$/.test(text.trim())) {
      narrow.push({ r, c, text });
      return text;
    }
    if (fromExcel) {
      const read = fromExcelCell(cell, plain?.[r]?.[c]?.text);
      if (read === null) return text;
      if (read.unread) {
        unread.push({ r, c, text });
        return text;
      }
      if (read.guessed) guessed.push({ r, c, text });
      if (read.rounded) rounded.push({ r, c, text, value: read.value });
      return read.value;
    }
    const parsed = parseTypedNumber(text);
    if (parsed.kind === 'number') {
      if (usesOtherDecimal(text)) swapped.push({ r, c, text: text.trim(), value: parsed.value });
      return parsed.value;
    }
    if (parsed.kind === 'unclear') {
      unclear.push({ r, c, text: text.trim() });
      return text;
    }
    if (parsed.kind === 'ambiguous') {
      guessed.push({ r, c, text });
      const body = text.trim();
      return Number(body.includes(decimal) ? body.replace(decimal, '.') : body.replace(/[.,]/, ''));
    }
    return text;
  }));
  return { rows: out, guessed, unread, rounded, narrow, swapped, unclear, decimal, hadEvidence: evidence !== null, fromExcel };
}

// Đọc bảng trong HTML clipboard. Trả null nếu không có bảng. Chỉ chạy trên trình duyệt.
export function cellsFromHtml(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const formats = new Map();
  const alignedClasses = new Set();
  for (const m of html.matchAll(/\.(xl\d+)\s*\{([^}]*)\}/g)) {
    const format = m[2].match(/mso-number-format:\s*("(?:[^"\\]|\\.)*"|[^;]+)/);
    if (format) formats.set(m[1], format[1].replace(/^"|"$/g, '').trim());
    if (/text-align:/.test(m[2])) alignedClasses.add(m[1]);
  }
  const taken = new Set();
  const rows = [];
  [...table.rows].forEach((tr, r) => {
    const row = [];
    let c = 0;
    for (const td of tr.cells) {
      while (taken.has(`${r},${c}`)) {
        row[c] = { text: '' };
        c++;
      }
      const colspan = td.colSpan || 1;
      const rowspan = td.rowSpan || 1;
      const cls = [...td.classList].find((x) => formats.has(x));
      const aligned = [...td.classList].some((x) => alignedClasses.has(x));
      row[c] = { text: td.textContent.replace(/\s+/g, ' ').trim(), format: cls ? formats.get(cls) : null, numeric: (td.getAttribute('align') || '').toLowerCase() === 'right', aligned };
      if (colspan > 1 || rowspan > 1) rows.merged = true;
      for (let i = 0; i < rowspan; i++) for (let j = 0; j < colspan; j++) if (i || j) taken.add(`${r + i},${c + j}`);
      for (let j = 1; j < colspan; j++) row[c + j] = { text: '' };
      c += colspan;
    }
    rows.push(row);
  });
  return rows;
}

export function cellsFromText(text) {
  return text.replace(/\r\n/g, '\n').replace(/\n$/, '').split('\n').map((line) => line.split('\t').map((t) => ({ text: t })));
}
