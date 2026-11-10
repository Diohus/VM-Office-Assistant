// Lời thông báo cho người dùng. Không có từ kỹ thuật, tối đa 3 câu, luôn nói file có bị đổi không (acceptance_m1.md mục 7).
import { DECIMAL, formatNumber } from './numberFormat.js';

const SEP_NAME = { '.': 'dấu chấm', ',': 'dấu phẩy' };

const UNSHOWN_LABEL = {
  chart: 'biểu đồ',
  comment: 'nhận xét',
  threaded_comment: 'nhận xét',
  shape: 'hình vẽ',
  table_style: 'kiểu trình bày của bảng',
  data_bar: 'thanh dữ liệu',
  color_scale: 'thang màu',
  icon_set: 'bộ biểu tượng trong ô',
  conditional_format: 'quy tắc tô màu',
  pivot: 'bảng tổng hợp',
  sparkline: 'biểu đồ nhỏ trong ô',
  object: 'đối tượng nhúng',
  data_table: 'bảng dữ liệu',
  formula_unparsed: 'công thức app chưa đọc được',
  image_other: 'ảnh',
};

function joinVi(parts) {
  if (parts.length <= 1) return parts.join('');
  return `${parts.slice(0, -1).join(', ')} và ${parts[parts.length - 1]}`;
}

function firstCell(ref) {
  return ref ? ref.split(/[\s:]/)[0] : null;
}

function isComment(item) {
  return item.kind === 'comment' || item.kind === 'threaded_comment';
}

function unshownNotice(doc) {
  // Mục trên sheet ẩn không nêu ở đây: sheet ẩn vẫn ẩn, và vẫn còn nguyên trong file.
  const items = doc.unshown.filter((u) => !doc.hidden_sheets.includes(u.sheet));
  if (!items.length) return null;
  const counts = new Map();
  for (const it of items) {
    const label = UNSHOWN_LABEL[it.kind] || 'nội dung khác';
    counts.set(label, (counts.get(label) || 0) + 1);
  }
  const what = joinVi([...counts].map(([label, n]) => `${n} ${label}`));
  return {
    id: 'unshown',
    kind: 'info',
    text: `File này có ${what} app chưa hiện được. Chúng vẫn còn nguyên trong file và khi xuất file. Muốn xem thì xuất file rồi mở bằng Excel.`,
    places: items.map((it) => ({
      label: `${UNSHOWN_LABEL[it.kind] || 'Nội dung khác'}: sheet ${it.sheet}${it.cell ? `, ô ${it.cell}` : ''}${isComment(it) && it.detail ? `: "${it.detail}"` : ''}`,
      sheet: it.sheet,
      cell: firstCell(it.cell),
    })),
  };
}

function externalNotice(doc) {
  const visible = (b) => !doc.hidden_sheets.includes(b.sheet);
  const cells = doc.blocked_formulas.filter((b) => b.where === 'cell' && visible(b));
  const rules = doc.blocked_formulas.filter((b) => b.where !== 'cell' && visible(b));
  // Một liên kết có thể khai nhiều đường dẫn trong cùng một phần của file: chỉ đếm một lần.
  const links = [...new Set(doc.external.filter((e) => e.kind !== 'hyperlink').map((e) => e.part))];
  if (!cells.length && !rules.length && !links.length) return null;
  const parts = [];
  if (cells.length) parts.push(`${cells.length} ô lấy dữ liệu từ ngoài file`);
  if (rules.length) parts.push(`${rules.length} quy tắc tô màu hoặc danh sách chọn lấy dữ liệu từ ngoài file`);
  if (links.length) parts.push(`${links.length} liên kết tới nơi khác`);
  return {
    id: 'external',
    kind: 'warn',
    text: `File này có ${joinVi(parts)}. App không gọi ra ngoài, nên các ô đó chỉ hiện số đã lưu trong file, nếu có, và có thể không mới. File của anh/chị không bị đổi; hãy kiểm tra các ô được nêu.`,
    places: [...cells, ...rules].map((b) => ({ label: `Sheet ${b.sheet}, ô ${b.cell}`, sheet: b.sheet, cell: firstCell(b.cell) })),
  };
}

// Nguyên văn do PO/BA chốt, ui_spec.md mục 11 (Q100).
export const EXPORT_WARNING = [
  'Người nhận chỉ cần xem: bấm Tải PDF.',
  'Tải file Excel: khi mở, ô công thức, kể cả cột ngày, trống và biểu đồ vẽ từ các ô đó hiện 0 cho tới khi bấm Enable Editing. Bấm xong, Excel tự tính lại, số không mất. Nếu gửi file này cho người khác, hãy mở bằng Excel, lưu lại rồi mới gửi.',
];

export const EXPORT_UNCHANGED_LEGACY = ['Bản này chưa sửa trong app. File tải về là bản .xlsx app đã chuyển từ file đời cũ. File đời cũ của anh/chị vẫn còn nguyên.', 'Người nhận chỉ cần xem: bấm Tải PDF.'];
export const SAVE_FAILED_NOTE = ' Các sửa của anh/chị vẫn còn trên màn hình, chưa được lưu.';
export const EXPORT_UNCHANGED = ['Bản này chưa sửa trong app, nên file tải về giống hệt file đã tải lên.', 'Người nhận chỉ cần xem: bấm Tải PDF.'];

export function savedNotice(version, cells) {
  return { id: 'saved', kind: 'info', text: `Đã lưu bản ${version} trong app, ghi ${cells} ô. File trong folder của anh/chị chưa đổi. Muốn có file mới, bấm Xuất.`, places: [] };
}

export function ambiguousNumberNotice({ sheet, cell, text, unclear }) {
  if (unclear) {
    return {
      id: 'ambiguous',
      kind: 'warn',
      text: `Ô ${cell}: "${text}" không phải một số rõ ràng, nên app chưa ghi. File của anh/chị không bị đổi. Hãy gõ lại, dùng ${SEP_NAME[DECIMAL]} cho số lẻ, hoặc bấm Esc để bỏ.`,
      places: [{ label: `Sheet ${sheet}, ô ${cell}`, sheet, cell }],
    };
  }
  const compact = text.replace(/[.,]/g, '');
  const decimal = text.replace(/[.,]/, DECIMAL);
  return {
    id: 'ambiguous',
    kind: 'warn',
    text: `Ô ${cell}: "${text}" có thể là số nghìn ${compact} hoặc số lẻ ${decimal}, nên app chưa ghi. File của anh/chị không bị đổi. Hãy gõ ${compact} nếu là số nghìn, hoặc ${decimal}0 nếu là số lẻ.`,
    places: [{ label: `Sheet ${sheet}, ô ${cell}`, sheet, cell }],
  };
}

// Người dùng gõ số lẻ bằng dấu khác dấu của app, ví dụ 5,5 khi app dùng dấu chấm. App đã ghi đúng 5.5, không bao giờ thành 55 (Q122).
export function decimalSwappedNotice({ sheet, cell, text, value }) {
  return {
    id: 'decimal',
    kind: 'info',
    text: `Ô ${cell}: app đã hiểu "${text}" là ${formatNumber(value)}, vì app dùng ${SEP_NAME[DECIMAL]} cho số lẻ. Nếu không đúng, bấm Hoàn tác.`,
    places: [{ label: `Sheet ${sheet}, ô ${cell}`, sheet, cell }],
  };
}

// F2 trên ô có nhiều chữ số hơn ô đang sửa hiện được. Ghi thì mất các chữ số không hiện, nên hỏi trước (Q122).
export function hiddenDigitsNotice({ sheet, cell, full, shown, next }) {
  return {
    id: 'hidden-digits',
    kind: 'warn',
    text: `Ô ${cell} đang có số đầy đủ ${formatNumber(full)}, nhưng khi sửa chỉ hiện ${formatNumber(shown)}. Nếu ghi, ô sẽ thành ${formatNumber(next)} và mất các chữ số không hiện. App chưa ghi. Nhấn Enter lần nữa, không sửa gì thêm, để ghi số mới; hoặc bấm Esc để giữ số cũ.`,
    places: [{ label: `Sheet ${sheet}, ô ${cell}`, sheet, cell }],
  };
}

// Báo một lần cho cả vùng dán. Trả null nếu mọi ô đều hiểu chắc chắn.
export function pastedNotice({ sheet, tooBig, rows, cols, narrow = [], guessed = [], unread = [], rounded = [], swapped = [], unclear = [], decimal, hadEvidence, fromExcel }) {
  if (narrow.length) {
    return { id: 'pasted', kind: 'warn', text: `${narrow.length} ô đang hiện dấu thăng trong Excel vì cột hẹp, nên app chưa dán gì. File của anh/chị không bị đổi. Hãy nới rộng cột trong Excel, chép lại rồi dán.`, places: [] };
  }
  if (tooBig) {
    return { id: 'pasted', kind: 'warn', text: `Vùng dán có ${rows} dòng, ${cols} cột, vượt quá biểu mẫu nên app chưa dán. File của anh/chị không bị đổi. Hãy chọn ô đầu khác hoặc dán vùng nhỏ hơn.`, places: [] };
  }
  const places = [
    ...rounded.map((x) => ({ label: `Sheet ${sheet}, ô ${x.cell}: đã dán ${formatNumber(x.value)}, có thể đã làm tròn`, sheet, cell: x.cell })),
    ...guessed.filter((x) => !rounded.some((y) => y.cell === x.cell)).map((x) => ({ label: `Sheet ${sheet}, ô ${x.cell}: ${x.text}`, sheet, cell: x.cell })),
    ...unread.map((x) => ({ label: `Sheet ${sheet}, ô ${x.cell}: ${x.text}`, sheet, cell: x.cell })),
    ...swapped.map((x) => ({ label: `Sheet ${sheet}, ô ${x.cell}: "${x.text}" là ${formatNumber(x.value)}`, sheet, cell: x.cell })),
    ...unclear.map((x) => ({ label: `Sheet ${sheet}, ô ${x.cell}: "${x.text}" không phải số rõ ràng`, sheet, cell: x.cell })),
  ];
  const parts = [];
  if (guessed.length) {
    if (fromExcel) {
      parts.push(`${guessed.length} ô được canh riêng trong Excel nên app không biết chắc là số hay chữ. App đã hiểu là số. Hãy xem lại các ô này, nếu sai thì bấm Hoàn tác.`);
    } else {
      const how = `${SEP_NAME[decimal]} là số lẻ, ${SEP_NAME[decimal === '.' ? ',' : '.']} là số nghìn`;
      const why = hadEvidence ? 'theo các số khác trong vùng dán' : 'vì trong vùng dán không có gì để biết chắc, app tạm theo kiểu đang hiện trên bảng';
      parts.push(`${guessed.length} ô có thể hiểu hai cách. App đã hiểu ${how}, ${why}. Hãy xem lại các ô này, nếu sai thì bấm Hoàn tác.`);
    }
  }
  if (rounded.length) parts.push(`${rounded.length} ô có thể đã bị làm tròn, bấm Xem vị trí để thấy từng ô và số đã dán. Các ô này đã được dán theo số đang hiện trong Excel, có thể không phải số gốc. Excel không cho app đọc số gốc của các ô này. Muốn có số gốc, hãy đổi các ô đó sang định dạng General trong Excel, nới rộng cột, chép lại rồi dán. Số dài từ 10 chữ số trở lên thì hãy gõ tay.`);
  if (swapped.length) parts.push(`${swapped.length} ô viết số lẻ bằng ${SEP_NAME[DECIMAL === '.' ? ',' : '.']}. App dùng ${SEP_NAME[DECIMAL]} cho số lẻ nên đã hiểu, ví dụ "${swapped[0].text}" là ${formatNumber(swapped[0].value)}. Nếu không đúng, bấm Hoàn tác.`);
  if (unclear.length) parts.push(`${unclear.length} ô không phải một số rõ ràng, ví dụ "${unclear[0].text}", nên app giữ nguyên dạng chữ. Công thức sẽ không tính các ô này. Hãy gõ lại các ô này, dùng ${SEP_NAME[DECIMAL]} cho số lẻ.`);
  if (unread.length) parts.push(`${unread.length} ô app không đọc được thành số nên giữ nguyên dạng chữ, ví dụ ngày tháng hoặc số có đơn vị tiền. Công thức sẽ không tính các ô này. Hãy gõ lại nếu cần.`);
  return parts.length ? { id: 'pasted', kind: 'warn', text: parts.join(' '), places } : null;
}

export function lockedCellNotice({ sheet, cell }) {
  return {
    id: 'locked',
    kind: 'warn',
    text: `Ô ${cell} đang bị khóa bảo vệ trong file nên không sửa được. File của anh/chị không bị đổi. Nếu cần sửa, hãy hỏi người giữ biểu mẫu để bỏ khóa.`,
    places: [{ label: `Sheet ${sheet}, ô ${cell}`, sheet, cell }],
  };
}

export function refusedNotice(cells, message) {
  return {
    id: 'refused',
    kind: 'error',
    text: message,
    places: cells.map((c) => ({ label: `Sheet ${c.sheet}, ô ${c.cell}: ${c.reason}`, sheet: c.sheet, cell: c.cell })),
  };
}

export function pdfMismatchNotice(cells, message) {
  return {
    id: 'pdf-mismatch',
    kind: 'error',
    text: message,
    places: cells.map((c) => ({ label: `Sheet ${c.sheet}, ô ${c.cell}: trên bảng ${c.grid ?? 'trống'}, trong PDF ${c.pdf ?? 'trống'}`, sheet: c.sheet, cell: c.cell })),
  };
}

export function mismatchNotice(mismatches) {
  if (!mismatches.length) return null;
  return {
    id: 'mismatch',
    kind: 'warn',
    text: `Có ${mismatches.length} ô công thức app tính ra khác số đã lưu trong file. File của anh/chị không bị đổi. Hãy kiểm tra các ô này bằng Excel.`,
    places: mismatches.map((m) => ({ label: `Sheet ${m.sheet}, ô ${m.cell}`, sheet: m.sheet, cell: m.cell })),
  };
}

export function documentNotices(doc) {
  const out = [];
  for (const w of doc.warnings) out.push({ id: w.code, kind: 'info', text: w.message, places: [] });
  const ext = externalNotice(doc);
  if (ext) out.push(ext);
  const un = unshownNotice(doc);
  if (un) out.push(un);
  return out;
}
