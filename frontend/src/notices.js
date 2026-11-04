// Lời thông báo cho người dùng. Không có từ kỹ thuật, tối đa 3 câu, luôn nói file có bị đổi không (acceptance_m1.md mục 7).

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
    text: `File này có ${joinVi(parts)}. App không gọi ra ngoài, nên các số đó là số đã lưu trong file và có thể không mới. File của bạn không bị đổi; hãy kiểm tra các ô được nêu.`,
    places: [...cells, ...rules].map((b) => ({ label: `Sheet ${b.sheet}, ô ${b.cell}`, sheet: b.sheet, cell: firstCell(b.cell) })),
  };
}

export function mismatchNotice(mismatches) {
  if (!mismatches.length) return null;
  return {
    id: 'mismatch',
    kind: 'warn',
    text: `Có ${mismatches.length} ô công thức app tính ra khác số đã lưu trong file. File của bạn không bị đổi. Hãy kiểm tra các ô này bằng Excel.`,
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
