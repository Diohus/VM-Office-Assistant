// Dấu thập phân của cả app: một chỗ duy nhất (Q122, Owner câu 41: "Dùng dấu chấm đi 5.51").
//
// Muốn đổi sang dấu phẩy (5,51): đổi DECIMAL thành ','. Không cần sửa chỗ nào khác:
//   - ô trên lưới hiện theo NUMFMT_LOCALE,
//   - số gõ, câu thông báo, dán chữ thuần lấy dấu từ đây,
//   - PDF: frontend gửi DECIMAL kèm yêu cầu xuất, worker LibreOffice in theo dấu đó.
// Đổi xong: dựng lại app (sh tools/dev/up.sh) và chạy sh tests/integration/run.sh. Test đang viết cho dấu chấm.
// Dán từ Excel không theo cấu hình này: luôn đọc theo mẫu định dạng của ô trong clipboard (Q113).
export const DECIMAL = '.';

export const OTHER = DECIMAL === '.' ? ',' : '.';
// Tên vùng số của Univer: 'en' hiện 1,234.5; 'vi' hiện 1.234,5. Chữ giao diện vẫn tiếng Việt.
export const NUMFMT_LOCALE = DECIMAL === '.' ? 'en' : 'vi';

// Số viết theo dấu của app, đủ chữ số, không có dấu nghìn. Dùng trong câu thông báo.
export function formatNumber(value) {
  if (typeof value !== 'number') return String(value);
  const text = value.toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 20 });
  return DECIMAL === '.' ? text : text.replace('.', ',');
}

// Đổi chữ số gõ theo kiểu máy (dấu chấm) sang dấu của app, ví dụ gợi ý trong thông báo.
export function withDecimal(text) {
  return DECIMAL === '.' ? text : text.replace('.', ',');
}

// Chữ số gõ (đã hiểu được là số) có dấu thập phân khác dấu của app, ví dụ 5,5 khi app dùng dấu chấm.
// Dấu đứng sau cùng là dấu thập phân, nên 1.234,5 cũng tính. 1,234,567 là dấu nghìn, không tính.
export function usesOtherDecimal(text) {
  const body = String(text ?? '').trim().replace(/^-/, '');
  const at = Math.max(body.lastIndexOf('.'), body.lastIndexOf(','));
  if (at < 0 || body[at] !== OTHER) return false;
  return body.split(OTHER).length === 2;
}
