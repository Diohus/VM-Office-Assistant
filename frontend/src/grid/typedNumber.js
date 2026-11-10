// Hiểu số người dùng gõ (US-1.2, Q80). Dấu phẩy và dấu chấm đều có thể là dấu thập phân, nên ca nhập nhằng phải hỏi lại, không đoán.
//   { kind: 'number', value }     hiểu chắc chắn
//   { kind: 'ambiguous', text }   ví dụ 1.234, 1,234: có thể là một nghìn hoặc một phẩy
//   { kind: 'unclear', text }     chỉ gồm chữ số và dấu nhưng không ra số rõ ràng, ví dụ 1,2,3 hoặc 5,5.1: không để thư viện đoán
//   { kind: 'other' }             không phải số, để nguyên
export function parseTypedNumber(raw) {
  const text = String(raw ?? '').trim();
  if (!/^-?[\d.,]*\d[\d.,]*$/.test(text)) return { kind: 'other' };
  const body = text.replace(/^-/, '');
  const sign = text.startsWith('-') ? -1 : 1;
  const hasComma = body.includes(',');
  const hasDot = body.includes('.');
  if (!hasComma && !hasDot) return { kind: 'number', value: sign * Number(body) };
  if (hasComma && hasDot) {
    // Dấu đứng sau cùng là dấu thập phân: 1.234,5 và 1,234.5 đều là 1234,5.
    const decimal = body.lastIndexOf(',') > body.lastIndexOf('.') ? ',' : '.';
    const group = decimal === ',' ? '.' : ',';
    const [whole, fraction, extra] = body.split(decimal);
    if (extra !== undefined || !new RegExp(`^\\d{1,3}(\\${group}\\d{3})+$`).test(whole) || !/^\d+$/.test(fraction)) return { kind: 'unclear', text };
    return { kind: 'number', value: sign * Number(`${whole.split(group).join('')}.${fraction}`) };
  }
  const sep = hasComma ? ',' : '.';
  const parts = body.split(sep);
  if (parts.some((p) => !p)) return { kind: 'unclear', text };
  if (parts.length > 2) {
    // 1.234.567 chỉ có thể là phân nhóm nghìn.
    if (/^\d{1,3}$/.test(parts[0]) && parts.slice(1).every((p) => /^\d{3}$/.test(p))) return { kind: 'number', value: sign * Number(parts.join('')) };
    return { kind: 'unclear', text };
  }
  if (/^[1-9]\d{0,2}$/.test(parts[0]) && parts[1].length === 3) return { kind: 'ambiguous', text };
  return { kind: 'number', value: sign * Number(`${parts[0]}.${parts[1]}`) };
}
