import { DECIMAL } from './numberFormat.js';

const TIMEOUT_MS = 150000;
// Giới hạn ô sửa trong một lần lưu, phải bằng MAX_EDITS ở backend/app/main.py (Q124). Kiểm trước để không gửi đi rồi mới bị từ chối.
export const MAX_EDITS = 5000;
export const EDITS_TOO_MANY = `Lần lưu này có quá ${MAX_EDITS} ô sửa nên app chưa ghi. Bản đang lưu không bị đổi. Hãy lưu sau mỗi lần sửa ít hơn ${MAX_EDITS} ô.`;
// Máy chủ trả lỗi mà không có câu của app: không phải mất kết nối, không được nói là mất kết nối (Q124).
const SERVER_MESSAGE = (status) => `App chưa làm được việc này vì máy chủ của app trả lỗi (mã ${status}). File của bạn không bị đổi. Hãy thử lại; nếu vẫn lỗi, báo người quản lý app.`;
// 502, 503, 504: máy chủ trung gian không gọi được API, tức API đang tắt hoặc đang bật lại.
const UNREACHABLE = new Set([502, 503, 504]);
const NETWORK_MESSAGE = 'Chưa làm được vì mất kết nối với máy chủ của app. File của bạn không bị đổi. Hãy kiểm tra kết nối rồi thử lại.';

export class ApiError extends Error {
  constructor(message, code, detail) {
    super(message);
    this.code = code;
    this.detail = detail;
  }
}

async function call(path, options = {}) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
  let res;
  try {
    res = await fetch(path, { ...options, signal: ctrl.signal });
  } catch {
    throw new ApiError(NETWORK_MESSAGE, 'NETWORK');
  } finally {
    clearTimeout(timer);
  }
  if (res.ok) return res;
  // Máy chủ trung gian có thể trả trang HTML khi API tắt: không đưa nội dung đó cho người dùng.
  const body = await res.json().catch(() => null);
  if (typeof body?.message === 'string') throw new ApiError(body.message, body.code || 'SERVER', body.detail);
  if (UNREACHABLE.has(res.status)) throw new ApiError(NETWORK_MESSAGE, 'NETWORK');
  throw new ApiError(SERVER_MESSAGE(res.status), 'SERVER');
}

function postJson(path, body) {
  return call(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }).then((r) => r.json());
}

export async function uploadDocument(file) {
  const res = await call(`/api/documents?filename=${encodeURIComponent(file.name)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/octet-stream' },
    body: file,
  });
  return res.json();
}

export async function fetchGrid(documentId, version = null) {
  return (await call(`/api/documents/${documentId}/grid${version ? `?version=${version}` : ''}`)).json();
}

export async function fetchRecent() {
  return (await call('/api/documents')).json();
}

export async function fetchDocument(documentId) {
  return (await call(`/api/documents/${documentId}`)).json();
}

export async function fetchVersions(documentId) {
  return (await call(`/api/documents/${documentId}/versions`)).json();
}

export function saveVersion(documentId, body) {
  return postJson(`/api/documents/${documentId}/versions`, body);
}

export function restoreVersion(documentId, version, body) {
  return postJson(`/api/documents/${documentId}/versions/${version}/restore`, body);
}

// PDF in theo dấu thập phân của app (numberFormat.js), để số trong PDF cùng kiểu với lưới.
export async function exportPdf(documentId, body) {
  const res = await call(`/api/documents/${documentId}/pdf`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ...body, decimal: DECIMAL }) });
  const name = decodeURIComponent((res.headers.get('Content-Disposition') || '').split("UTF-8''")[1] || 'bieu_mau.pdf');
  return { blob: await res.blob(), name };
}

export function exportUrl(documentId, version = null) {
  return `/api/documents/${documentId}/export${version ? `?version=${version}` : ''}`;
}
