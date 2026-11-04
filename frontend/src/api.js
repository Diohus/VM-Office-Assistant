const TIMEOUT_MS = 150000;
const NETWORK_MESSAGE = 'Chưa tải được file vì mất kết nối với máy chủ của app. File của bạn không bị đổi. Hãy kiểm tra kết nối rồi thử lại.';

export class ApiError extends Error {
  constructor(message, code) {
    super(message);
    this.code = code;
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
  throw new ApiError(body?.message || NETWORK_MESSAGE, body?.code || 'NETWORK');
}

export async function uploadDocument(file) {
  const res = await call(`/api/documents?filename=${encodeURIComponent(file.name)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/octet-stream' },
    body: file,
  });
  return res.json();
}

export async function fetchGrid(documentId) {
  return (await call(`/api/documents/${documentId}/grid`)).json();
}

export function exportUrl(documentId) {
  return `/api/documents/${documentId}/export`;
}
