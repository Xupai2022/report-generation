export function toSessionId() {
  const random = Array.from({ length: 8 }, () => Math.floor(Math.random() * 16).toString(16)).join('');
  const now = new Date();
  const ts = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, '0'),
    String(now.getDate()).padStart(2, '0'),
    String(now.getHours()).padStart(2, '0'),
    String(now.getMinutes()).padStart(2, '0'),
    String(now.getSeconds()).padStart(2, '0'),
    String(now.getMilliseconds() * 1000).padStart(6, '0'),
  ].join('');

  return `${random}_${ts}`;
}

export function toClientId() {
  return `client_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

export function formatDateTime(input?: string) {
  if (!input) return '--';
  try {
    return new Date(input).toLocaleString('zh-CN');
  } catch {
    return input;
  }
}

export function ensureArray<T>(value?: T[] | null): T[] {
  return Array.isArray(value) ? value : [];
}
