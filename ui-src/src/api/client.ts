import type { ApiSuccess } from '../types/api';

export class HttpError extends Error {
  readonly status: number;
  readonly payload: unknown;

  constructor(status: number, message: string, payload: unknown) {
    super(message);
    this.status = status;
    this.payload = payload;
  }
}

export async function apiFetch<T>(path: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch(path, {
    credentials: 'include',
    headers: {
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
    ...init,
  });

  const text = await response.text();
  let payload: any = null;
  if (text) {
    try {
      payload = JSON.parse(text);
    } catch {
      payload = text;
    }
  }

  if (!response.ok) {
    const message = payload?.error?.message || payload?.detail || response.statusText || 'Request failed';
    throw new HttpError(response.status, message, payload);
  }

  if (payload && typeof payload === 'object' && 'data' in payload) {
    return (payload as ApiSuccess<T>).data;
  }

  return payload as T;
}

export async function apiNoJson<T>(path: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch(path, {
    credentials: 'include',
    ...init,
  });

  if (!response.ok) {
    throw new HttpError(response.status, response.statusText || 'Request failed', null);
  }

  return (await response.blob()) as unknown as T;
}
