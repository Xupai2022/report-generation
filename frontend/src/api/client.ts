import {
  GenerateResponse,
  InputCatalogItem,
  PreviewResponse,
  RewriteResponse,
  TemplateCatalogItem,
} from "./types";

const API_PREFIX = "/api";

async function http<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${API_PREFIX}${path}`, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(text || res.statusText);
  }
  return res.json();
}

export const api = {
  listTemplates: () => http<TemplateCatalogItem[]>("/templates"),
  listInputs: () => http<InputCatalogItem[]>("/inputs"),
  generate: (input_id: string, template_id: string, use_mock = true) =>
    http<GenerateResponse>("/generate", {
      method: "POST",
      body: JSON.stringify({ input_id, template_id, use_mock }),
    }),
  rewrite: (job_id: string, slide_key: string, new_content: Record<string, unknown>) =>
    http<RewriteResponse>("/rewrite", {
      method: "POST",
      body: JSON.stringify({ job_id, slide_key, new_content }),
    }),
  getLogs: (limit = 50) => http<{ lines: string[] }>(`/logs?limit=${limit}`),
  preview: (job_id: string, regenerate_if_missing = true) =>
    http<PreviewResponse>(
      `/preview?job_id=${encodeURIComponent(job_id)}&regenerate_if_missing=${regenerate_if_missing}`
    ),
};
