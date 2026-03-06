import type {
  AdminJobDetailsResp,
  AdminJobsResp,
  AdminStatisticsResp,
  AdminVerifyResp,
  AiRewriteResp,
  CreateReportReq,
  CreateReportResp,
  InputItem,
  JobStatusResp,
  PreviewResp,
  RatingResp,
  RewriteSlideReq,
  RewriteSlidesResp,
  TemplateItem,
  TemplateSlideMeta,
} from '../types/api';
import { apiFetch } from './client';

export const MainApi = {
  listTemplates: () => apiFetch<TemplateItem[]>('/api/v1/templates'),
  listInputs: () => apiFetch<InputItem[]>('/api/v1/inputs'),
  getTemplateSlides: (templateId: string) => apiFetch<TemplateSlideMeta[]>(`/api/v1/templates/${encodeURIComponent(templateId)}/slides`),
  createReport: (body: CreateReportReq) => apiFetch<CreateReportResp>('/api/v1/reports', { method: 'POST', body: JSON.stringify(body) }),
  getJobStatus: (jobId: string) => apiFetch<JobStatusResp>(`/api/v1/jobs/${encodeURIComponent(jobId)}/status`),
  getPreview: (jobId: string, forceRegenerate = false) => apiFetch<PreviewResp>(`/api/v1/reports/${encodeURIComponent(jobId)}/preview?regenerate_if_missing=true&force_regenerate=${forceRegenerate ? 'true' : 'false'}`),
  rewriteSlides: (jobId: string, slides: RewriteSlideReq[], clientId?: string) =>
    apiFetch<RewriteSlidesResp>(`/api/v1/reports/${encodeURIComponent(jobId)}/slides`, {
      method: 'PATCH',
      body: JSON.stringify({ slides, client_id: clientId }),
    }),
  aiRewriteSlide: (jobId: string, slideKey: string, userPrompt: string, targetTokens: string[], clientId?: string) =>
    apiFetch<AiRewriteResp>(`/api/v1/reports/${encodeURIComponent(jobId)}/slides/ai-rewrite`, {
      method: 'POST',
      body: JSON.stringify({ slide_key: slideKey, user_prompt: userPrompt, target_tokens: targetTokens, client_id: clientId }),
    }),
  submitRating: (jobId: string, rating: 'liked' | 'disliked', comment?: string | null) =>
    apiFetch<RatingResp>(`/api/v1/ratings/jobs/${encodeURIComponent(jobId)}/rating`, {
      method: 'POST',
      body: JSON.stringify({ rating, comment: comment || null }),
    }),
};

export const AdminApi = {
  login: (username: string, password: string) =>
    apiFetch('/api/v1/admin/login', {
      method: 'POST',
      body: JSON.stringify({ username, password }),
    }),
  logout: () => apiFetch('/api/v1/admin/logout', { method: 'POST' }),
  verify: () => apiFetch<AdminVerifyResp>('/api/v1/admin/verify'),
  listJobs: (query: URLSearchParams) => apiFetch<AdminJobsResp>(`/api/v1/admin/jobs?${query.toString()}`),
  getJobDetail: (jobId: string) => apiFetch<AdminJobDetailsResp>(`/api/v1/admin/jobs/${jobId.replace(':', '_')}`),
  deleteJobs: (jobIds: string[]) =>
    apiFetch('/api/v1/admin/jobs', {
      method: 'DELETE',
      body: JSON.stringify({ job_ids: jobIds }),
    }),
  stats: () => apiFetch<AdminStatisticsResp>('/api/v1/admin/statistics'),
};
