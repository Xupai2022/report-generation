export type Lang = 'zh-CN' | 'en-US';

export interface ApiErrorPayload {
  code?: string;
  message?: string;
  details?: Record<string, unknown>;
}

export interface ApiSuccess<T> {
  data: T;
}

export interface InputItem {
  id: string;
  file?: string;
  description?: string;
  [key: string]: unknown;
}

export interface TemplateItem {
  template_id: string;
  name?: string;
  audience?: string;
  language?: string;
  slides_count?: number;
  [key: string]: unknown;
}

export interface TemplateSlideMeta {
  slide_no: number;
  slide_key: string;
  title?: string;
  ai_rewrite_tokens?: string[];
}

export interface CreateReportReq {
  input_id: string;
  template_id: string;
  use_mock: boolean;
  session_id: string;
  client_id: string;
  idempotency_key: string;
  focus_options: string[];
}

export interface CreateReportResp {
  job_id: string;
  session_id?: string;
  status: 'running' | 'completed' | string;
  message?: string;
  progress?: number;
  existing_job?: boolean;
  from_cache?: boolean;
  slidespec?: SlideSpec;
  preview_urls?: string[];
}

export interface JobStatusResp {
  job_id: string;
  status: 'pending' | 'running' | 'completed' | 'failed' | 'cancelled' | string;
  progress: number;
  message?: string;
  last_error?: string;
  error_code?: string;
  preview_urls?: string[];
  metadata?: Record<string, unknown>;
}

export interface SlideSpecSlide {
  slide_no: number;
  slide_key: string;
  title?: string;
  placeholders: Record<string, unknown>;
}

export interface SlideSpec {
  template_id: string;
  slides: SlideSpecSlide[];
}

export interface GenerateResult {
  job_id: string;
  report_path?: string;
  slidespec_path?: string;
  slidespec?: SlideSpec;
  preview_urls?: string[];
  warnings?: string[];
  [key: string]: unknown;
}

export interface PreviewResp {
  preview_urls?: string[];
  images?: string[];
  timings?: Record<string, number>;
}

export interface RewriteSlideReq {
  slide_key: string;
  new_content: Record<string, unknown>;
}

export interface RewriteSlidesResp {
  job_id: string;
  updated_slides: string[];
  updated_count: number;
  warnings?: string[];
  slidespec?: SlideSpec;
}

export interface AiRewriteResp extends RewriteSlidesResp {
  slide_key?: string;
  updated_tokens?: string[];
}

export interface RatingResp {
  job_id: string;
  rating: 'liked' | 'disliked' | null;
  message?: string;
}

export interface AdminVerifyResp {
  authenticated: boolean;
  username?: string;
}

export interface AdminJob {
  job_id: string;
  input_id: string;
  template_id: string;
  status: string;
  created_at: string;
  progress?: number;
  preview_thumbnail?: string;
  rating?: {
    rating: 'liked' | 'disliked' | null;
    comment?: string;
    rated_at?: string;
  } | null;
  metadata?: Record<string, unknown>;
}

export interface AdminJobsResp {
  jobs: AdminJob[];
  total: number;
  page: number;
  limit: number;
}

export interface AdminJobDetailsResp {
  job: AdminJob;
  preview_urls: string[];
}

export interface AdminStatisticsResp {
  total_jobs: number;
  success_rate: number;
  avg_generation_time_ms?: number;
  by_rating?: Record<string, number>;
}

export interface I18nMap {
  [key: string]: string;
}

export interface WsProgressMessage {
  type: 'progress';
  progress: number;
  message?: string;
  details?: Record<string, unknown>;
}

export interface WsCompleteMessage {
  type: 'completed';
  result: GenerateResult;
}

export interface WsFailedMessage {
  type: 'failed';
  result: {
    error?: string;
    error_code?: string;
  };
}

export type WsMessage = WsProgressMessage | WsCompleteMessage | WsFailedMessage | { type: string; [key: string]: unknown };
