export interface TemplateCatalogItem {
  template_id: string;
  name: string;
  version: string;
  slides_count?: number;
}

export interface InputCatalogItem {
  id: string;
  tenant_id: string;
  tenant_name: string;
  period: string;
}

export interface SlideSpecItem {
  slide_no: number;
  slide_key: string;
  data: Record<string, unknown>;
}

export interface SlideSpec {
  template_id: string;
  slides: SlideSpecItem[];
}

export interface GenerateResponse {
  job_id: string;
  report_path: string;
  warnings: string[];
  slidespec: SlideSpec;
  slidespec_path?: string;
}

export interface RewriteResponse {
  job_id: string;
  slide_key: string;
  report_path: string;
  warnings: string[];
  slidespec: SlideSpec;
}

export interface PreviewResponse {
  job_id: string;
  images: string[];
}
