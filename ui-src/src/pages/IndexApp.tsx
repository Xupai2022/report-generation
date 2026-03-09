import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  AlertTriangle,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Download,
  Edit3,
  FileSpreadsheet,
  FileText,
  Funnel,
  Globe,
  LayoutTemplate,
  LayoutGrid,
  Info,
  List,
  Maximize,
  Play,
  RefreshCw,
  Settings2,
  Sparkles,
  UploadCloud,
  X,
} from 'lucide-react';

import { MainApi } from '../api/main';
import { HttpError } from '../api/client';
import type {
  CreateReportResp,
  GenerateResult,
  InputItem,
  JobStatusResp,
  SlideSpec,
  SlideSpecSlide,
  TemplateItem,
  TemplateSlideMeta,
  WsMessage,
} from '../types/api';
import { useI18n } from '../shared/useI18n';
import { ensureArray, toClientId, toSessionId } from '../shared/utils';

const FOCUS_OPTIONS = [
  { value: 'business_protection', labelKey: 'focusOptionBusinessProtection', descKey: 'focusOptionBusinessProtectionDesc' },
  { value: 'vulnerability', labelKey: 'focusOptionVulnerability', descKey: 'focusOptionVulnerabilityDesc' },
  { value: 'alert', labelKey: 'focusOptionAlert', descKey: 'focusOptionAlertDesc' },
] as const;

type FocusValue = (typeof FOCUS_OPTIONS)[number]['value'];
const FOCUS_VALUE_TO_PREFERENCE_TEXT: Record<FocusValue, string> = {
  vulnerability: '漏洞优先',
  alert: '告警优先',
  business_protection: '业务保护',
};
type ToastType = 'success' | 'error' | 'warning' | 'info';

interface ToastState {
  id: number;
  message: string;
  type: ToastType;
}

interface AiRewriteModeBySlide {
  [slideKey: string]: 'all' | 'selected';
}

interface AiRewriteTokenSelection {
  [slideKey: string]: string[];
}

interface RecentAiPromptsBySlide {
  [slideKey: string]: string[];
}

interface StringFieldMap {
  [field: string]: string;
}

interface ModifiedSlideFields {
  [slideKey: string]: StringFieldMap;
}

function toTemplateName(tpl: TemplateItem): string {
  if (typeof tpl.name === 'string' && tpl.name.trim()) {
    return tpl.name;
  }
  return tpl.template_id;
}

function getTemplatePreviewImageUrl(templateId: string, frameIndex: number) {
  const folderByTemplate: Record<string, string> = {
    mss_classic_ops: '/static/previews/template_classic_preview',
  };
  const folder = folderByTemplate[templateId];
  if (!folder) return '';
  const safeIndex = Number.isFinite(frameIndex) ? Math.max(0, frameIndex) : 0;
  return `${folder}/slide${safeIndex + 1}.png`;
}

function createWsUrl(clientId: string) {
  const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
  return `${protocol}//${window.location.host}/ws/${clientId}`;
}

function sameKeys(a: unknown, b: unknown): boolean {
  if (Array.isArray(a) || Array.isArray(b)) {
    return Array.isArray(a) && Array.isArray(b);
  }
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object') {
    return true;
  }
  const keysA = Object.keys(a as Record<string, unknown>).sort();
  const keysB = Object.keys(b as Record<string, unknown>).sort();
  if (keysA.length !== keysB.length) {
    return false;
  }
  return keysA.every((key, idx) => key === keysB[idx]);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

function hasSameStructuredKeys(templateValue: unknown, editedValue: unknown): boolean {
  if (Array.isArray(templateValue)) {
    if (!Array.isArray(editedValue)) return false;
    if (templateValue.length === 0) return true;
    return editedValue.every((item, idx) => hasSameStructuredKeys(templateValue[Math.min(idx, templateValue.length - 1)], item));
  }
  if (isPlainObject(templateValue)) {
    if (!isPlainObject(editedValue)) return false;
    if (!sameKeys(templateValue, editedValue)) return false;
    return Object.keys(templateValue).every((key) => hasSameStructuredKeys(templateValue[key], editedValue[key]));
  }
  return true;
}

function restoreStructuredKeysByTemplate(templateValue: unknown, editedValue: unknown): unknown {
  if (Array.isArray(templateValue)) {
    if (!Array.isArray(editedValue)) return templateValue;
    if (templateValue.length === 0) return editedValue;
    return editedValue.map((item, idx) => restoreStructuredKeysByTemplate(templateValue[Math.min(idx, templateValue.length - 1)], item));
  }
  if (isPlainObject(templateValue)) {
    if (!isPlainObject(editedValue)) return templateValue;
    const restored: Record<string, unknown> = {};
    const editedEntries = Object.entries(editedValue);
    Object.entries(templateValue).forEach(([templateKey, templateChild], idx) => {
      const editedChild = idx < editedEntries.length ? editedEntries[idx][1] : templateChild;
      restored[templateKey] = restoreStructuredKeysByTemplate(templateChild, editedChild);
    });
    return restored;
  }
  return typeof editedValue === 'undefined' ? templateValue : editedValue;
}

function parseEditedValue(raw: string, original: unknown, field: string): unknown {
  if (typeof original === 'number') {
    const value = Number(raw);
    if (Number.isNaN(value)) {
      throw new Error(`字段 ${field} 必须是数字`);
    }
    return value;
  }
  if (typeof original === 'boolean') {
    const normalized = raw.trim().toLowerCase();
    if (normalized === 'true') return true;
    if (normalized === 'false') return false;
    throw new Error(`字段 ${field} 必须是 true 或 false`);
  }
  if (original && typeof original === 'object') {
    try {
      const parsed = JSON.parse(raw);
      return parsed;
    } catch (error) {
      if (error instanceof Error) throw error;
      throw new Error(`字段 ${field} JSON 无效`);
    }
  }
  return raw;
}

function stringifyValue(value: unknown): string {
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  return JSON.stringify(value, null, 2);
}

function firstSlide(slidespec: SlideSpec | null): string | null {
  if (!slidespec || !Array.isArray(slidespec.slides) || slidespec.slides.length === 0) {
    return null;
  }
  return slidespec.slides[0]?.slide_key || null;
}

function isReloadNavigation(): boolean {
  try {
    const entries = window.performance.getEntriesByType('navigation') as PerformanceNavigationTiming[];
    if (entries.length > 0) return entries[0].type === 'reload';
  } catch {
    // ignore
  }
  const legacy = (window.performance as unknown as { navigation?: { type?: number } }).navigation;
  return legacy?.type === 1;
}

export function IndexApp() {
  const { lang, toggleLang, t } = useI18n();

  const [templates, setTemplates] = useState<TemplateItem[]>([]);
  const [inputs, setInputs] = useState<InputItem[]>([]);
  const [templateSlidesById, setTemplateSlidesById] = useState<Record<string, TemplateSlideMeta[]>>({});
  const [templateFrameIndexById, setTemplateFrameIndexById] = useState<Record<string, number>>({});
  const [templatePreviewStamp, setTemplatePreviewStamp] = useState<number>(() => Date.now());

  const [selectedTemplate, setSelectedTemplate] = useState('');
  const [selectedInput, setSelectedInput] = useState('');
  const [useMock, setUseMock] = useState(true);
  const [selectedFocus, setSelectedFocus] = useState<FocusValue[]>(['business_protection']);

  const [inWorkspace, setInWorkspace] = useState(false);
  const [loading, setLoading] = useState(false);
  const [wsReady, setWsReady] = useState(false);
  const [generationInProgress, setGenerationInProgress] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressText, setProgressText] = useState('');

  const [jobId, setJobId] = useState<string | null>(null);
  const [sessionId, setSessionId] = useState<string>('');
  const [clientId, setClientId] = useState<string>('');

  const [slidespec, setSlidespec] = useState<SlideSpec | null>(null);
  const [previews, setPreviews] = useState<string[]>([]);
  const [activeSlideKey, setActiveSlideKey] = useState<string | null>(null);
  const [previewLayout, setPreviewLayout] = useState<'list' | 'two-up'>('list');
  const [modifiedSlides, setModifiedSlides] = useState<ModifiedSlideFields>({});
  const [restoredFieldMarks, setRestoredFieldMarks] = useState<Record<string, boolean>>({});

  const [exportOpen, setExportOpen] = useState(false);
  const [exportFormat, setExportFormat] = useState<'ppt' | 'pdf'>('ppt');

  const [aiModalOpen, setAiModalOpen] = useState(false);
  const [aiPrompt, setAiPrompt] = useState('');
  const [aiModeBySlide, setAiModeBySlide] = useState<AiRewriteModeBySlide>({});
  const [aiSelectedTokensBySlide, setAiSelectedTokensBySlide] = useState<AiRewriteTokenSelection>({});
  const [recentAiPromptsBySlide, setRecentAiPromptsBySlide] = useState<RecentAiPromptsBySlide>({});

  const [toasts, setToasts] = useState<ToastState[]>([]);
  const [ratingOpen, setRatingOpen] = useState(false);
  const [ratingChoice, setRatingChoice] = useState<'liked' | 'disliked' | null>(null);
  const [ratingComment, setRatingComment] = useState('');
  const [ratedJobs, setRatedJobs] = useState<Set<string>>(new Set());
  const [previewLoadErrorKeys, setPreviewLoadErrorKeys] = useState<Record<string, true>>({});

  const [presentOpen, setPresentOpen] = useState(false);
  const [presentIndex, setPresentIndex] = useState(0);
  const presenterRef = useRef<HTMLDivElement | null>(null);
  const editorPaneRef = useRef<HTMLElement | null>(null);
  const exportPopoverRef = useRef<HTMLDivElement | null>(null);
  const aiPopoverRef = useRef<HTMLDivElement | null>(null);

  const wsRef = useRef<WebSocket | null>(null);
  const wsRetryRef = useRef<number | null>(null);
  const pollRef = useRef<number | null>(null);
  const progressValueRef = useRef<number>(0);
  const actionRef = useRef<'generate' | 'rewrite' | 'ai-rewrite' | null>(null);
  const completionHandledJobRef = useRef<string | null>(null);
  const previewWheelAtRef = useRef<number>(0);
  const previewWheelAccumRef = useRef<number>(0);
  const previewScrollRafRef = useRef<number | null>(null);

  const activeSlide = useMemo<SlideSpecSlide | null>(() => {
    if (!slidespec || !activeSlideKey) return null;
    return slidespec.slides.find((s) => s.slide_key === activeSlideKey) || null;
  }, [slidespec, activeSlideKey]);

  const activeSlideIndex = useMemo(() => {
    if (!slidespec || !activeSlideKey) return -1;
    return slidespec.slides.findIndex((s) => s.slide_key === activeSlideKey);
  }, [slidespec, activeSlideKey]);

  const activeEditorFields = useMemo(() => {
    if (!activeSlide) return {} as StringFieldMap;
    const initial: StringFieldMap = {};
    Object.entries(activeSlide.placeholders || {}).forEach(([k, v]) => {
      initial[k] = stringifyValue(v);
    });
    const edited = modifiedSlides[activeSlide.slide_key];
    return edited ? { ...initial, ...edited } : initial;
  }, [activeSlide, modifiedSlides]);

  const activeFieldKeys = useMemo(() => {
    if (!activeSlide) return [] as string[];
    return Object.keys(activeSlide.placeholders || {});
  }, [activeSlide]);

  const activeAiTokens = useMemo(() => {
    if (!selectedTemplate || !activeSlideKey) return [] as string[];
    const slides = templateSlidesById[selectedTemplate] || [];
    const meta = slides.find((s) => s.slide_key === activeSlideKey);
    return ensureArray(meta?.ai_rewrite_tokens);
  }, [selectedTemplate, activeSlideKey, templateSlidesById]);

  const selectedAiMode = useMemo(() => {
    if (!activeSlideKey) return 'all' as const;
    return aiModeBySlide[activeSlideKey] || 'all';
  }, [activeSlideKey, aiModeBySlide]);

  const selectedAiTokens = useMemo(() => {
    if (!activeSlideKey) return [];
    return ensureArray(aiSelectedTokensBySlide[activeSlideKey]);
  }, [activeSlideKey, aiSelectedTokensBySlide]);

  const aiRewriteTargetCount = useMemo(() => {
    if (selectedAiMode === 'selected') return selectedAiTokens.length;
    return activeAiTokens.length > 0 ? activeAiTokens.length : activeFieldKeys.length;
  }, [selectedAiMode, selectedAiTokens, activeAiTokens, activeFieldKeys]);

  const recentAiPrompts = useMemo(() => {
    if (!activeSlideKey) return [] as string[];
    return ensureArray(recentAiPromptsBySlide[activeSlideKey]);
  }, [activeSlideKey, recentAiPromptsBySlide]);

  const showToast = (type: ToastType, message: string) => {
    const normalized = localizeServerMessage(message || '');
    const id = Date.now() + Math.floor(Math.random() * 9999);
    setToasts((prev) => [...prev, { id, type, message: normalized }]);
    window.setTimeout(() => {
      setToasts((prev) => prev.filter((item) => item.id !== id));
    }, 4200);
  };

  const lt = (key: string, zh: string, en: string) => {
    const fallback = lang === 'zh-CN' ? zh : en;
    const value = t(key, fallback);
    const hasZh = /[\u4e00-\u9fa5]/.test(value);
    const hasEn = /[A-Za-z]/.test(value);
    if (lang === 'zh-CN' && hasEn && !hasZh) return zh;
    if (lang === 'en-US' && hasZh) return en;
    return value;
  };

  const localizeServerMessage = (raw: string): string => {
    const msg = (raw || '').trim();
    if (!msg) return msg;
    const lower = msg.toLowerCase();
    if (lower.includes('already generating') || lower.includes('reusing the running task')) {
      return lt('msgReusingRunningTask', '检测到你已有正在生成的任务，已继续该任务进度。', 'An existing running task was found in this browser. Continuing that task.');
    }
    if (lower.includes('generating report') && lower.includes('please wait')) {
      return lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...');
    }
    if (lower.includes('generation failed') || lower.includes('report generation failed')) {
      return lt('errorGenerationFailed', '报告生成失败', 'Report generation failed');
    }
    if (lower.includes('preview') && lower.includes('fail')) {
      return lt('errorPreviewFailed', '预览生成失败', 'Preview generation failed');
    }
    if (lower.includes('has no ai-generated placeholders to rewrite')) {
      return lt(
        'msgNoAiGeneratedPlaceholdersToRewrite',
        '当前页没有可用于 AI 重写的占位字段。',
        'This slide has no AI-generated placeholders to rewrite.',
      );
    }
    if (lower.includes('load') && lower.includes('fail')) {
      return lt('errorLoadFailed', '数据加载失败', 'Data loading failed');
    }
    return msg;
  };

  const getToastTitle = (type: ToastType): string => {
    if (type === 'success') return t('titleUpdateCompleted', lang === 'zh-CN' ? '完成' : 'Done');
    if (type === 'error') return t('titleUpdateFailed', lang === 'zh-CN' ? '错误' : 'Error');
    if (type === 'warning') return t('toastTitleHint', lang === 'zh-CN' ? '提示' : 'Notice');
    return t('toastTitleHint', lang === 'zh-CN' ? '提示' : 'Notice');
  };

  const localizeProgressMessage = (raw?: string): string => {
    const text = (raw || '').trim();
    if (!text) return t('progressInit', lang === 'zh-CN' ? '初始化' : 'Initializing');
    const lower = text.toLowerCase();
    if (lower.includes('already generating') || lower.includes('reusing the running task')) {
      return t('progressGenerateContent', lang === 'zh-CN' ? 'AI 生成内容' : 'AI generating content');
    }
    if (lower.includes('init')) return t('progressInit', lang === 'zh-CN' ? '初始化' : 'Initializing');
    if (lower.includes('load') && lower.includes('template')) return t('progressLoadTemplate', lang === 'zh-CN' ? '加载模板' : 'Loading template');
    if ((lower.includes('ai') && lower.includes('generat')) || lower.includes('generate content')) {
      return t('progressGenerateContent', lang === 'zh-CN' ? 'AI 生成内容' : 'AI generating content');
    }
    if (lower.includes('render') && lower.includes('preview')) {
      return t('progressRenderPreview', lang === 'zh-CN' ? '渲染预览图' : 'Rendering preview images');
    }
    if (lower.includes('render') && lower.includes('ppt')) {
      return t('progressRenderPpt', lang === 'zh-CN' ? '渲染 PPT' : 'Rendering PPT');
    }
    if (lower.includes('complete') || lower.includes('done') || lower.includes('finished')) {
      return t('progressComplete', lang === 'zh-CN' ? '完成' : 'Completed');
    }
    return t('statusProcessing', lang === 'zh-CN' ? '生成中' : 'Processing');
  };

  const applyIncomingProgress = (nextProgress: number | undefined, message?: string) => {
    const numeric = typeof nextProgress === 'number' ? nextProgress : 0;
    if (numeric < progressValueRef.current) return;
    progressValueRef.current = numeric;
    setProgress(numeric);
    if (message) {
      const localized = localizeProgressMessage(message);
      const completedLabel = t('progressComplete', lang === 'zh-CN' ? '完成' : 'Completed');
      if (numeric < 100 && localized === completedLabel) {
        if (actionRef.current === 'rewrite' || actionRef.current === 'ai-rewrite') {
          setProgressText(t('progressRenderPreview', lang === 'zh-CN' ? '渲染预览图' : 'Rendering preview images'));
        } else {
          setProgressText(t('statusProcessing', lang === 'zh-CN' ? '生成中' : 'Processing'));
        }
      } else {
        setProgressText(localized);
      }
    }
  };

  const loadTemplateSlides = async (templateId: string) => {
    if (!templateId || templateSlidesById[templateId]) return;
    try {
      const slides = await MainApi.getTemplateSlides(templateId);
      setTemplateSlidesById((prev) => ({ ...prev, [templateId]: slides }));
    } catch (error) {
      console.error('Failed to load template slide metadata', error);
    }
  };

  const connectWebSocket = (nextClientId: string) => {
    if (wsRef.current) {
      wsRef.current.close();
      wsRef.current = null;
    }
    const socket = new WebSocket(createWsUrl(nextClientId));
    wsRef.current = socket;
    socket.onopen = () => setWsReady(true);
    socket.onclose = () => {
      setWsReady(false);
      if (wsRetryRef.current) window.clearTimeout(wsRetryRef.current);
      wsRetryRef.current = window.setTimeout(() => {
        connectWebSocket(nextClientId);
      }, 3000);
    };
    socket.onerror = () => setWsReady(false);
    socket.onmessage = (event) => {
      let payload: WsMessage;
      try {
        payload = JSON.parse(event.data) as WsMessage;
      } catch {
        return;
      }
      if (payload.type === 'progress') {
        applyIncomingProgress(payload.progress, payload.message);
        return;
      }
      if (payload.type === 'completed') {
        void handleGenerationComplete(payload.result);
        return;
      }
      if (payload.type === 'failed') {
        handleGenerationFailed(payload.result?.error || t('errorGenerationFailed'));
      }
    };
  };

  const loadBootstrapData = async () => {
    try {
      const [templatesData, inputsData] = await Promise.all([MainApi.listTemplates(), MainApi.listInputs()]);
      setTemplates(templatesData);
      setInputs(inputsData);
      const firstTemplate = templatesData[0]?.template_id || '';
      const firstInput = inputsData[0]?.id || '';
      setSelectedTemplate(firstTemplate);
      setSelectedInput(firstInput);
      if (firstTemplate) void loadTemplateSlides(firstTemplate);
    } catch (error) {
      const message = error instanceof Error ? error.message : t('errorLoadFailed');
      showToast('error', message);
    }
  };

  useEffect(() => {
    if (isReloadNavigation()) {
      // Browser refresh should always return user to the configuration screen.
      setInWorkspace(false);
      setLoading(false);
      setGenerationInProgress(false);
      setProgress(0);
      setProgressText('');
      setJobId(null);
      setSlidespec(null);
      setPreviews([]);
      setActiveSlideKey(null);
      setModifiedSlides({});
      setExportOpen(false);
      setAiModalOpen(false);
    }
    const cid = toClientId();
    setClientId(cid);
    connectWebSocket(cid);
    void loadBootstrapData();
    return () => {
      if (wsRetryRef.current) window.clearTimeout(wsRetryRef.current);
      if (pollRef.current) window.clearInterval(pollRef.current);
      if (wsRef.current) wsRef.current.close();
    };
  }, []);

  useEffect(() => {
    if (!selectedTemplate) return;
    void loadTemplateSlides(selectedTemplate);
  }, [selectedTemplate]);

  useEffect(() => {
    setTemplatePreviewStamp(Date.now());
  }, [templates.length]);

  const stopPolling = () => {
    if (pollRef.current) {
      window.clearInterval(pollRef.current);
      pollRef.current = null;
    }
  };

  const syncPreviewUrls = async (currentJobId: string, forceRegenerate = false) => {
    const previewData = await MainApi.getPreview(currentJobId, forceRegenerate);
    const next = ensureArray(previewData.preview_urls).length > 0 ? ensureArray(previewData.preview_urls) : ensureArray(previewData.images);
    setPreviews(next);
  };

  const tryLoadSlidespecFromPath = async (slidespecPath?: string) => {
    if (!slidespecPath) return null;
    try {
      const response = await fetch(slidespecPath, { credentials: 'include' });
      if (!response.ok) return null;
      return (await response.json()) as SlideSpec;
    } catch {
      return null;
    }
  };

  const handleGenerationComplete = async (result: GenerateResult) => {
    const nextJobId = result.job_id || jobId;
    if (!nextJobId) return;
    if (completionHandledJobRef.current === nextJobId) return;
    completionHandledJobRef.current = nextJobId;
    const shouldToastSuccess = actionRef.current === 'generate';
    setGenerationInProgress(false);
    stopPolling();
    progressValueRef.current = 100;
    setProgress(100);
    setProgressText(t('progressComplete', lang === 'zh-CN' ? '完成' : 'Completed'));
    setJobId(nextJobId);
    const templateId = nextJobId.split(':')[1] || selectedTemplate;
    if (templateId) {
      setSelectedTemplate(templateId);
      void loadTemplateSlides(templateId);
    }
    let nextSlidespec = result.slidespec || null;
    if (!nextSlidespec) {
      nextSlidespec = await tryLoadSlidespecFromPath(typeof result.slidespec_path === 'string' ? result.slidespec_path : undefined);
    }
    if (nextSlidespec) {
      setSlidespec(nextSlidespec);
      setActiveSlideKey((prev) => prev || firstSlide(nextSlidespec));
      setModifiedSlides({});
    }
    const resultPreviews = ensureArray(result.preview_urls);
    const expectedPreviewCount = nextSlidespec?.slides?.length || 0;
    if (resultPreviews.length > 0 && (expectedPreviewCount === 0 || resultPreviews.length >= expectedPreviewCount)) {
      setPreviews(resultPreviews);
    } else {
      try {
        await syncPreviewUrls(nextJobId);
      } catch (error) {
        console.error('Failed to load previews after completion', error);
      }
    }
    if (shouldToastSuccess) {
      showToast('success', lt('msgSuccess', '报告生成成功！', 'Report generated successfully!'));
    }
    actionRef.current = null;
  };

  const handleGenerationFailed = (message: string) => {
    setGenerationInProgress(false);
    stopPolling();
    progressValueRef.current = 0;
    setProgress(0);
    setProgressText('');
    actionRef.current = null;
    showToast('error', message || lt('errorGenerationFailed', '报告生成失败', 'Report generation failed'));
  };

  const checkCurrentJobStatus = async () => {
    if (!jobId || !generationInProgress) return;
    if (actionRef.current && actionRef.current !== 'generate') return;
    try {
      const statusData = (await MainApi.getJobStatus(jobId)) as JobStatusResp & { slidespec_path?: string };
      if (statusData.status === 'running' || statusData.status === 'pending') {
        applyIncomingProgress(statusData.progress, statusData.message);
        return;
      }
      if (statusData.status === 'failed' || statusData.status === 'cancelled') {
        handleGenerationFailed(statusData.last_error || t('msgRegenerateAfterError'));
        return;
      }
      if (statusData.status === 'completed') {
        const fallbackSlidespec = await tryLoadSlidespecFromPath(statusData.slidespec_path);
        await handleGenerationComplete({
          job_id: statusData.job_id,
          preview_urls: statusData.preview_urls,
          slidespec: fallbackSlidespec || undefined,
        });
      }
    } catch (error) {
      console.warn('poll status error', error);
    }
  };

  useEffect(() => {
    if (!generationInProgress || !jobId) {
      stopPolling();
      return;
    }
    stopPolling();
    pollRef.current = window.setInterval(() => {
      void checkCurrentJobStatus();
    }, 2000);
    return () => {
      stopPolling();
    };
  }, [generationInProgress, jobId]);

  const normalizeFocus = (values: FocusValue[]) => {
    const set = new Set<FocusValue>();
    values.forEach((item) => {
      if (item === 'vulnerability' || item === 'alert' || item === 'business_protection') set.add(item);
    });
    return Array.from(set);
  };

  const toPreferenceTitles = (values: FocusValue[]) => {
    return values.map((value) => FOCUS_VALUE_TO_PREFERENCE_TEXT[value]);
  };

  const beginGenerate = async () => {
    if (!selectedInput) return showToast('warning', lt('msgSelectInput', '请选择输入数据', 'Please select input data'));
    if (!selectedTemplate) return showToast('warning', lt('msgSelectTemplate', '请选择报告模板', 'Please select report template'));
    const focus = normalizeFocus(selectedFocus);
    if (focus.length === 0) return showToast('warning', lt('msgSelectAtLeastOneFocus', '请至少选择一个关注焦点', 'Please select at least one focus area'));
    const focusTitles = toPreferenceTitles(focus);
    if (generationInProgress) return showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
    if (!inWorkspace) setInWorkspace(true);
    const sid = sessionId || toSessionId();
    setSessionId(sid);
    setLoading(true);
    setGenerationInProgress(true);
    actionRef.current = 'generate';
    completionHandledJobRef.current = null;
    progressValueRef.current = 0;
    setProgress(0);
    setProgressText(t('progressInit', lang === 'zh-CN' ? '初始化' : 'Initializing'));
    try {
      const firstSlideKey = firstSlide(slidespec);
      if (firstSlideKey) {
        setActiveSlideKey(firstSlideKey);
      }
      const payload = {
        input_id: selectedInput,
        template_id: selectedTemplate,
        use_mock: useMock,
        session_id: sid,
        client_id: clientId,
        idempotency_key: sid,
        focus_options: focusTitles,
      };
      const result = await MainApi.createReport(payload);
      setJobId(result.job_id || null);
      if (result.status === 'completed' && result.from_cache) {
        await handleGenerationComplete(result as CreateReportResp & GenerateResult);
        return;
      }
      applyIncomingProgress(result.progress, result.message);
      setPreviews([]);
      setSlidespec(null);
      setActiveSlideKey(null);
      if (result.existing_job) {
        showToast('info', lt('msgReusingRunningTask', '检测到你已有正在生成的任务，已继续该任务进度。', 'An existing running task was found in this browser. Continuing that task.'));
      } else {
        showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
      }
      if (!wsReady) void checkCurrentJobStatus();
    } catch (error) {
      const message = error instanceof HttpError ? error.message : (error as Error).message;
      handleGenerationFailed(message || t('errorGenerationFailed'));
    } finally {
      setLoading(false);
    }
  };

  const updateActiveField = (field: string, value: string) => {
    if (!activeSlideKey) return;
    setModifiedSlides((prev) => {
      const prevSlide = prev[activeSlideKey] || {};
      return {
        ...prev,
        [activeSlideKey]: {
          ...prevSlide,
          [field]: value,
        },
      };
    });
  };

  const handleReconfigure = () => {
    setInWorkspace(false);
    // Clear identifiers so next generate always creates a new job.
    setSessionId('');
    setJobId(null);
    completionHandledJobRef.current = null;
  };

  const markRestoredFields = (restoredFields: string[]) => {
    if (restoredFields.length === 0) return;
    setRestoredFieldMarks((prev) => {
      const next = { ...prev };
      restoredFields.forEach((key) => {
        next[key] = true;
      });
      return next;
    });
    window.setTimeout(() => {
      setRestoredFieldMarks((prev) => {
        const next = { ...prev };
        restoredFields.forEach((key) => {
          delete next[key];
        });
        return next;
      });
    }, 1300);
  };

  const showStructuredRestoreToast = () => {
    showToast(
      'warning',
      t(
        'msgStructuredKeysAutoRestored',
        lang === 'zh-CN' ? '检测到结构化字段键名或 JSON 结构被修改，系统已自动恢复。' : 'Structured field keys or JSON structure were changed and have been auto-restored.',
      ),
    );
  };

  const inspectAndRestoreStructuredKeys = (source: ModifiedSlideFields) => {
    if (!slidespec) {
      return { nextModifiedSlides: source, payload: [] as { slide_key: string; new_content: Record<string, unknown> }[], restoredFields: [] as string[] };
    }
    const entries = Object.entries(source);
    const nextModifiedSlides: ModifiedSlideFields = { ...source };
    const restoredFields: string[] = [];
    const payload: { slide_key: string; new_content: Record<string, unknown> }[] = [];
    entries.forEach(([slideKey, editedFields]) => {
      const originalSlide = slidespec.slides.find((s) => s.slide_key === slideKey);
      if (!originalSlide) return;
      const nextContent: Record<string, unknown> = { ...originalSlide.placeholders };
      const nextEdited = { ...(nextModifiedSlides[slideKey] || {}) };
      Object.entries(editedFields).forEach(([field, raw]) => {
        const originalValue = originalSlide.placeholders[field];
        let parsedValue: unknown;
        try {
          parsedValue = parseEditedValue(raw, originalValue, field);
        } catch (error) {
          if (isPlainObject(originalValue) || Array.isArray(originalValue)) {
            nextContent[field] = originalValue;
            nextEdited[field] = stringifyValue(originalValue);
            restoredFields.push(`${slideKey}:${field}`);
            return;
          }
          throw error;
        }
        if (isPlainObject(originalValue) || Array.isArray(originalValue)) {
          if (!hasSameStructuredKeys(originalValue, parsedValue)) {
            const restored = restoreStructuredKeysByTemplate(originalValue, parsedValue);
            nextContent[field] = restored;
            nextEdited[field] = stringifyValue(restored);
            restoredFields.push(`${slideKey}:${field}`);
            return;
          }
        }
        nextContent[field] = parsedValue;
      });
      const changed = JSON.stringify(nextContent) !== JSON.stringify(originalSlide.placeholders || {});
      if (changed) {
        payload.push({
          slide_key: slideKey,
          new_content: nextContent,
        });
        nextModifiedSlides[slideKey] = nextEdited;
      } else {
        delete nextModifiedSlides[slideKey];
      }
    });
    return { nextModifiedSlides, payload, restoredFields };
  };

  useEffect(() => {
    const onOutsidePointerDown = (event: MouseEvent) => {
      if (!activeSlideKey || !slidespec) return;
      const target = event.target as Node | null;
      if (!target) return;
      if (editorPaneRef.current?.contains(target)) return;
      const activeEdited = modifiedSlides[activeSlideKey];
      if (!activeEdited || Object.keys(activeEdited).length === 0) return;
      try {
        const inspected = inspectAndRestoreStructuredKeys({ [activeSlideKey]: activeEdited });
        if (inspected.restoredFields.length > 0) {
          setModifiedSlides((prev) => {
            const next = { ...prev };
            if (inspected.nextModifiedSlides[activeSlideKey]) {
              next[activeSlideKey] = inspected.nextModifiedSlides[activeSlideKey];
            } else {
              delete next[activeSlideKey];
            }
            return next;
          });
          markRestoredFields(inspected.restoredFields);
          showStructuredRestoreToast();
        }
      } catch (error) {
        showToast('error', error instanceof Error ? error.message : t('titleUpdateFailed'));
      }
    };
    document.addEventListener('mousedown', onOutsidePointerDown, true);
    return () => {
      document.removeEventListener('mousedown', onOutsidePointerDown, true);
    };
  }, [activeSlideKey, modifiedSlides, slidespec]);

  useEffect(() => {
    const onGlobalPointerDown = (event: MouseEvent) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (exportOpen && exportPopoverRef.current && !exportPopoverRef.current.contains(target)) {
        setExportOpen(false);
      }
      if (aiModalOpen && aiPopoverRef.current && !aiPopoverRef.current.contains(target)) {
        setAiModalOpen(false);
      }
    };
    const onGlobalKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      if (exportOpen) setExportOpen(false);
      if (aiModalOpen) setAiModalOpen(false);
    };
    document.addEventListener('mousedown', onGlobalPointerDown);
    window.addEventListener('keydown', onGlobalKeyDown);
    return () => {
      document.removeEventListener('mousedown', onGlobalPointerDown);
      window.removeEventListener('keydown', onGlobalKeyDown);
    };
  }, [exportOpen, aiModalOpen]);

  const saveManualChanges = async () => {
    if (!jobId || !slidespec) return;
    if (generationInProgress) {
      showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
      return;
    }
    const entries = Object.entries(modifiedSlides);
    if (entries.length === 0) return showToast('warning', t('msgNoChangesToApply'));
    let nextModifiedSlides: ModifiedSlideFields = { ...modifiedSlides };
    let restoredFields: string[] = [];
    let payload: { slide_key: string; new_content: Record<string, unknown> }[] = [];
    try {
      const inspected = inspectAndRestoreStructuredKeys(modifiedSlides);
      nextModifiedSlides = inspected.nextModifiedSlides;
      restoredFields = inspected.restoredFields;
      payload = inspected.payload;
    } catch (error) {
      showToast('error', error instanceof Error ? error.message : t('titleUpdateFailed'));
      return;
    }
    if (restoredFields.length > 0) {
      setModifiedSlides(nextModifiedSlides);
      markRestoredFields(restoredFields);
      showStructuredRestoreToast();
    }
    if (payload.length === 0) {
      // If structured keys were restored, the restore toast already explains the outcome.
      if (restoredFields.length > 0) return;
      return showToast('warning', t('msgNoChangesToApply'));
    }
    const firstSlideKey = firstSlide(slidespec);
    if (firstSlideKey) {
      setActiveSlideKey(firstSlideKey);
      window.requestAnimationFrame(() => {
        const firstView = document.getElementById(`view-${firstSlideKey}`);
        if (firstView) firstView.scrollIntoView({ block: 'center', behavior: 'smooth' });
      });
    }
    setGenerationInProgress(true);
    actionRef.current = 'rewrite';
    progressValueRef.current = 0;
    setProgress(0);
    setProgressText(t('progressInit', lang === 'zh-CN' ? '初始化' : 'Initializing'));
    setLoading(true);
    try {
      const result = await MainApi.rewriteSlides(jobId, payload, clientId);
      if (result.slidespec) setSlidespec(result.slidespec);
      setModifiedSlides({});
      await syncPreviewUrls(jobId, true);
      progressValueRef.current = 100;
      setProgress(100);
      setProgressText(t('progressComplete', lang === 'zh-CN' ? '完成' : 'Completed'));
      showToast('success', t('msgUpdateSlidesSuccess').replace('{count}', String(result.updated_count || 0)));
    } catch (error) {
      showToast('error', error instanceof Error ? error.message : t('titleUpdateFailed'));
      progressValueRef.current = 0;
      setProgress(0);
      setProgressText('');
    } finally {
      setGenerationInProgress(false);
      setLoading(false);
      actionRef.current = null;
    }
  };

  const openAiRewrite = () => {
    if (!activeSlideKey) return showToast('warning', t('msgSelectSlideForAiRewrite'));
    setExportOpen(false);
    setAiModalOpen(true);
  };

  const submitAiRewrite = async () => {
    if (!jobId || !activeSlideKey) return;
    const normalizedPrompt = aiPrompt.trim();
    if (!normalizedPrompt) return showToast('warning', t('msgEnterAiRewritePrompt'));
    if (selectedAiMode === 'selected' && selectedAiTokens.length === 0) {
      return showToast('warning', t('msgSelectAtLeastOneAiToken'));
    }
    const firstSlideKey = firstSlide(slidespec);
    if (firstSlideKey) {
      setActiveSlideKey(firstSlideKey);
      window.requestAnimationFrame(() => {
        const firstView = document.getElementById(`view-${firstSlideKey}`);
        if (firstView) firstView.scrollIntoView({ block: 'center', behavior: 'smooth' });
      });
    }
    setAiModalOpen(false);
    setGenerationInProgress(true);
    actionRef.current = 'ai-rewrite';
    progressValueRef.current = 0;
    setProgress(0);
    setProgressText(t('progressInit', lang === 'zh-CN' ? '初始化' : 'Initializing'));
    setLoading(true);
    try {
      const result = await MainApi.aiRewriteSlide(
        jobId,
        activeSlideKey,
        normalizedPrompt,
        selectedAiMode === 'selected' ? selectedAiTokens : [],
        clientId,
      );
      if (result.slidespec) setSlidespec(result.slidespec);
      await syncPreviewUrls(jobId, true);
      setRecentAiPromptsBySlide((prev) => {
        const current = ensureArray(prev[activeSlideKey]);
        const merged = [normalizedPrompt, ...current.filter((item) => item !== normalizedPrompt)].slice(0, 3);
        return {
          ...prev,
          [activeSlideKey]: merged,
        };
      });
      setAiPrompt('');
      progressValueRef.current = 100;
      setProgress(100);
      setProgressText(t('progressComplete', lang === 'zh-CN' ? '完成' : 'Completed'));
      if ((result.updated_count || 0) > 0) showToast('success', t('msgAiRewriteSuccess').replace('{slide}', activeSlideKey));
      else showToast('warning', t('msgAiRewriteNoChanges'));
    } catch (error) {
      showToast('error', error instanceof Error ? error.message : t('titleUpdateFailed'));
      progressValueRef.current = 0;
      setProgress(0);
      setProgressText('');
    } finally {
      setGenerationInProgress(false);
      setLoading(false);
      actionRef.current = null;
    }
  };

  const handleRefreshPreview = async () => {
    if (!jobId) return;
    setLoading(true);
    try {
      await syncPreviewUrls(jobId, true);
    } catch (error) {
      showToast('error', error instanceof Error ? error.message : t('errorPreviewFailed'));
    } finally {
      setLoading(false);
    }
  };

  const parseDownloadFilename = (contentDisposition: string | null, fallback: string) => {
    if (!contentDisposition) return fallback;
    const filenameStar = /filename\*=UTF-8''([^;]+)/i.exec(contentDisposition);
    if (filenameStar?.[1]) {
      try {
        return decodeURIComponent(filenameStar[1]);
      } catch {
        return filenameStar[1];
      }
    }
    const filename = /filename="?([^";]+)"?/i.exec(contentDisposition);
    return filename?.[1] || fallback;
  };

  const performExport = async (targetFormat: 'ppt' | 'pdf' = exportFormat) => {
    if (!jobId) return;
    const path = targetFormat === 'pdf' ? 'download-pdf' : 'download';
    const url = `/api/v1/reports/${encodeURIComponent(jobId)}/${path}?regenerate_if_missing=true`;
    setExportOpen(false);
    if (!ratedJobs.has(jobId)) {
      setRatingOpen(true);
      setRatingChoice(null);
      setRatingComment('');
    }
    try {
      const response = await fetch(url, { credentials: 'include' });
      if (!response.ok) {
        const errorText = (await response.text()) || response.statusText || 'Export failed';
        throw new Error(errorText);
      }
      const blob = await response.blob();
      const ext = targetFormat === 'pdf' ? 'pdf' : 'pptx';
      const fallbackFilename = `${jobId}.${ext}`;
      const filename = parseDownloadFilename(response.headers.get('content-disposition'), fallbackFilename);
      const blobUrl = window.URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = blobUrl;
      anchor.download = filename;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => window.URL.revokeObjectURL(blobUrl), 1000);
    } catch (error) {
      showToast(
        'error',
        error instanceof Error
          ? error.message
          : t('errorExportFailed', lang === 'zh-CN' ? '导出失败，请稍后重试' : 'Export failed, please try again'),
      );
    }
  };

  const ratingCommentMinLength = 10;
  const ratingCommentLength = ratingComment.trim().length;
  const ratingRequiresComment = ratingChoice === 'disliked';
  const ratingSubmitDisabled = !ratingChoice || (ratingRequiresComment && ratingCommentLength < ratingCommentMinLength);

  const submitRating = async () => {
    if (!jobId || !ratingChoice) return;
    if (ratingChoice === 'disliked' && ratingComment.trim().length < ratingCommentMinLength) {
      return showToast('warning', t('ratingErrorCommentTooShort'));
    }
    try {
      await MainApi.submitRating(jobId, ratingChoice, ratingComment.trim() || null);
      setRatedJobs((prev) => {
        const next = new Set(prev);
        next.add(jobId);
        return next;
      });
      setRatingOpen(false);
      showToast('success', t('ratingSuccessMessage'));
    } catch (error) {
      showToast('error', error instanceof Error ? error.message : t('msgError'));
    }
  };

  const toggleFocus = (value: FocusValue) => {
    setSelectedFocus((prev) => {
      if (prev.includes(value)) {
        // Keep at least one focus option selected.
        if (prev.length <= 1) return prev;
        return prev.filter((item) => item !== value);
      }
      return [...prev, value];
    });
  };

  const focusSummaryText =
    selectedFocus
      .map((item) => {
        const meta = FOCUS_OPTIONS.find((f) => f.value === item);
        return meta ? t(meta.labelKey) : item;
      })
      .join(' / ') || t('focusSummaryNone');

  const preconfigTemplateCards = templates.length <= 1 ? templates : templates.slice(0, 6);
  const heroTitle = t('preConfigHeroTitle', lang === 'zh-CN' ? '创建报告' : 'Create your report');
  const heroSubtitle = t(
    'preConfigHeroSubtitle',
    lang === 'zh-CN'
      ? '配置数据源与偏好，AI 将生成专业汇报演示稿。'
      : 'Configure your data source and preferences. Our AI will generate a professional, boardroom-ready deck.',
  );
  const readyText = t('preConfigReadyText', lang === 'zh-CN' ? '准备生成？' : 'Ready to generate?');
  const generateCta = t('preConfigGenerateCta', lang === 'zh-CN' ? '生成报告' : 'Generate Report');

  const getTemplateFrameTitle = (frame: TemplateSlideMeta | undefined, index: number) => {
    if (frame?.title && String(frame.title).trim()) return String(frame.title);
    if (frame?.slide_key && String(frame.slide_key).trim()) return String(frame.slide_key);
    return t('preConfigTemplatePageFallback', `Page ${index + 1}`).replace('{index}', String(index + 1));
  };

  const scrubTemplateFrame = (templateId: string, clientX: number, target: EventTarget | null) => {
    if (!target || !(target instanceof HTMLElement)) return;
    const frames = templateSlidesById[templateId] || [];
    const total = Math.max(frames.length, 1);
    const rect = target.getBoundingClientRect();
    if (!rect.width) return;
    const ratio = Math.min(1, Math.max(0, (clientX - rect.left) / rect.width));
    const index = Math.min(total - 1, Math.floor(ratio * total));
    setTemplateFrameIndexById((prev) => {
      if (prev[templateId] === index) return prev;
      return { ...prev, [templateId]: index };
    });
  };

  const onChangeAiMode = (mode: 'all' | 'selected') => {
    if (!activeSlideKey) return;
    setAiModeBySlide((prev) => ({ ...prev, [activeSlideKey]: mode }));
  };

  const toggleAiToken = (token: string) => {
    if (!activeSlideKey) return;
    setAiSelectedTokensBySlide((prev) => {
      const current = ensureArray(prev[activeSlideKey]);
      const exists = current.includes(token);
      const next = exists ? current.filter((x) => x !== token) : [...current, token];
      return {
        ...prev,
        [activeSlideKey]: next,
      };
    });
  };

  const selectAllAiTokens = () => {
    if (!activeSlideKey) return;
    if (activeAiTokens.length === 0) return;
    setAiSelectedTokensBySlide((prev) => {
      const current = new Set(ensureArray(prev[activeSlideKey]));
      activeAiTokens.forEach((token) => current.add(token));
      return {
        ...prev,
        [activeSlideKey]: Array.from(current),
      };
    });
  };

  const clearAiTokenSelection = () => {
    if (!activeSlideKey) return;
    setAiSelectedTokensBySlide((prev) => ({
      ...prev,
      [activeSlideKey]: [],
    }));
  };

  const applyAiPromptPreset = (preset: string) => {
    setAiPrompt((prev) => {
      const lines = prev
        .split('\n')
        .map((line) => line.trim())
        .filter(Boolean);
      const existingIndex = lines.findIndex((line) => line === preset);
      if (existingIndex >= 0) {
        lines.splice(existingIndex, 1);
        return lines.join('\n');
      }
      return lines.length > 0 ? `${lines.join('\n')}\n${preset}` : preset;
    });
  };

  const onAiPromptKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
      event.preventDefault();
      void submitAiRewrite();
    }
  };

  const openPresentation = () => {
    if (generationInProgress || loading) return showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
    if (!slidespec || slidespec.slides.length === 0) return showToast('warning', t('msgNoSlidesForPresentation'));
    setPresentIndex(activeSlideIndex >= 0 ? activeSlideIndex : 0);
    setPresentOpen(true);
  };

  const openPresentationAtSlide = (slideKey: string) => {
    if (generationInProgress || loading) return showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
    if (!slidespec || slidespec.slides.length === 0) return showToast('warning', t('msgNoSlidesForPresentation'));
    const targetIndex = slidespec.slides.findIndex((slide) => slide.slide_key === slideKey);
    setPresentIndex(targetIndex >= 0 ? targetIndex : 0);
    setPresentOpen(true);
  };

  const closePresentation = () => {
    setPresentOpen(false);
    if (document.fullscreenElement) {
      void document.exitFullscreen().catch(() => {});
    }
  };

  const nextPresentation = (delta: number) => {
    if (!slidespec) return;
    setPresentIndex((prev) => {
      const target = prev + delta;
      if (target < 0) return 0;
      if (target > slidespec.slides.length - 1) return slidespec.slides.length - 1;
      return target;
    });
  };

  useEffect(() => {
    if (!presentOpen) return;
    if (presenterRef.current && !document.fullscreenElement) {
      void presenterRef.current.requestFullscreen?.().catch(() => {});
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') closePresentation();
      else if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') nextPresentation(-1);
      else if (event.key === 'ArrowRight' || event.key === 'ArrowDown') nextPresentation(1);
    };
    const onFsChange = () => {
      if (!document.fullscreenElement) setPresentOpen(false);
    };
    window.addEventListener('keydown', onKey);
    document.addEventListener('fullscreenchange', onFsChange);
    return () => {
      window.removeEventListener('keydown', onKey);
      document.removeEventListener('fullscreenchange', onFsChange);
    };
  }, [presentOpen]);

  const currentPresentUrl = presentIndex >= 0 ? previews[presentIndex] : undefined;
  const activeSlideTitle = activeSlide ? (activeSlide.title || activeSlide.slide_key) : '';
  const totalSlideCount = slidespec?.slides.length || 0;
  const currentSlideNumber = activeSlideIndex >= 0 ? activeSlideIndex + 1 : 0;
  const exportSlideCountText = t('exportSlideCount', lang === 'zh-CN' ? '{count} 页' : '{count} slides').replace('{count}', String(totalSlideCount || 0));
  const exportPrimaryLabel = t('btnExportPptx', lang === 'zh-CN' ? '导出PPTX' : 'Export PPTX');
  const exportConfirmLabel =
    exportFormat === 'pdf' ? t('btnExportPdf', lang === 'zh-CN' ? '导出PDF' : 'Export PDF') : t('btnExportPptx', lang === 'zh-CN' ? '导出PPTX' : 'Export PPTX');
  const aiPresetPrompts = [
    t('aiPresetExecutive', lang === 'zh-CN' ? '改成更偏管理层汇报的表达，突出结论和行动。' : 'Rewrite for executive audience with clear conclusions and actions.'),
    t('aiPresetConcise', lang === 'zh-CN' ? '更简洁，减少冗余句，保留关键数据。' : 'Make it concise and keep only key metrics.'),
    t('aiPresetProfessional', lang === 'zh-CN' ? '语气更专业，术语统一，结构更清晰。' : 'Use a more professional tone and clearer structure.'),
  ];
  const aiPromptPresetSet = useMemo(() => {
    const lines = aiPrompt
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean);
    return new Set(lines);
  }, [aiPrompt]);
  const aiSlideLabel = activeSlide ? `${currentSlideNumber}/${totalSlideCount} · ${activeSlide.title || activeSlide.slide_key}` : '-';
  const loadingPlaceholderCount = Math.max(previewLayout === 'two-up' ? 8 : 4, (templateSlidesById[selectedTemplate] || []).length || 0);
  const previewToggleTo = previewLayout === 'list' ? 'two-up' : 'list';
  const previewToggleTooltip = previewLayout === 'list' ? t('labelPreviewTwoUp') : t('labelPreviewList');

  const getPreviewAnchorSlideKey = (): string | null => {
    const items = Array.from(document.querySelectorAll<HTMLElement>('.preview-card[id^="view-"]'));
    if (items.length === 0) return activeSlideKey;
    const pane = document.querySelector<HTMLElement>('.preview-pane');
    if (!pane) return activeSlideKey;
    const paneRect = pane.getBoundingClientRect();
    const centerY = paneRect.top + paneRect.height / 2;
    let anchorKey: string | null = activeSlideKey;
    let minDistance = Number.POSITIVE_INFINITY;
    items.forEach((item) => {
      const rect = item.getBoundingClientRect();
      const distance = Math.abs(rect.top + rect.height / 2 - centerY);
      if (distance < minDistance) {
        minDistance = distance;
        const id = item.id || '';
        anchorKey = id.startsWith('view-') ? id.slice(5) : anchorKey;
      }
    });
    return anchorKey;
  };

  const togglePreviewLayout = () => {
    const anchor = getPreviewAnchorSlideKey();
    setPreviewLayout(previewToggleTo);
    if (anchor) setActiveSlideKey(anchor);
  };

  const focusSlidePreview = (slideKey: string, behavior: ScrollBehavior = 'auto') => {
    setActiveSlideKey(slideKey);
    const view = document.getElementById(`view-${slideKey}`);
    if (view) view.scrollIntoView({ block: 'center', behavior });
  };

  const stepActiveSlide = (delta: number) => {
    if (!slidespec || slidespec.slides.length === 0) return;
    const current = activeSlideIndex >= 0 ? activeSlideIndex : 0;
    const next = Math.min(slidespec.slides.length - 1, Math.max(0, current + delta));
    const nextKey = slidespec.slides[next]?.slide_key;
    if (!nextKey) return;
    focusSlidePreview(nextKey);
  };

  const onPreviewWheel = (event: React.WheelEvent<HTMLElement>) => {
    if (!slidespec || slidespec.slides.length === 0) return;
    if (previewLayout === 'list') return;
    const now = Date.now();
    if (Math.abs(event.deltaY) < 2) return;
    event.preventDefault();
    if (now - previewWheelAtRef.current > 280) previewWheelAccumRef.current = 0;
    previewWheelAtRef.current = now;
    previewWheelAccumRef.current += event.deltaY;
    const threshold = 220;
    if (Math.abs(previewWheelAccumRef.current) < threshold) return;
    const steps = Math.trunc(previewWheelAccumRef.current / threshold);
    previewWheelAccumRef.current -= steps * threshold;
    stepActiveSlide(steps);
  };

  const onPreviewScroll = () => {
    if (previewLayout !== 'list') return;
    if (previewScrollRafRef.current) window.cancelAnimationFrame(previewScrollRafRef.current);
    previewScrollRafRef.current = window.requestAnimationFrame(() => {
      previewScrollRafRef.current = null;
      const anchor = getPreviewAnchorSlideKey();
      if (anchor && anchor !== activeSlideKey) setActiveSlideKey(anchor);
    });
  };

  const onPresenterWheel = (event: React.WheelEvent<HTMLDivElement>) => {
    if (!presentOpen || !slidespec || slidespec.slides.length === 0) return;
    const now = Date.now();
    if (Math.abs(event.deltaY) < 10) return;
    if (now - previewWheelAtRef.current < 180) return;
    previewWheelAtRef.current = now;
    event.preventDefault();
    nextPresentation(event.deltaY > 0 ? 1 : -1);
  };

  const getPreviewImgKey = (channel: 'thumb' | 'main' | 'present', idx: number, url?: string) => `${channel}:${idx}:${url || ''}`;
  const hasPreviewImgError = (channel: 'thumb' | 'main' | 'present', idx: number, url?: string) =>
    Boolean(previewLoadErrorKeys[getPreviewImgKey(channel, idx, url)]);
  const markPreviewImgError = (channel: 'thumb' | 'main' | 'present', idx: number, url?: string) => {
    const key = getPreviewImgKey(channel, idx, url);
    setPreviewLoadErrorKeys((prev) => (prev[key] ? prev : { ...prev, [key]: true }));
  };

  useEffect(() => {
    setPreviewLoadErrorKeys({});
  }, [previews]);

  useEffect(() => {
    if (!activeSlideKey) return;
    const thumb = document.getElementById(`thumb-${activeSlideKey}`);
    if (thumb) thumb.scrollIntoView({ block: 'nearest' });
  }, [activeSlideKey]);

  useEffect(() => {
    return () => {
      if (previewScrollRafRef.current) window.cancelAnimationFrame(previewScrollRafRef.current);
    };
  }, []);

  return (
    <div className="rg-root">
      <div className="rg-toast-container">
        {toasts.map((toast) => (
          <div className={`rg-toast rg-toast-${toast.type}`} key={toast.id}>
            <span className="rg-toast-icon">
              {toast.type === 'success' ? <CheckCircle2 size={15} /> : null}
              {toast.type === 'error' ? <AlertCircle size={15} /> : null}
              {toast.type === 'warning' ? <AlertTriangle size={15} /> : null}
              {toast.type === 'info' ? <Info size={15} /> : null}
            </span>
            <span className="rg-toast-main">
              <strong>{getToastTitle(toast.type)}</strong>
              <small>{toast.message}</small>
            </span>
            <button onClick={() => setToasts((prev) => prev.filter((item) => item.id !== toast.id))}>
              <X size={14} />
            </button>
          </div>
        ))}
      </div>

      {!inWorkspace ? (
        <section className="preconfig-page">
          <header className="preconfig-header">
            <div>
              <div className="preconfig-badge">
                <Sparkles size={14} />
                {t('preConfigStudioBadge', 'AI Presentation Studio')}
              </div>
              <h1>{heroTitle}</h1>
              <p>{heroSubtitle}</p>
            </div>
            <button className="lang-btn" onClick={toggleLang} title={lang === 'zh-CN' ? 'Switch to English' : '切换到中文'}>
              <Globe size={16} /> {lang === 'zh-CN' ? 'EN' : '中文'}
            </button>
          </header>

          <div className="preconfig-grid">
            <div className="preconfig-left-col">
              <section className="preconfig-section">
                <h2>
                  <FileSpreadsheet size={16} className="section-icon" /> 1. {t('preConfigSectionData', 'Data Source')}
                </h2>
                <div className="data-source-card">
                  <div className="data-source-icon">
                    <UploadCloud size={30} />
                  </div>
                  <h3>{t('preConfigDataCardTitle', 'Select Data Source')}</h3>
                  <p>{t('preConfigDataCardDesc', 'Choose your Excel or CSV input data')}</p>
                  <div className="data-source-controls">
                    <select value={selectedInput} onChange={(e) => setSelectedInput(e.target.value)}>
                      <option value="">{t('msgSelectDataSource', 'Select data source...')}</option>
                      {inputs.map((input) => (
                        <option key={input.id} value={input.id}>
                          {input.description || input.id}
                        </option>
                      ))}
                    </select>
                    <label className="check-row">
                      <input type="checkbox" checked={useMock} onChange={(e) => setUseMock(e.target.checked)} />
                      <span>{t('labelUseMock')}</span>
                    </label>
                  </div>
                </div>
              </section>

              <section className="preconfig-section">
                <h2>
                  <Settings2 size={16} className="section-icon" /> 2. {t('preConfigFocusTitle', 'Report Focus')}
                </h2>
                <div className="focus-grid">
                  {FOCUS_OPTIONS.map((item) => {
                    const selected = selectedFocus.includes(item.value);
                    return (
                      <button key={item.value} className={`focus-card ${selected ? 'selected' : ''}`} onClick={() => toggleFocus(item.value)}>
                        <span className="focus-title">{t(item.labelKey)}</span>
                        <span className="focus-desc">{t(item.descKey)}</span>
                        <span className="focus-check">{selected ? <CheckCircle2 size={12} /> : null}</span>
                      </button>
                    );
                  })}
                </div>
              </section>
            </div>

            <div className="preconfig-right-col">
              <section className="preconfig-section preconfig-template-section">
                <h2>
                  <LayoutTemplate size={16} className="section-icon" /> 3. {t('preConfigSectionTemplate', 'Select Template')}
                </h2>
                <div className={`template-gallery ${preconfigTemplateCards.length === 1 ? 'single' : ''}`}>
                  {preconfigTemplateCards.map((tpl, idx) => {
                    const selected = selectedTemplate === tpl.template_id;
                    const frames = templateSlidesById[tpl.template_id] || [];
                    const total = Math.max(frames.length, 1);
                    const frameIndex = Math.min(templateFrameIndexById[tpl.template_id] || 0, total - 1);
                    const imageUrl = getTemplatePreviewImageUrl(tpl.template_id, frameIndex);
                    const fixedTemplateName = t('preConfigTemplateFixedName', lang === 'zh-CN' ? '经典模板' : 'Classic Template');
                    const slidesCountText = t('preConfigSlidesCount', '{count} slides').replace('{count}', String(total));
                    return (
                      <button
                        key={tpl.template_id}
                        className={`template-card ${selected ? 'selected' : ''}`}
                        onClick={() => setSelectedTemplate(tpl.template_id)}
                      >
                        <div
                          className="template-preview"
                          onPointerEnter={() => void loadTemplateSlides(tpl.template_id)}
                          onPointerMove={(event) => scrubTemplateFrame(tpl.template_id, event.clientX, event.currentTarget)}
                        >
                          {imageUrl ? (
                            <img
                              src={`${imageUrl}?v=${templatePreviewStamp}&p=${frameIndex + 1}`}
                              alt={t('preConfigTemplateImageAlt', 'Template page {page} thumbnail').replace('{page}', String(frameIndex + 1))}
                            />
                          ) : (
                            <img src={`/placeholders/template-${(idx % 4) + 1}.svg`} alt={toTemplateName(tpl)} />
                          )}
                          <div className="template-preview-meta">
                            <div className="template-meta-left">
                              <span className="template-fixed-name">{fixedTemplateName}</span>
                              <span className="template-slides-count">{slidesCountText}</span>
                            </div>
                          </div>
                        </div>
                        {selected ? (
                          <span className="template-selected-check">
                            <CheckCircle2 size={16} />
                          </span>
                        ) : null}
                      </button>
                    );
                  })}
                </div>
              </section>
            </div>
          </div>

          <footer className="preconfig-footer">
            <span>{readyText}</span>
            <button
              className="preconfig-generate-btn"
              onClick={() => void beginGenerate()}
              disabled={loading || !selectedInput || !selectedTemplate}
            >
              {loading ? (
                <>
                  <RefreshCw className="spin" size={16} />
                  {t('msgGenerating', lang === 'zh-CN' ? '生成中...' : 'Generating...')}
                </>
              ) : (
                <>
                  <Sparkles size={16} />
                  {generateCta}
                  <ChevronRight size={16} className="preconfig-generate-arrow" />
                </>
              )}
            </button>
          </footer>
        </section>
      ) : (
        <section className="workspace-page">
          <header className="workspace-header">
            <div className="header-left">
              <button
                className="workspace-icon-btn tooltip-card-btn"
                onClick={handleReconfigure}
                disabled={loading || generationInProgress}
                data-tooltip={t('btnReconfigure')}
                aria-label={t('btnReconfigure')}
              >
                <ChevronLeft size={16} />
              </button>
              <div className="title-wrap">
                <FileText size={15} />
                <strong>{jobId ? `${jobId}.pptx` : 'MSS_Report_Draft.pptx'}</strong>
                <span>{generationInProgress ? t('statusProcessing') : t('statusCompleted')}</span>
              </div>
              <div className="focus-badge">
                {t('focusSummaryLabel')}: {focusSummaryText}
              </div>
            </div>
            <div className="header-actions">
              <button
                className="btn-secondary workspace-top-action tooltip-card-btn"
                onClick={toggleLang}
                data-tooltip={lang === 'zh-CN' ? t('tooltipSwitchToEnglish', '切换到英文') : t('tooltipSwitchToChinese', '切换到中文')}
                aria-label={lang === 'zh-CN' ? t('tooltipSwitchToEnglish', '切换到英文') : t('tooltipSwitchToChinese', '切换到中文')}
              >
                <Globe size={14} /> {lang === 'zh-CN' ? 'EN' : '中文'}
              </button>
              <button
                className="workspace-icon-btn tooltip-card-btn preview-layout-toggle"
                onClick={togglePreviewLayout}
                data-tooltip={previewToggleTooltip}
                aria-label={previewToggleTooltip}
              >
                {previewLayout === 'list' ? <LayoutGrid size={16} /> : <List size={16} />}
              </button>
              <button
                className="btn-secondary workspace-top-action tooltip-card-btn"
                onClick={() => void handleRefreshPreview()}
                disabled={!jobId || loading}
                data-tooltip={t('btnRefreshPreview', '刷新预览')}
                aria-label={t('btnRefreshPreview', '刷新预览')}
              >
                <RefreshCw size={14} />
              </button>
              <div className={`workspace-action-anchor ${aiModalOpen ? 'open' : ''}`} ref={aiPopoverRef}>
                <button
                  className="btn-secondary workspace-top-action tooltip-card-btn"
                  onClick={() => {
                    if (aiModalOpen) setAiModalOpen(false);
                    else openAiRewrite();
                  }}
                  disabled={!jobId || !activeSlideKey || loading}
                  data-tooltip={t('btnAiRewriteCurrentSlide', 'AI 重写当前页')}
                  aria-label={t('btnAiRewriteCurrentSlide', 'AI 重写当前页')}
                >
                  <Sparkles size={14} />
                </button>
                {aiModalOpen && (
                  <div className="workspace-popover ai-rewrite-popover" role="dialog" aria-label={t('btnAiRewriteCurrentSlide', 'AI 重写当前页')}>
                    <div className="ai-popover-head">
                      <h3>{t('btnAiRewriteCurrentSlide')}</h3>
                      <p>{t('msgAiRewriteHint')}</p>
                      <div className="ai-popover-meta">
                        <span>{aiSlideLabel}</span>
                        <span>
                          {t('aiRewriteAffectCount', lang === 'zh-CN' ? '预计影响 {count} 个字段' : 'Estimated impact: {count} fields').replace(
                            '{count}',
                            String(aiRewriteTargetCount),
                          )}
                        </span>
                      </div>
                    </div>
                    <div className="mode-switch">
                      <button className={selectedAiMode === 'all' ? 'active' : ''} onClick={() => onChangeAiMode('all')}>
                        {t('aiRewriteModeAll')}
                      </button>
                      <button className={selectedAiMode === 'selected' ? 'active' : ''} onClick={() => onChangeAiMode('selected')}>
                        {t('aiRewriteModeSelected')}
                      </button>
                    </div>
                    {selectedAiMode === 'selected' && (
                      <div className="token-panel">
                        <div className="token-toolbar">
                          <span className="token-count">
                            {t('aiTokenCount', lang === 'zh-CN' ? '可选字段 {count}' : '{count} fields').replace('{count}', String(activeAiTokens.length))}
                          </span>
                          <button type="button" className="token-mini-btn" onClick={selectAllAiTokens}>
                            {t('aiTokenSelectAll', lang === 'zh-CN' ? '全选' : 'Select all')}
                          </button>
                          <button type="button" className="token-mini-btn" onClick={clearAiTokenSelection}>
                            {t('aiTokenClear', lang === 'zh-CN' ? '清空' : 'Clear')}
                          </button>
                        </div>
                        <div className="token-chip-grid">
                          {activeAiTokens.length === 0 && (
                            <div className="ai-empty-hint">{t('msgNoData')}</div>
                          )}
                          {activeAiTokens.map((token) => (
                            <button
                              type="button"
                              key={token}
                              className={`token-chip ${selectedAiTokens.includes(token) ? 'active' : ''}`}
                              onClick={() => toggleAiToken(token)}
                            >
                              <code>{token}</code>
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                    <div className="ai-summary-card">
                      <strong>{t('aiRewriteSummaryTitle', lang === 'zh-CN' ? '改写范围摘要' : 'Rewrite scope')}</strong>
                      <small>
                        {selectedAiMode === 'all'
                          ? t('aiRewriteSummaryAll', lang === 'zh-CN' ? '将按提示词尝试优化当前页全部可重写字段。' : 'AI will optimize all rewriteable fields on this slide.')
                          : t('aiRewriteSummarySelected', lang === 'zh-CN' ? '将仅改写你选中的字段。' : 'Only selected fields will be rewritten.')}
                      </small>
                    </div>
                    <div className="ai-preset-row">
                      {aiPresetPrompts.map((preset) => (
                        <button
                          key={preset}
                          type="button"
                          className={`ai-preset-btn ${aiPromptPresetSet.has(preset) ? 'active' : ''}`}
                          onClick={() => applyAiPromptPreset(preset)}
                        >
                          {preset}
                        </button>
                      ))}
                    </div>
                    {recentAiPrompts.length > 0 && (
                      <div className="ai-recent-row">
                        <span>{t('aiRecentPrompts', lang === 'zh-CN' ? '最近使用' : 'Recent')}</span>
                        {recentAiPrompts.map((item) => (
                          <button key={item} type="button" className="ai-recent-btn" onClick={() => setAiPrompt(item)}>
                            {item}
                          </button>
                        ))}
                      </div>
                    )}
                    <textarea
                      value={aiPrompt}
                      onChange={(e) => setAiPrompt(e.target.value)}
                      onKeyDown={onAiPromptKeyDown}
                      rows={6}
                      placeholder={t('placeholderAiRewritePrompt')}
                    />
                    <div className="ai-shortcut-hint">{t('aiShortcutSubmit', lang === 'zh-CN' ? '快捷提交：Ctrl/⌘ + Enter' : 'Shortcut: Ctrl/⌘ + Enter')}</div>
                    <div className="popover-actions">
                      <button className="btn-secondary" onClick={() => setAiModalOpen(false)}>
                        {t('btnCancel')}
                      </button>
                      <button className="btn-primary btn-b-accent" onClick={() => void submitAiRewrite()}>
                        {t('btnAiRewriteCurrentSlide')}
                      </button>
                    </div>
                  </div>
                )}
              </div>
              <button
                className="btn-secondary workspace-top-action tooltip-card-btn"
                onClick={openPresentation}
                disabled={!jobId || previews.length === 0 || loading || generationInProgress}
                data-tooltip={t('presentBtnTitle', '全屏播放幻灯片')}
                aria-label={t('presentBtnTitle', '全屏播放幻灯片')}
              >
                <Play size={14} />
              </button>
              <div className={`workspace-action-anchor ${exportOpen ? 'open' : ''}`} ref={exportPopoverRef}>
                <div className="workspace-export-split">
                  <button
                    className="workspace-export-btn workspace-export-main"
                    onClick={() => {
                      setExportFormat('ppt');
                      performExport('ppt');
                    }}
                    disabled={!jobId || generationInProgress}
                  >
                    <Download size={14} /> {exportPrimaryLabel}
                  </button>
                  <button
                    className="workspace-export-btn workspace-export-toggle"
                    onClick={() => {
                      setAiModalOpen(false);
                      setExportOpen((prev) => !prev);
                    }}
                    disabled={!jobId || generationInProgress}
                    aria-label={t('msgChooseExportFormat')}
                  >
                    <ChevronDown size={14} />
                  </button>
                </div>
                {exportOpen && (
                  <div className="workspace-popover export-popover" role="dialog" aria-label={t('msgChooseExportFormat')}>
                    <div className="export-popover-head">
                      <h3>{t('msgChooseExportFormat')}</h3>
                      <p>{jobId ? `${jobId}.pptx` : 'MSS_Report_Draft.pptx'} · {exportSlideCountText}</p>
                    </div>
                    <div className="export-format-grid">
                      <button className={`export-format-card ${exportFormat === 'ppt' ? 'active' : ''}`} onClick={() => setExportFormat('ppt')}>
                        <span className="format-badge">PPTX</span>
                        <span className="format-main">
                          <strong>{t('btnExportPptx', lang === 'zh-CN' ? '导出PPTX' : 'Export PPTX')}</strong>
                          <small>{lang === 'zh-CN' ? '推荐，保留动画与可编辑内容' : 'Recommended, keeps editable slides and animation'}</small>
                        </span>
                        <span className="format-tag">{lang === 'zh-CN' ? '推荐' : 'Recommended'}</span>
                      </button>
                      <button className={`export-format-card ${exportFormat === 'pdf' ? 'active' : ''}`} onClick={() => setExportFormat('pdf')}>
                        <span className="format-badge pdf">PDF</span>
                        <span className="format-main">
                          <strong>{t('btnExportPdf', lang === 'zh-CN' ? '导出PDF' : 'Export PDF')}</strong>
                          <small>{lang === 'zh-CN' ? '适合跨平台分享与打印' : 'Best for sharing and printing across platforms'}</small>
                        </span>
                      </button>
                    </div>
                    <div className="popover-actions">
                      <button className="btn-secondary" onClick={() => setExportOpen(false)}>
                        {t('btnCancel')}
                      </button>
                      <button className="btn-primary" onClick={performExport}>
                        {exportConfirmLabel}
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </header>

          <div className={`workspace-body ${previewLayout === 'two-up' ? 'two-up-mode' : ''}`}>
            <aside className="slide-sidebar">
              <div className="sidebar-title">{t('workspaceSlidesLabel', 'SLIDES')}</div>
              <div className="thumb-list">
                {ensureArray(slidespec?.slides).length > 0 ? (
                  ensureArray(slidespec?.slides).map((slide, idx) => (
                    <button
                      id={`thumb-${slide.slide_key}`}
                      key={slide.slide_key}
                      className={`thumb-card ${slide.slide_key === activeSlideKey ? 'active' : ''}`}
                      onClick={() => focusSlidePreview(slide.slide_key, 'smooth')}
                    >
                      <div className="thumb-index">{idx + 1}</div>
                      {previews[idx] && !hasPreviewImgError('thumb', idx, previews[idx]) ? (
                        <img
                          src={`${previews[idx]}?v=${Date.now()}`}
                          alt={slide.title || slide.slide_key}
                          onError={() => markPreviewImgError('thumb', idx, previews[idx])}
                        />
                      ) : generationInProgress ? (
                        <div className="thumb-placeholder loading">
                          <Funnel size={14} />
                          <span>{t('msgGeneratingPreview', lang === 'zh-CN' ? '预览生成中...' : 'Generating preview...')}</span>
                        </div>
                      ) : (
                        <div className="thumb-placeholder" />
                      )}
                      <div className="thumb-title">{slide.title || slide.slide_key}</div>
                    </button>
                  ))
                ) : generationInProgress ? (
                  Array.from({ length: loadingPlaceholderCount }).map((_, idx) => (
                    <div className="thumb-card loading" key={`thumb-loading-${idx}`}>
                      <div className="thumb-index">{idx + 1}</div>
                      <div className="thumb-placeholder loading">
                        <Funnel size={14} />
                        <span>{t('msgGeneratingPreview', lang === 'zh-CN' ? '预览生成中...' : 'Generating preview...')}</span>
                      </div>
                    </div>
                  ))
                ) : null}
              </div>
            </aside>

            <main className="preview-pane" onWheel={onPreviewWheel} onScroll={onPreviewScroll}>
              <div className={`preview-list ${previewLayout}`}>
                {ensureArray(slidespec?.slides).length > 0 ? (
                  ensureArray(slidespec?.slides).map((slide, idx) => (
                    <article
                      id={`view-${slide.slide_key}`}
                      key={slide.slide_key}
                      className={`preview-card ${slide.slide_key === activeSlideKey ? 'active' : ''}`}
                      onClick={() => setActiveSlideKey(slide.slide_key)}
                      onDoubleClick={() => {
                        setActiveSlideKey(slide.slide_key);
                        openPresentationAtSlide(slide.slide_key);
                      }}
                    >
                      {previews[idx] && !hasPreviewImgError('main', idx, previews[idx]) ? (
                        <img
                          src={`${previews[idx]}?v=${Date.now()}`}
                          alt={slide.title || slide.slide_key}
                          onError={() => markPreviewImgError('main', idx, previews[idx])}
                        />
                      ) : (
                        <div className="preview-placeholder">{t('msgGeneratingPreview')}</div>
                      )}
                    <button
                      className="float-ai"
                      onClick={(e) => {
                        e.stopPropagation();
                        setActiveSlideKey(slide.slide_key);
                        openAiRewrite();
                      }}
                    >
                      <Sparkles size={12} /> {t('btnAiRewriteShort', lang === 'zh-CN' ? 'AI重写' : 'AI Rewrite')}
                    </button>
                    </article>
                  ))
                ) : generationInProgress ? (
                  Array.from({ length: loadingPlaceholderCount }).map((_, idx) => (
                    <article className="preview-card loading" key={`preview-loading-${idx}`}>
                      <div className="preview-placeholder loading">
                        <Funnel size={18} />
                        <span>{t('msgGeneratingPreview', lang === 'zh-CN' ? '预览生成中...' : 'Generating preview...')}</span>
                      </div>
                    </article>
                  ))
                ) : (
                  <div className="empty-preview">{t('msgNoPreview')}</div>
                )}
              </div>
              <div className="preview-page-bar">
                <span className="preview-page-num">
                  {currentSlideNumber > 0
                    ? `${t('msgPageNumber')} ${currentSlideNumber} / ${totalSlideCount}`
                    : `${t('msgPageNumber')} - / ${totalSlideCount}`}
                </span>
                <span className="preview-page-title">{activeSlideTitle || '-'}</span>
              </div>
            </main>

            <aside className="editor-pane" ref={editorPaneRef}>
              <div className="editor-header">
                <h3>
                  <Edit3 size={15} /> {t('editorTitle')}
                </h3>
              </div>
              <div className="editor-fields">
                {activeSlide ? (
                  activeFieldKeys.map((field) => {
                    const value = activeEditorFields[field] ?? '';
                    const useTextArea = value.length > 80 || value.includes('\n') || value.startsWith('{') || value.startsWith('[');
                    const restoreKey = `${activeSlide.slide_key}:${field}`;
                    return (
                      <label className={`field-card ${restoredFieldMarks[restoreKey] ? 'restored' : ''}`} key={field}>
                        <span>{field}</span>
                        {useTextArea ? (
                          <textarea value={value} onChange={(e) => updateActiveField(field, e.target.value)} rows={6} />
                        ) : (
                          <input value={value} onChange={(e) => updateActiveField(field, e.target.value)} />
                        )}
                      </label>
                    );
                  })
                ) : (
                  <div className="empty-editor">{t('editorPlaceholder')}</div>
                )}
              </div>
              <div className="editor-footer">
                <button className="btn-primary btn-b-accent" onClick={() => void saveManualChanges()} disabled={!jobId || loading}>
                  {loading ? <RefreshCw className="spin" size={14} /> : <RefreshCw size={14} />}
                  {t('btnApplyRegenerate')}
                </button>
              </div>
            </aside>
          </div>
        </section>
      )}

      {ratingOpen && (
        <div className="rating-modal-overlay">
          <div className="rating-modal" role="dialog" aria-modal="false" aria-label={t('ratingTitle')}>
            <button className="rating-modal-close" onClick={() => setRatingOpen(false)} aria-label={t('btnCancel')}>
              <X size={16} />
            </button>
            <div className="rating-modal-head">
              <h4>{t('ratingTitle')}</h4>
              <p>{t('ratingPrompt')}</p>
              {jobId && <span className="rating-modal-job">{jobId}</span>}
            </div>
            <div className="rating-choice-grid">
              <button className={`rating-choice-card positive ${ratingChoice === 'liked' ? 'active' : ''}`} onClick={() => setRatingChoice('liked')}>
                <CheckCircle2 size={16} />
                <span>{t('ratingLiked', lang === 'zh-CN' ? '满意' : 'Satisfied')}</span>
                <small>{lang === 'zh-CN' ? '质量达预期，可直接使用' : 'High quality and ready to use'}</small>
              </button>
              <button className={`rating-choice-card negative ${ratingChoice === 'disliked' ? 'active' : ''}`} onClick={() => setRatingChoice('disliked')}>
                <AlertTriangle size={16} />
                <span>{t('ratingDisliked', lang === 'zh-CN' ? '不满意' : 'Needs improvement')}</span>
                <small>{lang === 'zh-CN' ? '结构或内容仍需优化' : 'Structure or content needs tuning'}</small>
              </button>
            </div>
            {ratingRequiresComment && (
              <div className="rating-feedback">
                <label htmlFor="rating-comment">
                  {lang === 'zh-CN' ? '请告诉我们哪里需要改进' : 'Tell us what should be improved'}
                </label>
                <textarea
                  id="rating-comment"
                  value={ratingComment}
                  onChange={(e) => setRatingComment(e.target.value)}
                  placeholder={t('ratingCommentPlaceholder')}
                  rows={4}
                />
                <div className={`rating-feedback-meta ${ratingCommentLength < ratingCommentMinLength ? 'warning' : ''}`}>
                  {lang === 'zh-CN'
                    ? `至少 ${ratingCommentMinLength} 个字，当前 ${ratingCommentLength} 个字`
                    : `At least ${ratingCommentMinLength} chars, current ${ratingCommentLength}`}
                </div>
              </div>
            )}
            <div className="rating-actions">
              <button className="btn-secondary" onClick={() => setRatingOpen(false)}>
                {t('btnCancel')}
              </button>
              <button className="btn-primary" onClick={() => void submitRating()} disabled={ratingSubmitDisabled}>
                {t('ratingSubmit')}
              </button>
            </div>
          </div>
        </div>
      )}

      {presentOpen && (
        <div className="presenter" onClick={closePresentation} ref={presenterRef} onWheel={onPresenterWheel}>
          <button
            className="present-btn left"
            onClick={(e) => {
              e.stopPropagation();
              nextPresentation(-1);
            }}
          >
            <ChevronLeft size={20} />
          </button>
          <div className="present-body" onClick={(e) => e.stopPropagation()}>
            {currentPresentUrl && !hasPreviewImgError('present', presentIndex, currentPresentUrl) ? (
              <img src={`${currentPresentUrl}?v=${Date.now()}`} alt="slide" onError={() => markPreviewImgError('present', presentIndex, currentPresentUrl)} />
            ) : (
              <div className="present-empty">{t('msgNoPreview')}</div>
            )}
            <div className="present-foot">
              {t('msgPageNumber')} {presentIndex + 1} / {slidespec?.slides.length || 0}
            </div>
          </div>
          <button
            className="present-btn right"
            onClick={(e) => {
              e.stopPropagation();
              nextPresentation(1);
            }}
          >
            <ChevronRight size={20} />
          </button>
          <button className="present-close" onClick={closePresentation}>
            <Maximize size={16} /> {t('msgFullscreenExitHint')}
          </button>
        </div>
      )}

      {generationInProgress && (
        <div className="progress-float">
          <div className="progress-float-title">
            <RefreshCw size={13} className="spin" /> {lt('msgGenerating', '正在生成报告', 'Generating report')}
          </div>
          <div className="progress-float-stage">{progressText || t('progressInit', lang === 'zh-CN' ? '初始化' : 'Initializing')}</div>
          <div className="progress-float-rail">
            <span style={{ width: `${Math.max(0, Math.min(100, progress))}%` }} />
          </div>
          <div className="progress-float-value">{Math.round(progress)}%</div>
        </div>
      )}

    </div>
  );
}
