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
const TIMEOUT_ERROR_CODES = new Set(['LLM_TIMEOUT_EXHAUSTED', 'LLM_UPSTREAM_CONNECTION_FAILED']);
const TASK_SUPERSEDED_ERROR_CODE = 'TASK_SUPERSEDED';
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

interface ModifiedSlideSummary {
  slideKey: string;
  slideNumber: number;
  title: string;
  fieldCount: number;
}

function toTemplateName(tpl: TemplateItem): string {
  if (tpl.template_id === 'mss_classic_ops_2') {
    return '价值复盘';
  }
  if (typeof tpl.name === 'string' && tpl.name.trim()) {
    return tpl.name;
  }
  return tpl.template_id;
}

function getDefaultInputIdForTemplate(templateId: string, inputs: InputItem[]): string {
  const matched = inputs.find((item) => item.template_id === templateId);
  if (matched?.id) return matched.id;

  const fallbackByTemplate: Record<string, string> = {
    mss_classic_ops: 'classic_ops_dataxlsx',
    mss_classic_ops_2: 'plus_ops_dataxlsx',
  };
  const fallback = fallbackByTemplate[templateId];
  if (fallback && inputs.some((item) => item.id === fallback)) return fallback;

  return inputs[0]?.id || '';
}

function getTemplatePreviewImageUrl(templateId: string, frameIndex: number) {
  const folderByTemplate: Record<string, string> = {
    mss_classic_ops: '/static/previews/template_classic_preview',
    mss_classic_ops_2: '/static/previews/template_classic_ops_2_preview',
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
  const [hoverTemplate, setHoverTemplate] = useState<string>('');
  const [selectedInput, setSelectedInput] = useState('');
  const [uploadingExcel, setUploadingExcel] = useState(false);
  const [excelDragActive, setExcelDragActive] = useState(false);
  const [useMock, setUseMock] = useState(false);
  const [selectedFocus, setSelectedFocus] = useState<FocusValue[]>(['business_protection']);

  const [inWorkspace, setInWorkspace] = useState(false);
  const [loading, setLoading] = useState(false);
  const [wsReady, setWsReady] = useState(false);
  const [generationInProgress, setGenerationInProgress] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');

  const [jobId, setJobId] = useState<string | null>(null);
  const [sessionId, setSessionId] = useState<string>('');
  const [clientId, setClientId] = useState<string>('');
  const [activeJobId, setActiveJobId] = useState<string | null>(null);
  const [activeSessionId, setActiveSessionId] = useState<string>('');
  const [activeOperationToken, setActiveOperationToken] = useState<string>('');
  const [timeoutDialogOpen, setTimeoutDialogOpen] = useState(false);
  const [timeoutDialogErrorCode, setTimeoutDialogErrorCode] = useState<string>('');
  const [timeoutDialogMessage, setTimeoutDialogMessage] = useState<string>('');

  const [slidespec, setSlidespec] = useState<SlideSpec | null>(null);
  const [previews, setPreviews] = useState<string[]>([]);
  const [previewStamp, setPreviewStamp] = useState<number>(() => Date.now());
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
  const editorPaneRef = useRef<HTMLElement | null>(null);
  const exportPopoverRef = useRef<HTMLDivElement | null>(null);
  const aiPopoverRef = useRef<HTMLDivElement | null>(null);
  const excelUploadInputRef = useRef<HTMLInputElement | null>(null);
  const excelDragDepthRef = useRef(0);

  const wsRef = useRef<WebSocket | null>(null);
  const wsRetryRef = useRef<number | null>(null);
  const pollRef = useRef<number | null>(null);
  const progressValueRef = useRef<number>(0);
  const actionRef = useRef<'generate' | 'rewrite' | 'ai-rewrite' | null>(null);
  const completionHandledJobRef = useRef<string | null>(null);
  const statusPollTokenRef = useRef<string>('');
  const activeSessionIdRef = useRef<string>('');
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

  const getTemplateSlideMeta = (slideKey?: string | null) => {
    if (!selectedTemplate || !slideKey) return null;
    return (templateSlidesById[selectedTemplate] || []).find((slide) => slide.slide_key === slideKey) || null;
  };

  const formatSlideLabel = (slideKey?: string | null, fallbackTitle?: string | null) => {
    const normalizedKey = typeof slideKey === 'string' ? slideKey.trim() : '';
    const normalizedFallback = typeof fallbackTitle === 'string' ? fallbackTitle.trim() : '';
    if (lang !== 'zh-CN') return normalizedKey || normalizedFallback || '';
    const title = getTemplateSlideMeta(normalizedKey)?.title;
    return (typeof title === 'string' && title.trim()) || normalizedFallback || normalizedKey || '';
  };

  const formatPlaceholderLabel = (token: string, slideKey?: string | null) => {
    if (lang !== 'zh-CN') return token;
    const targetSlideKey = slideKey || activeSlideKey;
    const slideMeta = getTemplateSlideMeta(targetSlideKey);
    const label = slideMeta?.placeholder_cn_names[token];
    return typeof label === 'string' && label.trim() ? label : token;
  };

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
    return activeAiTokens.length;
  }, [selectedAiMode, selectedAiTokens, activeAiTokens]);

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

  const shouldOpenTimeoutDialog = (errorCode?: string | null, message?: string) => {
    if (errorCode && TIMEOUT_ERROR_CODES.has(String(errorCode))) return true;
    const lower = (message || '').toLowerCase();
    if (!lower) return false;
    return (
      lower.includes('apitimeouterror')
      || lower.includes('retryerror')
      || lower.includes('timeout')
      || lower.includes('timed out')
    );
  };

  type ProgressStage = 'init' | 'parse' | 'retrieve' | 'ai' | 'renderPpt' | 'renderPreview' | 'complete';

  const localizeProgressMessage = (nextProgress?: number, raw?: string): string => {
    const text = (raw || '').trim();
    const lower = text.toLowerCase();
    const progressNumber = typeof nextProgress === 'number' ? nextProgress : 0;
    const has = (keyword: string) => lower.includes(keyword);
    const action = actionRef.current;

    let stage: ProgressStage = 'init';
    if (
      progressNumber >= 100
      || has('complete')
      || has('completed')
      || has('done')
      || has('finished')
      || has('finalized')
    ) {
      stage = 'complete';
    } else if (action === 'rewrite') {
      if (progressNumber >= 78 || ((has('preview') || has('预览')) && (has('render') || has('渲染')))) {
        stage = 'renderPreview';
      } else if (progressNumber >= 66 || has('rendering ppt') || has('渲染 ppt') || has('saving slide specification') || has('finalizing report')) {
        stage = 'renderPpt';
      } else {
        stage = 'parse';
      }
    } else if (action === 'ai-rewrite') {
      if (progressNumber >= 78 || ((has('preview') || has('预览')) && (has('render') || has('渲染')))) {
        stage = 'renderPreview';
      } else if (progressNumber >= 66 || has('rendering ppt') || has('渲染 ppt') || has('saving slide specification') || has('finalizing report')) {
        stage = 'renderPpt';
      } else if (has('rag') || has('retriev') || has('knowledge') || has('知识库') || has('检索') || (progressNumber >= 30 && progressNumber < 48)) {
        stage = 'retrieve';
      } else if (has('ai') || has('generat') || has('rewrite') || has('调用') || has('重写') || (progressNumber >= 48 && progressNumber < 66)) {
        stage = 'ai';
      } else {
        stage = 'parse';
      }
    } else if (progressNumber >= 78 || ((has('preview') || has('预览')) && (has('render') || has('渲染')))) {
      stage = 'renderPreview';
    } else if (progressNumber >= 68 || has('rendering ppt') || has('渲染 ppt') || has('saving slide specification') || has('保存文件') || has('finalizing report')) {
      stage = 'renderPpt';
    } else if (
      has('rag')
      || has('retriev')
      || has('knowledge')
      || has('知识库')
      || has('检索')
    ) {
      stage = 'retrieve';
    } else if (
      ((has('ai') || has('模型')) && (has('generat') || has('rewrite') || has('调用') || has('重写')))
      || has('generate content')
      || has('ai 生成中')
      || has('ai generating')
    ) {
      stage = 'ai';
    } else if (
      progressNumber >= 5
      || has('parse')
      || has('excel')
      || has('input')
      || has('extract')
      || has('data')
      || has('加载模板')
      || has('提取数据')
      || has('占位符')
      || (has('load') && has('template'))
    ) {
      stage = 'parse';
    }

    if (stage === 'complete') return t('progressComplete', lang === 'zh-CN' ? '处理完成' : 'Completed');
    if (stage === 'renderPreview') return t('progressRenderPreview', lang === 'zh-CN' ? '生成预览图片' : 'Generating preview images');
    if (stage === 'renderPpt') return t('progressRenderPpt', lang === 'zh-CN' ? '排版并生成 PPT' : 'Composing and building PPT');
    if (stage === 'ai') return t('progressCallAi', lang === 'zh-CN' ? '生成报告内容' : 'Generating report content');
    if (stage === 'retrieve') return t('progressRetrieveKnowledge', lang === 'zh-CN' ? '匹配相关知识参考' : 'Gathering relevant references');
    if (stage === 'parse') return t('progressParseData', lang === 'zh-CN' ? '读取并整理数据' : 'Reading and structuring data');
    return t('progressInit', lang === 'zh-CN' ? '准备生成任务' : 'Preparing task');
  };

  const applyIncomingProgress = (nextProgress: number | undefined, message?: string) => {
    const numeric = typeof nextProgress === 'number' ? nextProgress : 0;
    if (numeric < progressValueRef.current) return;
    progressValueRef.current = numeric;
    setProgress(numeric);
    setProgressMessage(typeof message === 'string' ? message : '');
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
      const incomingSession = typeof payload.session_id === 'string' ? payload.session_id : '';
      if (activeSessionIdRef.current && incomingSession && incomingSession !== activeSessionIdRef.current) {
        return;
      }
      if (payload.type === 'progress' && typeof payload.progress === 'number') {
        applyIncomingProgress(payload.progress, typeof payload.message === 'string' ? payload.message : undefined);
        return;
      }
      if (payload.type === 'completed' && payload.result && typeof payload.result === 'object') {
        void handleGenerationComplete(payload.result as GenerateResult);
        return;
      }
      if (payload.type === 'failed') {
        const failedResult = isPlainObject(payload.result) ? payload.result : {};
        handleGenerationFailed(
          typeof failedResult.error === 'string' ? failedResult.error : t('errorGenerationFailed'),
          typeof failedResult.error_code === 'string' ? failedResult.error_code : undefined,
        );
      }
    };
  };

  const loadBootstrapData = async () => {
    try {
      const [templatesData, inputsData] = await Promise.all([MainApi.listTemplates(), MainApi.listInputs()]);
      setTemplates(templatesData);
      setInputs(inputsData);
      const firstTemplate = templatesData[0]?.template_id || '';
      const firstInput = getDefaultInputIdForTemplate(firstTemplate, inputsData);
      setSelectedTemplate(firstTemplate);
      setSelectedInput(firstInput);
      if (firstTemplate) void loadTemplateSlides(firstTemplate);
    } catch (error) {
      const message = error instanceof Error ? error.message : t('errorLoadFailed');
      showToast('error', message);
    }
  };

  useEffect(() => {
    activeSessionIdRef.current = activeSessionId;
  }, [activeSessionId]);

  useEffect(() => {
    if (isReloadNavigation()) {
      // Browser refresh should always return user to the configuration screen.
      setInWorkspace(false);
      setLoading(false);
      setGenerationInProgress(false);
      setProgress(0);
      setProgressMessage('');
      setJobId(null);
      setActiveJobId(null);
      setActiveSessionId('');
      setActiveOperationToken('');
      setSlidespec(null);
      clearPreviews();
      setActiveSlideKey(null);
      setModifiedSlides({});
      setExportOpen(false);
      setAiModalOpen(false);
      setTimeoutDialogOpen(false);
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

  const replacePreviews = (nextPreviews: string[]) => {
    setPreviews([...nextPreviews]);
    setPreviewStamp(Date.now());
  };

  const clearPreviews = () => {
    setPreviews([]);
    setPreviewStamp(Date.now());
  };

  const syncPreviewUrls = async (currentJobId: string, forceRegenerate = false) => {
    const previewData = await MainApi.getPreview(currentJobId, forceRegenerate);
    const next = ensureArray(previewData.preview_urls).length > 0 ? ensureArray(previewData.preview_urls) : ensureArray(previewData.images);
    replacePreviews(next);
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
    if (activeJobId && nextJobId !== activeJobId) return;
    if (completionHandledJobRef.current === nextJobId) return;
    completionHandledJobRef.current = nextJobId;
    const shouldToastSuccess = actionRef.current === 'generate';
    setGenerationInProgress(false);
    stopPolling();
    progressValueRef.current = 100;
    setProgress(100);
    setProgressMessage('completed');
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
      replacePreviews(resultPreviews);
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
    setTimeoutDialogOpen(false);
    setTimeoutDialogErrorCode('');
    setTimeoutDialogMessage('');
    actionRef.current = null;
  };

  const handleGenerationFailed = (message: string, errorCode?: string) => {
    if (errorCode === TASK_SUPERSEDED_ERROR_CODE) {
      return;
    }
    setGenerationInProgress(false);
    stopPolling();
    progressValueRef.current = 0;
    setProgress(0);
    setProgressMessage('');
    actionRef.current = null;
    const timeoutDetected = shouldOpenTimeoutDialog(errorCode, message);
    const finalMessage = timeoutDetected
      ? t(
        'msgTimeoutRegenerateHint',
        lang === 'zh-CN'
          ? 'AI 服务长时间未返回结果，已停止当前任务。你可以重新生成一个新任务再试一次。'
          : 'The AI service took too long and the current task was stopped. You can start a brand-new task and try again.',
      )
      : (message || lt('errorGenerationFailed', '报告生成失败', 'Report generation failed'));
    showToast('error', finalMessage);
    if (timeoutDetected) {
      setTimeoutDialogErrorCode(errorCode || '');
      setTimeoutDialogMessage(finalMessage);
      setTimeoutDialogOpen(true);
    }
  };

  const checkCurrentJobStatus = async (pollToken?: string) => {
    if (!jobId || !generationInProgress) return;
    if (actionRef.current && actionRef.current !== 'generate') return;
    if (pollToken && statusPollTokenRef.current && pollToken !== statusPollTokenRef.current) return;
    try {
      const statusData = (await MainApi.getJobStatus(jobId)) as JobStatusResp & { slidespec_path?: string };
      if (pollToken && statusPollTokenRef.current && pollToken !== statusPollTokenRef.current) return;
      if (activeJobId && statusData.job_id && statusData.job_id !== activeJobId) return;
      if (statusData.status === 'running' || statusData.status === 'pending') {
        applyIncomingProgress(statusData.progress, statusData.message);
        return;
      }
      if (statusData.status === 'failed' || statusData.status === 'cancelled') {
        handleGenerationFailed(statusData.message || statusData.last_error || t('msgRegenerateAfterError'), statusData.error_code);
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
    const pollToken = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    statusPollTokenRef.current = pollToken;
    pollRef.current = window.setInterval(() => {
      void checkCurrentJobStatus(pollToken);
    }, 2000);
    return () => {
      stopPolling();
    };
  }, [generationInProgress, jobId, activeJobId]);

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

  const beginGenerate = async (options?: { forceNewTask?: boolean }) => {
    const forceNewTask = !!options?.forceNewTask;
    if (!selectedInput) return showToast('warning', lt('msgSelectInput', '请选择输入数据', 'Please select input data'));
    if (!selectedTemplate) return showToast('warning', lt('msgSelectTemplate', '请选择报告模板', 'Please select report template'));
    const focus = normalizeFocus(selectedFocus);
    if (focus.length === 0) return showToast('warning', lt('msgSelectAtLeastOneFocus', '请至少选择一个关注焦点', 'Please select at least one focus area'));
    const focusTitles = toPreferenceTitles(focus);
    if (generationInProgress) return showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
    if (!inWorkspace) setInWorkspace(true);
    const sid = forceNewTask ? toSessionId() : (sessionId || toSessionId());
    const idempotencyKey = forceNewTask ? toSessionId() : sid;
    const operationToken = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    setSessionId(sid);
    setActiveSessionId(sid);
    setActiveOperationToken(operationToken);
    statusPollTokenRef.current = operationToken;
    setTimeoutDialogOpen(false);
    setTimeoutDialogErrorCode('');
    setTimeoutDialogMessage('');
    setLoading(true);
    setGenerationInProgress(true);
    actionRef.current = 'generate';
    completionHandledJobRef.current = null;
    progressValueRef.current = 0;
    setProgress(0);
    setProgressMessage('');
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
        idempotency_key: idempotencyKey,
        focus_options: focusTitles,
        force_new_task: forceNewTask,
      };
      const result = await MainApi.createReport(payload);
      setJobId(result.job_id || null);
      setActiveJobId(result.job_id || null);
      if (result.session_id) {
        setSessionId(result.session_id);
        setActiveSessionId(result.session_id);
      }
      if (result.status === 'completed' && result.from_cache) {
        await handleGenerationComplete(result as CreateReportResp & GenerateResult);
        return;
      }
      applyIncomingProgress(result.progress, result.message);
      clearPreviews();
      setSlidespec(null);
      setActiveSlideKey(null);
      if (result.existing_job) {
        showToast('info', lt('msgReusingRunningTask', '检测到你已有正在生成的任务，已继续该任务进度。', 'An existing running task was found in this browser. Continuing that task.'));
      } else {
        showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
      }
      if (!wsReady) void checkCurrentJobStatus(operationToken);
    } catch (error) {
      const message = error instanceof HttpError ? error.message : (error as Error).message;
      handleGenerationFailed(message || t('errorGenerationFailed'));
    } finally {
      setLoading(false);
    }
  };

  const updateActiveField = (field: string, value: string) => {
    if (!activeSlideKey) return;
    const originalSlide = slidespec?.slides.find((slide) => slide.slide_key === activeSlideKey);
    if (!originalSlide) return;
    const originalValue = stringifyValue(originalSlide.placeholders?.[field]);
    setModifiedSlides((prev) => {
      const prevSlide = prev[activeSlideKey] || {};
      const nextSlide: StringFieldMap = { ...prevSlide };
      if (value === originalValue) {
        delete nextSlide[field];
      } else {
        nextSlide[field] = value;
      }
      const next = { ...prev };
      if (Object.keys(nextSlide).length === 0) {
        delete next[activeSlideKey];
      } else {
        next[activeSlideKey] = nextSlide;
      }
      return {
        ...next,
      };
    });
  };

  const regenerateAsNewTask = async () => {
    setTimeoutDialogOpen(false);
    await beginGenerate({ forceNewTask: true });
  };

  const openExcelFilePicker = () => {
    if (uploadingExcel || loading || generationInProgress) return;
    excelUploadInputRef.current?.click();
  };

  const uploadExcelFile = async (file: File | null | undefined) => {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      showToast('warning', lt('msgUploadExcelOnlyXlsx', '仅支持上传 .xlsx 文件', 'Only .xlsx files are supported'));
      return;
    }
    if (uploadingExcel || loading || generationInProgress) return;
    setUploadingExcel(true);
    try {
      const result = await MainApi.uploadExcel(file, selectedTemplate);
      const sid = String(result.session_id || '').trim();
      if (!sid) throw new Error(lt('msgUploadMissingSessionId', '上传成功但未返回 session_id', 'Upload succeeded but session_id is missing'));
      setSessionId(sid);
      setSelectedInput('custom');
      showToast('success', lt('msgUploadExcelSuccess', 'Excel 上传成功，已切换为“custom”数据源', 'Excel uploaded. Switched data source to "custom".'));
    } catch (error) {
      const message = error instanceof Error ? error.message : lt('msgUploadExcelFailed', 'Excel 上传失败', 'Excel upload failed');
      showToast('error', message);
    } finally {
      setUploadingExcel(false);
    }
  };

  const handleExcelFileSelected = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    await uploadExcelFile(file);
  };

  const handleExcelDragEnter = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    if (uploadingExcel || loading || generationInProgress) return;
    excelDragDepthRef.current += 1;
    setExcelDragActive(true);
  };

  const handleExcelDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    if (uploadingExcel || loading || generationInProgress) return;
    if (!excelDragActive) setExcelDragActive(true);
  };

  const handleExcelDragLeave = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    excelDragDepthRef.current = Math.max(0, excelDragDepthRef.current - 1);
    if (excelDragDepthRef.current === 0) {
      setExcelDragActive(false);
    }
  };

  const handleExcelDrop = async (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    excelDragDepthRef.current = 0;
    setExcelDragActive(false);
    const file = event.dataTransfer.files?.[0];
    await uploadExcelFile(file);
  };

  const handleReconfigure = () => {
    setInWorkspace(false);
    // Clear identifiers so next generate always creates a new job.
    setSessionId('');
    setActiveSessionId('');
    setActiveJobId(null);
    setActiveOperationToken('');
    setJobId(null);
    setTimeoutDialogOpen(false);
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
    // Keep current slide focus during regeneration to avoid jumping back to page 1.
    setGenerationInProgress(true);
    actionRef.current = 'rewrite';
    progressValueRef.current = 0;
    setProgress(0);
    setProgressMessage('');
    setLoading(true);
    try {
      const result = await MainApi.rewriteSlides(jobId, payload, clientId);
      if (result.slidespec) setSlidespec(result.slidespec);
      setModifiedSlides({});
      const returnedPreviews = ensureArray(result.preview_urls);
      if (returnedPreviews.length > 0) replacePreviews(returnedPreviews);
      else await syncPreviewUrls(jobId, false);
      progressValueRef.current = 100;
      setProgress(100);
      setProgressMessage('completed');
      showToast('success', t('msgUpdateSlidesSuccess').replace('{count}', String(result.updated_count || 0)));
    } catch (error) {
      showToast('error', error instanceof Error ? error.message : t('titleUpdateFailed'));
      progressValueRef.current = 0;
      setProgress(0);
      setProgressMessage('');
    } finally {
      setGenerationInProgress(false);
      setLoading(false);
      actionRef.current = null;
    }
  };

  const openAiRewrite = () => {
    if (!activeSlideKey) return showToast('warning', t('msgSelectSlideForAiRewrite'));
    if (activeAiTokens.length === 0) {
      return showToast(
        'warning',
        lt(
          'msgNoAiGeneratedPlaceholdersToRewrite',
          '当前页没有可用于 AI 重写的占位字段。',
          'This slide has no AI-generated placeholders to rewrite.',
        ),
      );
    }
    setExportOpen(false);
    setAiModalOpen(true);
  };

  const submitAiRewrite = async () => {
    if (!jobId || !activeSlideKey) return;
    const normalizedPrompt = aiPrompt.trim();
    if (!normalizedPrompt) return showToast('warning', t('msgEnterAiRewritePrompt'));
    if (selectedAiMode === 'all' && activeAiTokens.length === 0) {
      return showToast(
        'warning',
        lt(
          'msgNoAiGeneratedPlaceholdersToRewrite',
          '当前页没有可用于 AI 重写的占位字段。',
          'This slide has no AI-generated placeholders to rewrite.',
        ),
      );
    }
    if (selectedAiMode === 'selected' && selectedAiTokens.length === 0) {
      return showToast('warning', t('msgSelectAtLeastOneAiToken'));
    }
    setAiModalOpen(false);
    setGenerationInProgress(true);
    actionRef.current = 'ai-rewrite';
    progressValueRef.current = 0;
    setProgress(0);
    setProgressMessage('');
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
      const returnedPreviews = ensureArray(result.preview_urls);
      if (returnedPreviews.length > 0) replacePreviews(returnedPreviews);
      else await syncPreviewUrls(jobId, false);
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
      setProgressMessage('completed');
      if ((result.updated_count || 0) > 0) showToast('success', t('msgAiRewriteSuccess').replace('{slide}', activeSlideKey));
      else showToast('warning', t('msgAiRewriteNoChanges'));
    } catch (error) {
      showToast('error', error instanceof Error ? error.message : t('titleUpdateFailed'));
      progressValueRef.current = 0;
      setProgress(0);
      setProgressMessage('');
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

  const templateCards = templates;
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

  const selectTemplate = (templateId: string) => {
    setSelectedTemplate(templateId);
    if (selectedInput !== 'custom') {
      setSelectedInput(getDefaultInputIdForTemplate(templateId, inputs));
    }
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

  const requestFullscreenBeforePresenting = async () => {
    if (document.fullscreenElement) return true;
    try {
      const host = document.documentElement as HTMLElement | null;
      await host?.requestFullscreen?.();
      return true;
    } catch {
      return false;
    }
  };

  const openPresentationAtIndex = async (index: number) => {
    if (generationInProgress || loading) return showToast('info', lt('msgGeneratingPleaseWait', '正在生成报告，请稍候...', 'Generating report, please wait...'));
    if (!slidespec || slidespec.slides.length === 0) return showToast('warning', t('msgNoSlidesForPresentation'));
    const targetIndex = Math.max(0, Math.min(slidespec.slides.length - 1, index));
    setPresentIndex(targetIndex);
    await requestFullscreenBeforePresenting();
    setPresentOpen(true);
  };

  const openPresentation = () => {
    void openPresentationAtIndex(activeSlideIndex >= 0 ? activeSlideIndex : 0);
  };

  const openPresentationAtSlide = (slideKey: string) => {
    const targetIndex = slidespec?.slides.findIndex((slide) => slide.slide_key === slideKey) ?? 0;
    void openPresentationAtIndex(targetIndex >= 0 ? targetIndex : 0);
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
  const activeSlideTitle = activeSlide ? formatSlideLabel(activeSlide.slide_key, activeSlide.title) : '';
  const totalSlideCount = slidespec?.slides.length || 0;
  const currentSlideNumber = activeSlideIndex >= 0 ? activeSlideIndex + 1 : 0;
  const modifiedSlideSummaries = useMemo<ModifiedSlideSummary[]>(() => {
    if (!slidespec) return [];
    return slidespec.slides
      .map((slide, idx) => {
        const edited = modifiedSlides[slide.slide_key];
        const fieldCount = edited ? Object.keys(edited).length : 0;
        if (fieldCount === 0) return null;
        return {
          slideKey: slide.slide_key,
          slideNumber: idx + 1,
          title: formatSlideLabel(slide.slide_key, slide.title),
          fieldCount,
        };
      })
      .filter((item): item is ModifiedSlideSummary => item !== null);
  }, [slidespec, modifiedSlides, lang, selectedTemplate, templateSlidesById]);
  const modifiedSlideCount = modifiedSlideSummaries.length;
  const applyButtonLabel = modifiedSlideCount > 0
    ? t(
      'btnApplyRegenerateWithCount',
      lang === 'zh-CN' ? '应用并重新生成（已修改 {count} 页）' : 'Apply & Regenerate ({count} modified pages)',
    ).replace('{count}', String(modifiedSlideCount))
    : t('btnApplyRegenerate');
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
  const aiSlideLabel = activeSlide ? `${currentSlideNumber}/${totalSlideCount} · ${formatSlideLabel(activeSlide.slide_key, activeSlide.title)}` : '-';
  const reconfigureTooltip = generationInProgress
    ? lt('tooltipReconfigureAfterPreview', '预览生成中，完成后可重新配置', 'Preview is generating. Reconfigure after it completes.')
    : t('btnReconfigure');
  const aiRewriteTooltip = activeAiTokens.length === 0
    ? lt('tooltipAiRewriteUnavailableNoToken', 'AI重写不可用：当前页无AI生成内容', 'AI rewrite unavailable: no AI-generated content on this slide')
    : t('btnAiRewriteCurrentSlide', 'AI 重写当前页');
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
            <button
              className="lang-btn tooltip-card-btn"
              onClick={toggleLang}
              data-tooltip={lang === 'zh-CN' ? t('tooltipSwitchToEnglish', '切换到英文') : t('tooltipSwitchToChinese', '切换到中文')}
              aria-label={lang === 'zh-CN' ? t('tooltipSwitchToEnglish', '切换到英文') : t('tooltipSwitchToChinese', '切换到中文')}
            >
              <Globe size={16} /> {lang === 'zh-CN' ? 'EN' : '中文'}
            </button>
          </header>

          <div className="preconfig-grid">
            <section className="preconfig-section preconfig-data-section">
                <h2>
                  <FileSpreadsheet size={16} className="section-icon" /> 1. {t('preConfigSectionData', 'Data Source')}
                </h2>
                <div
                  className={`data-source-card ${excelDragActive ? 'drag-active' : ''}`}
                  onDragEnter={handleExcelDragEnter}
                  onDragOver={handleExcelDragOver}
                  onDragLeave={handleExcelDragLeave}
                  onDrop={handleExcelDrop}
                >
                  <button
                    type="button"
                    className="data-source-cover-btn"
                    onClick={openExcelFilePicker}
                    disabled={uploadingExcel || loading || generationInProgress}
                  >
                    <div className="data-source-icon">
                      <UploadCloud size={30} />
                    </div>
                    <h3>{t('preConfigDataCardTitle', lang === 'zh-CN' ? '本地目录数据' : 'Local directory data')}</h3>
                    <p>
                      {t(
                        'preConfigDataCardDesc',
                        lang === 'zh-CN'
                          ? '默认使用本地目录数据，上传 Excel 后将自动替换为用户数据。'
                          : 'Uses local directory data by default. Uploading Excel will replace it with your data.',
                      )}
                    </p>
                  </button>
                  <div className="data-source-footer">
                    <label className="check-row">
                      <input type="checkbox" checked={useMock} onChange={(e) => setUseMock(e.target.checked)} />
                      <span>{t('labelUseMock')}</span>
                    </label>
                  </div>
                  <input
                    ref={excelUploadInputRef}
                    type="file"
                    accept=".xlsx"
                    style={{ display: 'none' }}
                    onChange={handleExcelFileSelected}
                  />
                </div>
            </section>

            <div className="preconfig-right-col">
              <section className="preconfig-section preconfig-template-section">
                <h2>
                  <LayoutTemplate size={16} className="section-icon" /> 3. {t('preConfigSectionTemplate', 'Select Template')}
                </h2>
                <div
                  className="template-split"
                  onMouseLeave={() => setHoverTemplate('')}
                >
                  <div className="template-list-panel">
                    <div className="template-list-header">
                      <span>{t('preConfigTemplateListTitle', lang === 'zh-CN' ? '模板列表' : 'Template list')}</span>
                    </div>
                    <div className="template-list">
                      {templateCards.map((tpl) => {
                        const selected = selectedTemplate === tpl.template_id;
                        const hovered = hoverTemplate === tpl.template_id;
                        const templateName = toTemplateName(tpl);
                        return (
                          <button
                            key={tpl.template_id}
                            className={`template-list-item ${selected ? 'selected' : ''} ${hovered && !selected ? 'hovered' : ''}`}
                            onClick={() => selectTemplate(tpl.template_id)}
                            onMouseEnter={() => {
                              setHoverTemplate(tpl.template_id);
                              void loadTemplateSlides(tpl.template_id);
                            }}
                          >
                            <div className="template-list-main">
                              <span className="template-list-name">{templateName}</span>
                            </div>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                  <div className="template-gallery-panel">
                    <div className="template-gallery single">
                      {(() => {
                        const activeTemplateId = hoverTemplate || selectedTemplate;
                        const tpl = templateCards.find((item) => item.template_id === activeTemplateId)
                          || templateCards.find((item) => item.template_id === selectedTemplate)
                          || templateCards[0];
                        if (!tpl) return null;
                        const selected = selectedTemplate === tpl.template_id;
                        const hovered = hoverTemplate === tpl.template_id;
                        const frames = templateSlidesById[tpl.template_id] || [];
                        const catalogSlidesCount = typeof tpl.slides_count === 'number' ? tpl.slides_count : Number(tpl.slides_count) || 0;
                        const total = Math.max(frames.length, catalogSlidesCount, 1);
                        const frameIndex = Math.min(templateFrameIndexById[tpl.template_id] || 0, total - 1);
                        const imageUrl = getTemplatePreviewImageUrl(tpl.template_id, frameIndex);
                        const templateName = toTemplateName(tpl);
                        const slidesCountText = t('preConfigSlidesCount', '{count} slides').replace('{count}', String(total));
                        return (
                          <button
                            key={tpl.template_id}
                            className={`template-card ${selected ? 'selected' : ''} ${hovered && !selected ? 'hovered' : ''}`}
                            onClick={() => selectTemplate(tpl.template_id)}
                          >
                            <div
                              className="template-preview"
                              onPointerEnter={() => {
                                setHoverTemplate(tpl.template_id);
                                void loadTemplateSlides(tpl.template_id);
                              }}
                              onPointerMove={(event) => scrubTemplateFrame(tpl.template_id, event.clientX, event.currentTarget)}
                            >
                              {imageUrl ? (
                                <img
                                  src={`${imageUrl}?v=${templatePreviewStamp}&p=${frameIndex + 1}`}
                                  alt={t('preConfigTemplateImageAlt', 'Template page {page} thumbnail').replace('{page}', String(frameIndex + 1))}
                                />
                              ) : (
                                <img src="/placeholders/template-1.svg" alt={toTemplateName(tpl)} />
                              )}
                              <div className="template-preview-meta">
                                <div className="template-meta-left">
                                  <span className="template-fixed-name">{templateName}</span>
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
                      })()}
                    </div>
                  </div>
                </div>
              </section>
            </div>

            <section className="preconfig-section preconfig-focus-section">
              <h2>
                <Settings2 size={16} className="section-icon" /> 2. {t('preConfigFocusTitle', 'Report Focus')}
                <button
                  type="button"
                  className="section-help tooltip-card-btn"
                  data-tooltip={t(
                    'preConfigFocusHelpText',
                    lang === 'zh-CN'
                      ? '请至少选择 1 项关注焦点（支持多选），AI 将围绕你选择的重点生成更高质量内容。'
                      : 'Select at least one focus area (multiple selections supported). AI will prioritize your choices to generate higher-quality content.',
                  )}
                  aria-label={t(
                    'preConfigFocusHelpText',
                    lang === 'zh-CN'
                      ? '请至少选择 1 项关注焦点（支持多选），AI 将围绕你选择的重点生成更高质量内容。'
                      : 'Select at least one focus area (multiple selections supported). AI will prioritize your choices to generate higher-quality content.',
                  )}
                >
                  ?
                </button>
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
                className="workspace-icon-btn tooltip-card-btn tooltip-left-align"
                onClick={handleReconfigure}
                disabled={loading || generationInProgress}
                data-tooltip={reconfigureTooltip}
                aria-label={reconfigureTooltip}
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
                  disabled={!jobId || !activeSlideKey || loading || activeAiTokens.length === 0}
                  data-tooltip={aiRewriteTooltip}
                  aria-label={aiRewriteTooltip}
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
                              <code>{formatPlaceholderLabel(token, activeSlideKey)}</code>
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
                      <button className="btn-primary" onClick={() => void performExport()}>
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
                          src={`${previews[idx]}?v=${previewStamp}`}
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
                      <div className="thumb-title">{formatSlideLabel(slide.slide_key, slide.title)}</div>
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
                          src={`${previews[idx]}?v=${previewStamp}`}
                          alt={slide.title || slide.slide_key}
                          onError={() => markPreviewImgError('main', idx, previews[idx])}
                        />
                      ) : (
                        <div className="preview-placeholder">{t('msgGeneratingPreview')}</div>
                      )}
                    {slide.slide_key === activeSlideKey && activeAiTokens.length > 0 && (
                      <button
                        className="float-ai"
                        onClick={(e) => {
                          e.stopPropagation();
                          openAiRewrite();
                        }}
                      >
                        <Sparkles size={12} /> {t('btnAiRewriteShort', lang === 'zh-CN' ? 'AI重写' : 'AI Rewrite')}
                      </button>
                    )}
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
              <div className="editor-modified-summary">
                <div className="editor-modified-head">
                  <strong>{t('editorModifiedSummaryTitle', lang === 'zh-CN' ? '本次编辑改动' : 'Current edits')}</strong>
                  <span>
                    {t(
                      'editorModifiedSummaryCount',
                      lang === 'zh-CN' ? '已修改 {count} 页' : '{count} pages modified',
                    ).replace('{count}', String(modifiedSlideCount))}
                  </span>
                </div>
                {modifiedSlideCount > 0 ? (
                  <div className="editor-modified-list">
                    {modifiedSlideSummaries.map((item) => (
                      <button
                        type="button"
                        key={item.slideKey}
                        className={`editor-modified-chip ${item.slideKey === activeSlideKey ? 'active' : ''}`}
                        onClick={() => focusSlidePreview(item.slideKey, 'smooth')}
                      >
                        <span className="editor-modified-chip-page">P{item.slideNumber}</span>
                        <span className="editor-modified-chip-title">{item.title}</span>
                        <span className="editor-modified-chip-count">
                          {t(
                            'editorModifiedFieldCount',
                            lang === 'zh-CN' ? '{count} 项' : '{count} fields',
                          ).replace('{count}', String(item.fieldCount))}
                        </span>
                      </button>
                    ))}
                  </div>
                ) : (
                  <div className="editor-modified-empty">
                    {t('editorModifiedSummaryEmpty', lang === 'zh-CN' ? '当前还没有页面被修改' : 'No edited pages yet')}
                  </div>
                )}
              </div>
              <div className="editor-fields">
                {activeSlide ? (
                  activeFieldKeys.map((field) => {
                    const value = activeEditorFields[field] ?? '';
                    const useTextArea = value.length > 80 || value.includes('\n') || value.startsWith('{') || value.startsWith('[');
                    const restoreKey = `${activeSlide.slide_key}:${field}`;
                    return (
                      <label className={`field-card ${restoredFieldMarks[restoreKey] ? 'restored' : ''}`} key={field}>
                        <span>{formatPlaceholderLabel(field, activeSlide?.slide_key)}</span>
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
                  {applyButtonLabel}
                </button>
              </div>
            </aside>
          </div>
        </section>
      )}

      {timeoutDialogOpen && (
        <div className="timeout-modal-overlay">
          <div className="timeout-modal" role="dialog" aria-modal="true" aria-label={t('msgRegenerateAfterError')}>
            <div className="timeout-modal-head">
              <h4>{t('msgRequestTimeout', lang === 'zh-CN' ? '本次生成超时' : 'This generation timed out')}</h4>
              <p>
                {timeoutDialogMessage
                  || t(
                    'msgTimeoutRegenerateHint',
                    lang === 'zh-CN'
                      ? 'AI 服务长时间未返回结果，已停止当前任务。你可以重新生成一个新任务再试一次。'
                      : 'The AI service took too long and the current task was stopped. You can start a brand-new task and try again.',
                  )}
              </p>
              {timeoutDialogErrorCode ? <small>{timeoutDialogErrorCode}</small> : null}
            </div>
            <div className="timeout-modal-actions">
              <button className="btn-secondary" onClick={() => setTimeoutDialogOpen(false)}>
                {t('btnCancel')}
              </button>
              <button className="btn-primary" onClick={() => void regenerateAsNewTask()}>
                {t('btnRegenerateNow', lang === 'zh-CN' ? '立即重新生成' : 'Regenerate now')}
              </button>
            </div>
          </div>
        </div>
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
        <div className="presenter" onClick={closePresentation} onWheel={onPresenterWheel}>
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
              <img src={`${currentPresentUrl}?v=${previewStamp}`} alt="slide" onError={() => markPreviewImgError('present', presentIndex, currentPresentUrl)} />
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
          <div className="progress-float-stage">{localizeProgressMessage(progress, progressMessage)}</div>
          <div className="progress-float-rail">
            <span style={{ width: `${Math.max(0, Math.min(100, progress))}%` }} />
          </div>
          <div className="progress-float-value">{Math.round(progress)}%</div>
        </div>
      )}

    </div>
  );
}
