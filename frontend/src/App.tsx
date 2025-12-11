import { useEffect, useMemo, useState } from "react";
import { api } from "./api/client";
import { GenerateResponse, InputCatalogItem, TemplateCatalogItem } from "./api/types";
import { LayoutShell } from "./components/LayoutShell";
import { SlideEditor } from "./components/SlideEditor";
import { SlideList } from "./components/SlideList";
import { SlidePreview } from "./components/SlidePreview";
import { Toolbar } from "./components/Toolbar";
import { useSlides } from "./store/useSlides";

function safeParseJson(text: string): Record<string, unknown> {
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error("JSON 解析失败，请检查格式");
  }
}

export default function App() {
  const [templates, setTemplates] = useState<TemplateCatalogItem[]>([]);
  const [inputs, setInputs] = useState<InputCatalogItem[]>([]);
  const [selectedInput, setSelectedInput] = useState<string>("");
  const [selectedTemplate, setSelectedTemplate] = useState<string>("");
  const [logs, setLogs] = useState<string[]>([]);

  const { slidespec, jobId, setSlidespec, loading, setLoading, setError, error } = useSlides();
  const { previews, setPreviews } = useSlides.getState();
  const slides = slidespec?.slides || [];
  const [activeSlideKey, setActiveSlideKey] = useState<string | null>(null);

  useEffect(() => {
    api.listTemplates().then(setTemplates).catch((e) => setError(e.message));
    api.listInputs().then(setInputs).catch((e) => setError(e.message));
    api.getLogs(20).then((res) => setLogs(res.lines)).catch(() => undefined);
  }, [setError]);

  const activeSlide = useMemo(
    () => slides.find((s) => s.slide_key === activeSlideKey) || slides[0],
    [slides, activeSlideKey]
  );

  const fetchPreview = async (job_id: string) => {
    try {
      const res = await api.preview(job_id, true);
      setPreviews(res.images);
    } catch (e: any) {
      // 如果预览失败，仅记录错误，不中断
      setError(e.message);
    }
  };

  const doGenerate = async (inputId: string, templateId: string, useMock = true) => {
    setSelectedInput(inputId);
    setSelectedTemplate(templateId);
    if (!inputId || !templateId) return;
    try {
      setLoading(true);
      const res: GenerateResponse = await api.generate(inputId, templateId, useMock);
      setSlidespec(res.slidespec, res.job_id);
      setActiveSlideKey(res.slidespec.slides[0]?.slide_key || null);
      await fetchPreview(res.job_id);
      const latestLogs = await api.getLogs(20);
      setLogs(latestLogs.lines);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const doRewrite = async (jsonText: string) => {
    if (!jobId || !activeSlide) return;
    try {
      setLoading(true);
      const parsed = safeParseJson(jsonText);
      const res = await api.rewrite(jobId, activeSlide.slide_key, parsed);
      setSlidespec(res.slidespec, res.job_id);
      setActiveSlideKey(res.slide_key);
      await fetchPreview(res.job_id);
      const latestLogs = await api.getLogs(20);
      setLogs(latestLogs.lines);
    } catch (e: any) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  const activePreviewUrl = useMemo(() => {
    if (!slidespec || !previews.length) return undefined;
    const slide = activeSlide || slides[0];
    if (!slide) return undefined;
    // Previews are named slide1.png, slide2.png ...
    const idx = slide.slide_no - 1;
    return previews[idx];
  }, [activeSlide, slides, slidespec, previews]);

  return (
    <LayoutShell>
      <div className="space-y-4">
        <Toolbar
          inputs={inputs}
          templates={templates}
          selectedInput={selectedInput}
          selectedTemplate={selectedTemplate}
          onGenerate={doGenerate}
          loading={loading}
        />
        {error ? (
          <div className="card border-red-200 bg-red-50 text-red-800 px-4 py-3 text-sm">
            {error}
          </div>
        ) : null}
        <div className="grid grid-cols-12 gap-4 items-start">
          <div className="col-span-12 lg:col-span-2">
            <SlideList
              slides={slides}
              active={activeSlideKey}
              onSelect={setActiveSlideKey}
            />
          </div>
          <div className="col-span-12 lg:col-span-8">
            <SlidePreview slide={activeSlide} imageUrl={activePreviewUrl} />
          </div>
          <div className="col-span-12 lg:col-span-2">
            <SlideEditor slide={activeSlide} onSubmit={doRewrite} />
          </div>
        </div>
        <div className="card p-4">
          <div className="flex items-center justify-between">
            <div className="text-sm font-semibold text-slate-800">审计日志（最近）</div>
            <div className="text-xs text-slate-500">Job: {jobId || "未生成"}</div>
          </div>
          <div className="mt-2 text-xs text-slate-600 space-y-1 max-h-48 overflow-y-auto">
            {logs.map((line, idx) => (
              <div key={idx} className="font-mono">
                {line}
              </div>
            ))}
            {!logs.length && <div className="text-slate-400">暂无日志</div>}
          </div>
        </div>
      </div>
    </LayoutShell>
  );
}
