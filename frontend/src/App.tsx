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
    throw new Error("JSON Parse Error");
  }
}

export default function App() {
  const [templates, setTemplates] = useState<TemplateCatalogItem[]>([]);
  const [inputs, setInputs] = useState<InputCatalogItem[]>([]);
  const [selectedInput, setSelectedInput] = useState<string>("");
  const [selectedTemplate, setSelectedTemplate] = useState<string>("");
  const [logs, setLogs] = useState<string[]>([]);
  const [showLogs, setShowLogs] = useState(false);

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
    const idx = slide.slide_no - 1;
    return previews[idx];
  }, [activeSlide, slides, slidespec, previews]);

  return (
    <LayoutShell>
      <div className="flex flex-col h-full">
        <Toolbar
          inputs={inputs}
          templates={templates}
          selectedInput={selectedInput}
          selectedTemplate={selectedTemplate}
          onGenerate={doGenerate}
          loading={loading}
        />
        
        {error && (
          <div className="bg-red-50 border-b border-red-200 px-4 py-2 flex items-center justify-between text-xs text-red-800">
            <span>Error: {error}</span>
            <button onClick={() => setError(null)} className="font-bold hover:text-red-900">&times;</button>
          </div>
        )}

        <div className="flex-1 flex overflow-hidden">
          {/* Left Sidebar: Slide List & Logs */}
          <div className="w-64 flex-none border-r border-slate-200 bg-white flex flex-col z-10">
            <div className="flex-1 overflow-hidden">
               <SlideList
                 slides={slides}
                 active={activeSlideKey}
                 onSelect={setActiveSlideKey}
               />
            </div>
            
            {/* Logs Toggle/Preview in Sidebar */}
            <div className="flex-none border-t border-slate-200 bg-slate-50">
              <button 
                onClick={() => setShowLogs(!showLogs)}
                className="w-full flex items-center justify-between px-4 py-3 text-xs font-medium text-slate-600 hover:bg-slate-100 transition-colors"
              >
                <div className="flex items-center gap-2">
                  <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
                  </svg>
                  <span>System Logs</span>
                </div>
                <svg className={`w-3 h-3 transition-transform ${showLogs ? "rotate-180" : ""}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 15l7-7 7 7" />
                </svg>
              </button>
              
              {showLogs && (
                 <div className="h-48 overflow-y-auto border-t border-slate-200 bg-slate-900 p-3">
                   <div className="space-y-1">
                     {logs.length === 0 && <div className="text-slate-500 text-[10px] font-mono">No logs available</div>}
                     {logs.map((line, idx) => (
                       <div key={idx} className="font-mono text-[10px] text-emerald-400 break-all leading-tight">
                         <span className="opacity-50 mr-2">{idx+1}</span>{line}
                       </div>
                     ))}
                   </div>
                 </div>
              )}
            </div>
          </div>

          {/* Center: Preview */}
          <div className="flex-1 min-w-0 bg-slate-100 flex flex-col">
            <SlidePreview slide={activeSlide} imageUrl={activePreviewUrl} />
          </div>

          {/* Right Sidebar: Editor */}
          <div className="w-80 flex-none border-l border-slate-200 bg-white flex flex-col z-10">
            <SlideEditor slide={activeSlide} onSubmit={doRewrite} />
          </div>
        </div>
      </div>
    </LayoutShell>
  );
}
