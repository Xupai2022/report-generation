interface Props {
  inputs: { id: string; tenant_name?: string; period?: string }[];
  templates: { template_id: string; name: string }[];
  selectedInput: string | null;
  selectedTemplate: string | null;
  onGenerate: (inputId: string, templateId: string, useMock?: boolean) => void;
  loading: boolean;
}

export function Toolbar({
  inputs,
  templates,
  selectedInput,
  selectedTemplate,
  onGenerate,
  loading,
}: Props) {
  return (
    <div className="bg-white border-b border-slate-200 px-6 py-3 flex items-center justify-between shadow-sm z-10 relative">
      <div className="flex items-center gap-4">
        <div className="flex flex-col">
          <label className="text-[10px] font-semibold uppercase tracking-wider text-slate-500 mb-0.5">Input Data</label>
          <select
            value={selectedInput || ""}
            onChange={(e) => onGenerate(e.target.value, selectedTemplate || "", true)}
            className="block w-64 rounded-md border-0 py-1.5 pl-3 pr-10 text-slate-900 ring-1 ring-inset ring-slate-300 focus:ring-2 focus:ring-brand-600 sm:text-sm sm:leading-6 bg-slate-50 hover:bg-white transition-colors cursor-pointer"
          >
            <option value="">Select Data Source...</option>
            {inputs.map((i) => (
              <option key={i.id} value={i.id}>
                {i.tenant_name || i.id} {i.period ? `(${i.period})` : ""}
              </option>
            ))}
          </select>
        </div>

        <div className="w-px h-8 bg-slate-200 mx-2"></div>

        <div className="flex flex-col">
          <label className="text-[10px] font-semibold uppercase tracking-wider text-slate-500 mb-0.5">Template</label>
          <select
            value={selectedTemplate || ""}
            onChange={(e) => onGenerate(selectedInput || "", e.target.value, true)}
            className="block w-64 rounded-md border-0 py-1.5 pl-3 pr-10 text-slate-900 ring-1 ring-inset ring-slate-300 focus:ring-2 focus:ring-brand-600 sm:text-sm sm:leading-6 bg-slate-50 hover:bg-white transition-colors cursor-pointer"
          >
            <option value="">Select Design Template...</option>
            {templates.map((t) => (
              <option key={t.template_id} value={t.template_id}>
                {t.name}
              </option>
            ))}
          </select>
        </div>
      </div>

      <div className="flex items-center gap-3">
        <button
          onClick={() =>
            selectedInput && selectedTemplate && onGenerate(selectedInput, selectedTemplate, true)
          }
          disabled={!selectedInput || !selectedTemplate || loading}
          className={`
            inline-flex items-center gap-2 rounded-md px-4 py-2 text-sm font-semibold shadow-sm transition-all focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-brand-600
            ${
              loading || !selectedInput || !selectedTemplate
                ? "bg-slate-100 text-slate-400 cursor-not-allowed"
                : "bg-white text-slate-700 ring-1 ring-inset ring-slate-300 hover:bg-slate-50"
            }
          `}
        >
          {loading ? (
             <svg className="animate-spin -ml-1 mr-2 h-4 w-4 text-slate-500" fill="none" viewBox="0 0 24 24">
               <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
               <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
             </svg>
          ) : (
            <svg className="w-4 h-4 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M14.752 11.168l-3.197-2.132A1 1 0 0010 9.87v4.263a1 1 0 001.555.832l3.197-2.132a1 1 0 000-1.664z" />
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
          )}
          Generate (Mock)
        </button>

        <button
          onClick={() =>
            selectedInput && selectedTemplate && onGenerate(selectedInput, selectedTemplate, false)
          }
          disabled={!selectedInput || !selectedTemplate || loading}
          className={`
            inline-flex items-center gap-2 rounded-md px-4 py-2 text-sm font-semibold text-white shadow-sm focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-brand-600 transition-all
            ${
              loading || !selectedInput || !selectedTemplate
                ? "bg-brand-400 cursor-not-allowed opacity-70"
                : "bg-gradient-to-r from-brand-600 to-accent-600 hover:from-brand-500 hover:to-accent-500 shadow-brand-500/25"
            }
          `}
        >
          {loading ? "Processing..." : "Generate Report"}
          <svg className="w-4 h-4 text-white/80" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
          </svg>
        </button>
      </div>
    </div>
  );
}
