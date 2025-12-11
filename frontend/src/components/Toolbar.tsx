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
    <div className="card p-4 flex flex-wrap gap-3 items-center">
      <div className="text-sm font-semibold text-slate-800">生成报告</div>
      <select
        value={selectedInput || ""}
        onChange={(e) => onGenerate(e.target.value, selectedTemplate || "", true)}
        className="rounded-lg border border-slate-200 px-3 py-2 text-sm"
      >
        <option value="">选择输入数据</option>
        {inputs.map((i) => (
          <option key={i.id} value={i.id}>
            {i.id} {i.tenant_name ? `- ${i.tenant_name}` : ""} {i.period ? `(${i.period})` : ""}
          </option>
        ))}
      </select>
      <select
        value={selectedTemplate || ""}
        onChange={(e) => onGenerate(selectedInput || "", e.target.value, true)}
        className="rounded-lg border border-slate-200 px-3 py-2 text-sm"
      >
        <option value="">选择模板</option>
        {templates.map((t) => (
          <option key={t.template_id} value={t.template_id}>
            {t.template_id} - {t.name}
          </option>
        ))}
      </select>
      <button
        onClick={() =>
          selectedInput && selectedTemplate && onGenerate(selectedInput, selectedTemplate, true)
        }
        disabled={!selectedInput || !selectedTemplate || loading}
        className={`px-4 py-2 rounded-lg text-sm font-semibold text-white transition ${
          loading
            ? "bg-slate-400 cursor-not-allowed"
            : "bg-brand-600 hover:bg-brand-700"
        }`}
      >
        {loading ? "生成中..." : "AI生成（mock）"}
      </button>
      <button
        onClick={() =>
          selectedInput && selectedTemplate && onGenerate(selectedInput, selectedTemplate, false)
        }
        disabled={!selectedInput || !selectedTemplate || loading}
        className={`px-4 py-2 rounded-lg text-sm font-semibold text-brand-700 border border-brand-500 transition ${
          loading ? "opacity-60 cursor-not-allowed" : "hover:bg-brand-50"
        }`}
      >
        {loading ? "生成中..." : "生成（后端配置决定是否用 LLM/mock）"}
      </button>
    </div>
  );
}
