import { SlideSpecItem } from "../api/types";

interface Props {
  slides: SlideSpecItem[];
  active: string | null;
  onSelect: (key: string) => void;
}

export function SlideList({ slides, active, onSelect }: Props) {
  return (
    <div className="card p-3 space-y-2 h-full overflow-y-auto">
      <div className="text-sm font-semibold text-slate-700">幻灯片</div>
      <div className="space-y-1">
        {slides.map((s) => (
          <button
            key={s.slide_key}
            onClick={() => onSelect(s.slide_key)}
            className={`w-full text-left px-3 py-2 rounded-lg border transition ${
              active === s.slide_key
                ? "border-brand-500 bg-brand-50 text-brand-700"
                : "border-slate-200 hover:border-brand-200 hover:bg-slate-50"
            }`}
          >
            <div className="text-xs uppercase text-slate-500">#{s.slide_no}</div>
            <div className="text-sm font-semibold">{s.slide_key}</div>
          </button>
        ))}
      </div>
    </div>
  );
}
