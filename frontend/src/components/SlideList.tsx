import { SlideSpecItem } from "../api/types";

interface Props {
  slides: SlideSpecItem[];
  active: string | null;
  onSelect: (key: string) => void;
}

export function SlideList({ slides, active, onSelect }: Props) {
  if (slides.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full p-6 text-center text-slate-400">
        <svg className="w-12 h-12 mb-3 text-slate-300" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" />
        </svg>
        <p className="text-sm">No slides generated</p>
      </div>
    );
  }

  return (
    <div className="h-full overflow-y-auto p-3 space-y-2 custom-scrollbar">
      <div className="px-1 py-2 text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center justify-between">
        <span>Slides</span>
        <span className="bg-slate-100 text-slate-500 rounded-full px-2 py-0.5 text-[10px]">{slides.length}</span>
      </div>
      <div className="space-y-2">
        {slides.map((s, idx) => {
          const isActive = active === s.slide_key;
          return (
            <button
              key={s.slide_key}
              onClick={() => onSelect(s.slide_key)}
              className={`group w-full text-left p-3 rounded-lg border transition-all duration-200 relative ${
                isActive
                  ? "border-brand-500 bg-brand-50/50 shadow-sm ring-1 ring-brand-500/20"
                  : "border-slate-200 hover:border-brand-300 hover:shadow-sm bg-white"
              }`}
            >
              <div className="flex items-start gap-3">
                <div className={`
                  flex-none w-6 h-6 rounded flex items-center justify-center text-[10px] font-bold transition-colors
                  ${isActive ? "bg-brand-500 text-white" : "bg-slate-100 text-slate-500 group-hover:bg-brand-100 group-hover:text-brand-600"}
                `}>
                  {idx + 1}
                </div>
                <div className="min-w-0 flex-1">
                  <div className={`text-xs font-medium truncate ${isActive ? "text-brand-900" : "text-slate-700"}`}>
                    {s.slide_key}
                  </div>
                  <div className="text-[10px] text-slate-400 truncate mt-0.5">
                    {s.layout || "Default Layout"}
                  </div>
                </div>
              </div>
              {isActive && (
                <div className="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-8 bg-brand-500 rounded-r-full"></div>
              )}
            </button>
          );
        })}
      </div>
    </div>
  );
}
