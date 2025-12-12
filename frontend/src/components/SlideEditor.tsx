import { useState, useEffect } from "react";
import { SlideSpecItem } from "../api/types";

interface Props {
  slide?: SlideSpecItem | null;
  onSubmit: (jsonText: string) => void;
}

export function SlideEditor({ slide, onSubmit }: Props) {
  const [text, setText] = useState("");

  useEffect(() => {
    if (slide) {
      setText(JSON.stringify(slide.data, null, 2));
    } else {
      setText("");
    }
  }, [slide]);

  if (!slide) {
    return (
      <div className="h-full flex items-center justify-center text-slate-400 p-6 text-center">
        <div className="text-xs">Select a slide to edit properties</div>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col bg-white border-l border-slate-200">
      <div className="flex-none px-4 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50/50">
        <h3 className="text-xs font-bold text-slate-700 uppercase tracking-wider">Properties</h3>
        <button
          onClick={() => onSubmit(text)}
          className="px-3 py-1.5 rounded-md bg-brand-600 text-white text-xs font-medium hover:bg-brand-700 transition shadow-sm hover:shadow"
        >
          Apply Changes
        </button>
      </div>
      
      <div className="flex-1 overflow-hidden flex flex-col">
        <div className="flex-1 relative">
          <textarea
            value={text}
            onChange={(e) => setText(e.target.value)}
            spellCheck={false}
            className="absolute inset-0 w-full h-full p-4 font-mono text-xs leading-5 bg-white text-slate-800 resize-none focus:outline-none custom-scrollbar selection:bg-brand-100 selection:text-brand-900"
            style={{ tabSize: 2 }}
          />
        </div>
      </div>

      <div className="flex-none p-3 border-t border-slate-200 bg-slate-50">
        <div className="text-[10px] text-slate-500 flex items-center justify-between">
           <span>JSON Format</span>
           <span className={text.length > 1000 ? "text-amber-600" : "text-slate-400"}>
             {text.length} chars
           </span>
        </div>
      </div>
    </div>
  );
}
