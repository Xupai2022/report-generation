import { useState } from "react";
import { SlideSpecItem } from "../api/types";

interface Props {
  slide?: SlideSpecItem | null;
  onSubmit: (jsonText: string) => void;
}

export function SlideEditor({ slide, onSubmit }: Props) {
  const [text, setText] = useState(
    slide ? JSON.stringify(slide.data, null, 2) : "// 在此编辑 JSON 数据"
  );

  if (!slide) {
    return (
      <div className="card h-full min-h-[60vh] p-6 flex items-center justify-center text-slate-400">
        选择幻灯片后编辑内容
      </div>
    );
  }

  return (
    <div className="card h-full min-h-[60vh] p-4 flex flex-col gap-3">
      <div className="flex items-center justify-between">
        <div>
          <div className="text-xs text-slate-500">JSON 编辑</div>
          <div className="text-sm font-semibold text-slate-900">{slide.slide_key}</div>
        </div>
        <button
          onClick={() => onSubmit(text)}
          className="px-3 py-2 rounded-lg bg-brand-600 text-white text-sm hover:bg-brand-700 transition"
        >
          保存并重写
        </button>
      </div>
      <textarea
        value={text}
        onChange={(e) => setText(e.target.value)}
        className="flex-1 w-full rounded-lg border border-slate-200 bg-slate-50 p-3 font-mono text-xs text-slate-800 focus:border-brand-400 focus:outline-none min-h-[40vh]"
      />
    </div>
  );
}
