import { SlideSpecItem } from "../api/types";

interface Props {
  slide?: SlideSpecItem | null;
  imageUrl?: string;
}

export function SlidePreview({ slide, imageUrl }: Props) {
  if (!slide) {
    return (
      <div className="card h-full min-h-[60vh] p-6 flex items-center justify-center text-slate-400">
        选择左侧幻灯片查看预览
      </div>
    );
  }

  return (
    <div className="card h-full min-h-[60vh] p-6 space-y-3">
      <div className="flex items-center justify-between">
        <div>
          <div className="text-xs text-slate-500">#{slide.slide_no}</div>
          <div className="text-lg font-semibold text-slate-900">{slide.slide_key}</div>
        </div>
        <span className="pill">{imageUrl ? "PPT 预览" : "文本预览"}</span>
      </div>
      {imageUrl ? (
        <div
          className="border border-slate-200 rounded-lg overflow-hidden bg-slate-100 mx-auto"
          style={{ width: "1920px", height: "1080px" }}
        >
          <img src={imageUrl} alt={slide.slide_key} className="w-full h-full object-contain" />
        </div>
      ) : (
        <div
          className="rounded-lg bg-slate-900 text-slate-50 text-sm p-4 h-full overflow-y-auto whitespace-pre-wrap mx-auto"
          style={{ width: "1920px", height: "1080px" }}
        >
          {Object.entries(slide.data || {})
            .map(([k, v]) => `${k}: ${typeof v === "object" ? JSON.stringify(v) : v}`)
            .join("\n")}
        </div>
      )}
    </div>
  );
}
