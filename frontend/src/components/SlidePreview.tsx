import { SlideSpecItem } from "../api/types";

interface Props {
  slide?: SlideSpecItem | null;
  imageUrl?: string;
}

export function SlidePreview({ slide, imageUrl }: Props) {
  if (!slide) {
    return (
      <div className="h-full flex flex-col items-center justify-center bg-slate-100/50 text-slate-400">
        <div className="w-24 h-24 bg-slate-200 rounded-full flex items-center justify-center mb-4">
          <svg className="w-10 h-10 text-slate-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
          </svg>
        </div>
        <p className="text-sm font-medium">Select a slide to preview</p>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col bg-slate-100 relative">
      {/* Canvas Toolbar / Header */}
      <div className="flex-none h-12 bg-white border-b border-slate-200 px-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <span className="text-sm font-semibold text-slate-700">{slide.slide_key}</span>
          <span className="text-xs text-slate-400 px-2 py-0.5 bg-slate-100 rounded border border-slate-200">
             #{slide.slide_no}
          </span>
        </div>
        <div className="flex items-center gap-2">
           <div className="flex items-center bg-slate-100 rounded-md p-0.5 border border-slate-200">
             <button className="px-2 py-1 text-xs font-medium bg-white shadow-sm rounded text-slate-700">Preview</button>
             <button className="px-2 py-1 text-xs font-medium text-slate-500 hover:text-slate-700">Code</button>
           </div>
        </div>
      </div>

      {/* Canvas Area */}
      <div className="flex-1 overflow-auto bg-slate-50/50 p-8 flex items-center justify-center relative">
        {/* Background Grid Pattern */}
        <div className="absolute inset-0 opacity-[0.03] pointer-events-none" 
             style={{ backgroundImage: 'radial-gradient(#64748b 1px, transparent 1px)', backgroundSize: '24px 24px' }}>
        </div>

        {imageUrl ? (
          <div className="relative shadow-2xl shadow-slate-400/20 bg-white ring-1 ring-slate-900/5 transition-transform duration-300">
            <img 
              src={imageUrl} 
              alt={slide.slide_key} 
              className="max-w-full max-h-[calc(100vh-250px)] w-auto h-auto object-contain block select-none"
              style={{ minWidth: '400px' }} // prevent collapse
            />
          </div>
        ) : (
          <div className="relative shadow-xl bg-white w-full max-w-4xl aspect-video p-8 overflow-y-auto ring-1 ring-slate-900/5">
             <div className="space-y-4">
                <div className="h-8 w-1/3 bg-slate-100 rounded animate-pulse"></div>
                <div className="space-y-2">
                   <div className="h-4 w-full bg-slate-50 rounded"></div>
                   <div className="h-4 w-5/6 bg-slate-50 rounded"></div>
                   <div className="h-4 w-4/6 bg-slate-50 rounded"></div>
                </div>
                <div className="pt-8">
                  <h3 className="text-sm font-bold text-slate-400 uppercase tracking-wider mb-4">Raw Data</h3>
                  <pre className="text-xs font-mono text-slate-600 whitespace-pre-wrap bg-slate-50 p-4 rounded border border-slate-100">
                    {Object.entries(slide.data || {})
                      .map(([k, v]) => `${k}: ${typeof v === "object" ? JSON.stringify(v) : v}`)
                      .join("\n")}
                  </pre>
                </div>
             </div>
          </div>
        )}
      </div>
      
      {/* Zoom/Status Footer */}
      <div className="flex-none h-8 bg-white border-t border-slate-200 px-4 flex items-center justify-between text-[10px] text-slate-500">
         <div>1920 x 1080px</div>
         <div className="flex items-center gap-2">
           <span>Fit</span>
           <span>100%</span>
         </div>
      </div>
    </div>
  );
}
