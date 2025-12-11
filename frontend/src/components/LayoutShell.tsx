import { ReactNode } from "react";

export function LayoutShell({ children }: { children: ReactNode }) {
  return (
    <div className="min-h-screen bg-slate-50">
      <header className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
          <div className="flex items-center gap-2">
            <div className="h-8 w-8 rounded-lg bg-brand-500 text-white flex items-center justify-center font-bold">
              AI
            </div>
            <div>
              <div className="text-sm font-medium text-slate-500">MSS AI PPT</div>
              <div className="text-base font-semibold text-slate-900">
                报告生成与编辑实验台
              </div>
            </div>
          </div>
          <div className="text-xs text-slate-500">
            后端：<code className="bg-slate-100 px-2 py-1 rounded">127.0.0.1:8000</code>
          </div>
        </div>
      </header>
      <main className="mx-auto max-w-6xl px-6 py-4">{children}</main>
    </div>
  );
}
