import { ReactNode } from "react";

export function LayoutShell({ children }: { children: ReactNode }) {
  return (
    <div className="h-full flex flex-col bg-slate-50 text-slate-900 font-sans">
      {/* Header */}
      <header className="flex-none z-10 bg-white/80 backdrop-blur-md border-b border-slate-200">
        <div className="flex h-16 items-center justify-between px-6">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-gradient-to-br from-brand-500 to-accent-600 text-white shadow-lg shadow-brand-500/20">
              <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19.428 15.428a2 2 0 00-1.022-.547l-2.384-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z" />
              </svg>
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight text-slate-900 leading-tight">
                MSS AI PPT
              </h1>
              <p className="text-[10px] font-medium text-slate-500 uppercase tracking-wider">
                Intelligent Report Generator
              </p>
            </div>
          </div>
          
          <div className="flex items-center gap-4">
            <div className="hidden md:flex items-center gap-2 px-3 py-1.5 bg-slate-100 rounded-full border border-slate-200">
              <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></div>
              <span className="text-xs font-medium text-slate-600">System Online</span>
            </div>
            <div className="text-xs text-slate-400 font-mono">
              v0.1.0-beta
            </div>
          </div>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="flex-1 min-h-0 overflow-hidden relative">
        {children}
      </main>
    </div>
  );
}
