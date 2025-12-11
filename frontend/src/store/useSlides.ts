import { create } from "zustand";
import { SlideSpec, SlideSpecItem } from "../api/types";

interface State {
  jobId: string | null;
  templateId: string | null;
  slidespec: SlideSpec | null;
  loading: boolean;
  error: string | null;
  previews: string[];
  setSlidespec: (spec: SlideSpec, jobId: string) => void;
  updateSlideData: (slideKey: string, data: Record<string, unknown>) => void;
  setLoading: (val: boolean) => void;
  setError: (msg: string | null) => void;
  setPreviews: (urls: string[]) => void;
}

export const useSlides = create<State>((set) => ({
  jobId: null,
  templateId: null,
  slidespec: null,
  loading: false,
  error: null,
  previews: [],
  setSlidespec: (spec, jobId) =>
    set({ slidespec: spec, jobId, templateId: spec.template_id }),
  updateSlideData: (slideKey, data) =>
    set((state) => {
      if (!state.slidespec) return state;
      const slides: SlideSpecItem[] = state.slidespec.slides.map((s) =>
        s.slide_key === slideKey ? { ...s, data: { ...s.data, ...data } } : s
      );
      return { slidespec: { ...state.slidespec, slides } };
    }),
  setLoading: (val) => set({ loading: val }),
  setError: (msg) => set({ error: msg }),
  setPreviews: (urls) => set({ previews: urls }),
}));
