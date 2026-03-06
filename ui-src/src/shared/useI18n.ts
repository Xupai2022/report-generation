import { useEffect, useMemo, useState } from 'react';
import type { I18nMap, Lang } from '../types/api';

const FALLBACK_LANG: Lang = 'zh-CN';

export function useI18n() {
  const [lang, setLang] = useState<Lang>(() => {
    const saved = window.localStorage.getItem('ui_lang');
    if (saved === 'zh-CN' || saved === 'en-US') {
      return saved;
    }
    return FALLBACK_LANG;
  });
  const [messages, setMessages] = useState<I18nMap>({});

  useEffect(() => {
    let active = true;
    fetch(`/i18n/${lang}.json`)
      .then((r) => r.json())
      .then((data: I18nMap) => {
        if (active) {
          setMessages(data);
        }
      })
      .catch(() => {
        if (active) {
          setMessages({});
        }
      });

    window.localStorage.setItem('ui_lang', lang);
    return () => {
      active = false;
    };
  }, [lang]);

  const t = useMemo(() => {
    return (key: string, fallback?: string) => messages[key] || fallback || key;
  }, [messages]);

  const tf = useMemo(() => {
    return (key: string, params: Record<string, string | number>) => {
      let text = messages[key] || key;
      Object.entries(params).forEach(([k, v]) => {
        text = text.replace(new RegExp(`\\{${k}\\}`, 'g'), String(v));
      });
      return text;
    };
  }, [messages]);

  return {
    lang,
    setLang,
    toggleLang: () => setLang((prev) => (prev === 'zh-CN' ? 'en-US' : 'zh-CN')),
    t,
    tf,
  };
}
