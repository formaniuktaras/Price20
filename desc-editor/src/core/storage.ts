import dayjs from 'dayjs';
import { DescDoc, EditorDoc, EditorState, Lang, toDescDoc } from './state';

const STORAGE_KEY = 'desc-editor-autosave';

export interface PersistedState {
  activeLang: Lang;
  docs: Record<Lang, DescDoc>;
}

const hasStorage = () => typeof window !== 'undefined' && !!window.localStorage;

export const loadFromStorage = (): PersistedState | null => {
  if (!hasStorage()) return null;
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as PersistedState;
    return parsed;
  } catch (err) {
    console.warn('[DescEditor] failed to parse autosave payload', err);
    return null;
  }
};

export const saveToStorage = (state: EditorState) => {
  if (!hasStorage()) return;
  const docs = Object.entries(state.docs).reduce((acc, [lang, doc]) => {
    acc[lang as Lang] = toDescDoc(lang as Lang, doc);
    return acc;
  }, {} as Record<Lang, DescDoc>);
  const payload: PersistedState = {
    activeLang: state.activeLang,
    docs,
  };
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
};

export const clearStorage = () => {
  if (!hasStorage()) return;
  window.localStorage.removeItem(STORAGE_KEY);
};

export interface ExportPayload {
  lang: Lang;
  html: string;
  css: string;
  assets: DescDoc['assets'];
  history: DescDoc['history'];
  exportedAt: string;
}

export const exportState = (doc: DescDoc): Blob => {
  const payload: ExportPayload = {
    lang: doc.lang,
    html: doc.html,
    css: doc.css,
    assets: doc.assets,
    history: doc.history ?? [],
    exportedAt: dayjs().toISOString(),
  };
  return new Blob([JSON.stringify(payload, null, 2)], {
    type: 'application/json',
  });
};

export const parseImport = async (file: File): Promise<DescDoc> => {
  const text = await file.text();
  const parsed = JSON.parse(text) as ExportPayload;
  const doc: DescDoc = {
    lang: parsed.lang,
    html: parsed.html,
    css: parsed.css,
    assets: parsed.assets ?? [],
    history: parsed.history ?? [],
  };
  return doc;
};

export const toBundle = (doc: EditorDoc): string => {
  return `${doc.html}\n<style>${doc.css}</style>`;
};
