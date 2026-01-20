import React, { useCallback, useContext, useMemo, useReducer } from 'react';

export type Lang = 'uk' | 'ru' | 'en';

export interface AssetItem {
  name: string;
  dataUrl: string;
}

export interface HistoryEntry {
  ts: string;
  html: string;
  css: string;
}

export interface DescDoc {
  lang: Lang;
  html: string;
  css: string;
  assets: AssetItem[];
  history?: HistoryEntry[];
}

export interface EditorDoc {
  html: string;
  css: string;
  assets: AssetItem[];
  history: HistoryEntry[];
}

export interface EditorState {
  activeLang: Lang;
  mode: 'visual' | 'code' | 'preview';
  docs: Record<Lang, EditorDoc>;
}

type EditorAction =
  | { type: 'set-lang'; lang: Lang }
  | { type: 'set-mode'; mode: EditorState['mode'] }
  | { type: 'set-doc'; lang: Lang; doc: Partial<EditorDoc> }
  | { type: 'replace'; state: EditorState }
  | { type: 'push-history'; lang: Lang; entry: HistoryEntry };

const createEmptyDoc = (): EditorDoc => ({
  html: '<h1>{{brand}} {{model}}</h1>',
  css: 'body { font-family: "Inter", sans-serif; color: #1c1c1c; }',
  assets: [],
  history: [],
});

export const createInitialState = (initial?: Partial<Record<Lang, DescDoc>>): EditorState => {
  const base: Record<Lang, EditorDoc> = {
    uk: createEmptyDoc(),
    ru: createEmptyDoc(),
    en: createEmptyDoc(),
  };

  (Object.entries(initial || {}) as Array<[Lang, DescDoc]>).forEach(([lang, value]) => {
    base[lang] = {
      html: value.html,
      css: value.css,
      assets: value.assets ?? [],
      history: value.history ?? [],
    };
  });

  return {
    activeLang: 'uk',
    mode: 'visual',
    docs: base,
  };
};

const EditorContext = React.createContext<EditorContextValue | null>(null);

export const EditorProvider: React.FC<{ initial?: Partial<Record<Lang, DescDoc>> }> = ({ initial, children }) => {
  const [state, dispatch] = useReducer(editorReducer, createInitialState(initial));

  const value = useMemo<EditorContextValue>(() => ({
    state,
    dispatch,
  }), [state]);

  return <EditorContext.Provider value={value}>{children}</EditorContext.Provider>;
};

export const useEditorContext = (): EditorContextValue => {
  const ctx = useContext(EditorContext);
  if (!ctx) {
    throw new Error('useEditorContext must be used within EditorProvider');
  }
  return ctx;
};

const editorReducer = (state: EditorState, action: EditorAction): EditorState => {
  switch (action.type) {
    case 'set-lang':
      return { ...state, activeLang: action.lang };
    case 'set-mode':
      return { ...state, mode: action.mode };
    case 'set-doc': {
      const current = state.docs[action.lang];
      return {
        ...state,
        docs: {
          ...state.docs,
          [action.lang]: {
            ...current,
            ...action.doc,
          },
        },
      };
    }
    case 'push-history': {
      const current = state.docs[action.lang];
      const history = [action.entry, ...current.history].slice(0, 25);
      return {
        ...state,
        docs: {
          ...state.docs,
          [action.lang]: {
            ...current,
            history,
          },
        },
      };
    }
    case 'replace':
      return action.state;
    default:
      return state;
  }
};

export const useActiveDoc = (): [EditorDoc, (doc: Partial<EditorDoc>) => void] => {
  const {
    state: { activeLang, docs },
    dispatch,
  } = useEditorContext();

  const setDoc = useCallback(
    (doc: Partial<EditorDoc>) => {
      dispatch({ type: 'set-doc', lang: activeLang, doc });
    },
    [activeLang, dispatch]
  );

  return [docs[activeLang], setDoc];
};

export const useEditorActions = () => {
  const { state, dispatch } = useEditorContext();

  const setLang = useCallback((lang: Lang) => dispatch({ type: 'set-lang', lang }), [dispatch]);
  const setMode = useCallback((mode: EditorState['mode']) => dispatch({ type: 'set-mode', mode }), [dispatch]);
  const updateHtml = useCallback(
    (lang: Lang, html: string) => dispatch({ type: 'set-doc', lang, doc: { html } }),
    [dispatch]
  );
  const updateCss = useCallback(
    (lang: Lang, css: string) => dispatch({ type: 'set-doc', lang, doc: { css } }),
    [dispatch]
  );
  const setAssets = useCallback(
    (lang: Lang, assets: AssetItem[]) => dispatch({ type: 'set-doc', lang, doc: { assets } }),
    [dispatch]
  );
  const pushHistory = useCallback(
    (lang: Lang, entry: HistoryEntry) => dispatch({ type: 'push-history', lang, entry }),
    [dispatch]
  );
  const replaceState = useCallback((state: EditorState) => dispatch({ type: 'replace', state }), [dispatch]);

  return {
    state,
    setLang,
    setMode,
    updateHtml,
    updateCss,
    setAssets,
    pushHistory,
    replaceState,
  };
};

export interface DescEditorRef {
  getValue: (lang?: Lang) => DescDoc;
  setValue: (doc: DescDoc) => void;
  importJson: (file: File) => Promise<void>;
  exportJson: () => Blob;
  toHtmlBundle: (lang?: Lang) => string;
}

export interface EditorContextValue {
  state: EditorState;
  dispatch: React.Dispatch<EditorAction>;
}

export const toDescDoc = (lang: Lang, doc: EditorDoc): DescDoc => ({
  lang,
  html: doc.html,
  css: doc.css,
  assets: doc.assets,
  history: doc.history,
});

export const withHistorySnapshot = (doc: EditorDoc): EditorDoc => {
  const snapshot: HistoryEntry = {
    ts: new Date().toISOString(),
    html: doc.html,
    css: doc.css,
  };
  return {
    ...doc,
    history: [snapshot, ...doc.history].slice(0, 25),
  };
};

export const buildStateFromDocs = (docs: Record<Lang, DescDoc>, activeLang: Lang): EditorState => {
  const nextDocs = Object.entries(docs).reduce((acc, [lang, doc]) => {
    acc[lang as Lang] = {
      html: doc.html,
      css: doc.css,
      assets: doc.assets ?? [],
      history: doc.history ?? [],
    };
    return acc;
  }, {} as Record<Lang, EditorDoc>);

  return {
    activeLang,
    mode: 'visual',
    docs: nextDocs,
  };
};
