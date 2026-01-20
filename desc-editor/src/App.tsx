import React, { useEffect, useImperativeHandle, useRef, useState } from 'react';
import clsx from 'clsx';
import {
  DescDoc,
  DescEditorRef,
  EditorProvider,
  EditorState,
  Lang,
  buildStateFromDocs,
  toDescDoc,
  useEditorActions,
} from './core/state';
import VisualEditor from './editor/VisualEditor';
import CodeEditor from './editor/CodeEditor';
import PreviewPane from './editor/PreviewPane';
import Assets from './editor/Sidebar/Assets';
import Blocks from './editor/Sidebar/Blocks';
import History from './editor/Sidebar/History';
import { clearStorage, exportState, loadFromStorage, parseImport, saveToStorage, toBundle } from './core/storage';
import { isHosted, loadHostState, sendStateToHost } from './core/host';

const LangTabs: React.FC = () => {
  const {
    state: { activeLang },
    setLang,
  } = useEditorActions();

  return (
    <div className="lang-tabs">
      {(['uk', 'ru', 'en'] as const).map((lang) => (
        <button
          key={lang}
          type="button"
          className={clsx({ active: activeLang === lang })}
          onClick={() => setLang(lang)}
        >
          {lang.toUpperCase()}
        </button>
      ))}
    </div>
  );
};

const ModeTabs: React.FC = () => {
  const {
    state: { mode },
    setMode,
  } = useEditorActions();

  return (
    <div className="mode-tabs">
      <button type="button" className={clsx({ active: mode === 'visual' })} onClick={() => setMode('visual')}>
        Visual
      </button>
      <button type="button" className={clsx({ active: mode === 'code' })} onClick={() => setMode('code')}>
        Code
      </button>
      <button type="button" className={clsx({ active: mode === 'preview' })} onClick={() => setMode('preview')}>
        Preview
      </button>
    </div>
  );
};

const Workspace: React.FC = () => {
  const {
    state,
    setMode,
    pushHistory,
    setLang,
    updateHtml,
    updateCss,
    setAssets,
  } = useEditorActions();
  const doc = state.docs[state.activeLang];
  const [fullScreen, setFullScreen] = useState(false);
  const [isSending, setIsSending] = useState(false);
  const importInputRef = useRef<HTMLInputElement>(null);
  const lastSnapshotRef = useRef<Record<string, string>>({});

  useEffect(() => {
    const snapshotKey = `${doc.html}:::${doc.css}`;
    const map = lastSnapshotRef.current;
    if (map[state.activeLang] === snapshotKey) return;
    const timeout = setTimeout(() => {
      pushHistory(state.activeLang, {
        ts: new Date().toISOString(),
        html: doc.html,
        css: doc.css,
      });
      map[state.activeLang] = snapshotKey;
    }, 1200);
    return () => clearTimeout(timeout);
  }, [doc.html, doc.css, state.activeLang, pushHistory]);

  useEffect(() => {
    const handleKey = (event: KeyboardEvent) => {
      if (!event.ctrlKey) return;
      switch (event.key.toLowerCase()) {
        case 's':
          event.preventDefault();
          saveToStorage(state);
          break;
        case 'p':
          event.preventDefault();
          setMode('preview');
          break;
        case 'y':
          // handled by browser history when available
          break;
        case 'e':
          if (event.shiftKey) {
            event.preventDefault();
            setMode(state.mode === 'code' ? 'visual' : 'code');
          }
          break;
        case '/':
          event.preventDefault();
          alert('Ctrl+/ — швидка вставка блоків. Виберіть блок у правій панелі.');
          break;
        default:
          break;
      }
    };
    window.addEventListener('keydown', handleKey);
    return () => window.removeEventListener('keydown', handleKey);
  }, [state, setMode]);

  useEffect(() => {
    if (isHosted) return undefined;
    const interval = setInterval(() => saveToStorage(state), 4000);
    return () => clearInterval(interval);
  }, [state]);

  const handleExport = () => {
    const doc = toDescDoc(state.activeLang, state.docs[state.activeLang]);
    const blob = exportState(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${doc.lang}.desc.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImport = async (file?: File | null) => {
    if (!file) return;
    const imported = await parseImport(file);
    setLang(imported.lang);
    updateHtml(imported.lang, imported.html);
    updateCss(imported.lang, imported.css);
    setAssets(imported.lang, imported.assets ?? []);
  };

  const handleSendToApp = async () => {
    if (!isHosted || isSending) return;
    setIsSending(true);
    try {
      await sendStateToHost(state);
      alert('Дані передано у застосунок. Поверніться до нього та закрийте вкладку.');
      window.location.href = '/close';
    } catch (error) {
      console.error('Не вдалося передати дані у застосунок', error);
      alert(error instanceof Error ? error.message : 'Не вдалося передати дані у застосунок');
    } finally {
      setIsSending(false);
    }
  };

  return (
    <div className={clsx('workspace', { 'workspace--fullscreen': fullScreen })}>
      <header className="workspace__top">
        <LangTabs />
        <ModeTabs />
        <div className="workspace__actions">
          <button type="button" onClick={handleExport}>
            Зберегти .desc
          </button>
          <button type="button" onClick={() => importInputRef.current?.click()}>
            Завантажити
          </button>
          <button type="button" onClick={() => setFullScreen((v) => !v)}>
            {fullScreen ? 'Звичайний режим' : 'Повноекранний режим'}
          </button>
          <button type="button" onClick={() => { clearStorage(); window.location.reload(); }}>
            Очистити кеш
          </button>
          {isHosted && (
            <button type="button" onClick={handleSendToApp} disabled={isSending}>
              {isSending ? 'Збереження…' : 'Зберегти в застосунок'}
            </button>
          )}
          <input
            type="file"
            accept="application/json"
            hidden
            ref={importInputRef}
            onChange={(event) => {
              handleImport(event.target.files?.[0]);
              event.target.value = '';
            }}
          />
        </div>
      </header>
      <div className="workspace__body">
        <main className="workspace__main">
          {state.mode === 'visual' && <VisualEditor />}
          {state.mode === 'code' && <CodeEditor />}
          {state.mode === 'preview' && <PreviewPane />}
        </main>
        <aside className="workspace__sidebar">
          <Assets />
          <Blocks />
          <History />
        </aside>
      </div>
    </div>
  );
};

const AppShell = React.forwardRef<DescEditorRef, {}>((_, ref) => {
  const {
    state,
    replaceState,
    setLang,
    updateHtml,
    updateCss,
    setAssets,
    setMode,
  } = useEditorActions();

  useEffect(() => {
    if (isHosted) return;
    const persisted = loadFromStorage();
    if (persisted) {
      replaceState(buildStateFromDocs(persisted.docs as Record<'uk' | 'ru' | 'en', DescDoc>, persisted.activeLang));
    }
  }, [replaceState]);

  useEffect(() => {
    if (!isHosted) return;
    let cancelled = false;
    loadHostState()
      .then((payload) => {
        if (cancelled || !payload || !payload.docs) return;
        const docs = payload.docs as Record<'uk' | 'ru' | 'en', DescDoc>;
        const nextLang = (payload.activeLang ?? 'uk') as Lang;
        replaceState(buildStateFromDocs(docs, nextLang));
        setLang(nextLang);
        setMode('visual');
      })
      .catch((error) => {
        console.error('Не вдалося отримати стан із застосунку', error);
      });
    return () => {
      cancelled = true;
    };
  }, [replaceState, setLang, setMode]);

  useEffect(() => {
    const global = window as typeof window & { __DESC_EDITOR__?: Record<string, unknown> };
    const host = (global.__DESC_EDITOR__ = global.__DESC_EDITOR__ || {});
    host.setState = (payload: { activeLang?: Lang; docs?: Record<Lang, DescDoc> }) => {
      if (!payload || !payload.docs) return;
      const nextLang = (payload.activeLang ?? 'uk') as Lang;
      replaceState(buildStateFromDocs(payload.docs as Record<Lang, DescDoc>, nextLang));
    };
    host.setLang = (lang: Lang) => setLang(lang);
    host.setMode = (mode: EditorState['mode']) => setMode(mode);
    return () => {
      delete global.__DESC_EDITOR__;
    };
  }, [replaceState, setLang, setMode]);

  useEffect(() => {
    const global = window as typeof window & { __DESC_EDITOR__?: Record<string, unknown> };
    const host = (global.__DESC_EDITOR__ = global.__DESC_EDITOR__ || {});
    host.getState = () => {
      const docs = (['uk', 'ru', 'en'] as const).reduce(
        (acc, lang) => {
          acc[lang] = toDescDoc(lang, state.docs[lang]);
          return acc;
        },
        {} as Record<Lang, DescDoc>,
      );
      return JSON.stringify({ activeLang: state.activeLang, docs });
    };
  }, [state]);

  useImperativeHandle(ref, () => ({
    getValue: (lang = state.activeLang) => toDescDoc(lang, state.docs[lang]),
    setValue: (doc) => {
      setLang(doc.lang);
      updateHtml(doc.lang, doc.html);
      updateCss(doc.lang, doc.css);
      setAssets(doc.lang, doc.assets ?? []);
    },
    importJson: async (file: File) => {
      const imported = await parseImport(file);
      setLang(imported.lang);
      updateHtml(imported.lang, imported.html);
      updateCss(imported.lang, imported.css);
      setAssets(imported.lang, imported.assets ?? []);
    },
    exportJson: () => {
      const doc = toDescDoc(state.activeLang, state.docs[state.activeLang]);
      return exportState(doc);
    },
    toHtmlBundle: (lang = state.activeLang) => toBundle(state.docs[lang]),
  }), [setLang, setAssets, state, updateCss, updateHtml]);

  return <Workspace />;
});

AppShell.displayName = 'AppShell';

interface AppProps {
  initial?: Partial<Record<'uk' | 'ru' | 'en', DescDoc>>;
}

const DescEditorRoot = React.forwardRef<DescEditorRef, AppProps>(({ initial }, ref) => (
  <EditorProvider initial={initial}>
    <AppShell ref={ref} />
  </EditorProvider>
));

DescEditorRoot.displayName = 'DescEditorRoot';

const App: React.FC<AppProps> = ({ initial }) => <DescEditorRoot initial={initial} />;

export default App;
export { AppShell, DescEditorRoot };
