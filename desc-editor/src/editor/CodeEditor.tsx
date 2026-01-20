import React, { useMemo, useState } from 'react';
import Editor from '@monaco-editor/react';
import debounce from 'lodash.debounce';
import { useEditorActions } from '../core/state';

const CodeEditor: React.FC = () => {
  const {
    state: { activeLang, docs },
    updateHtml,
    updateCss,
  } = useEditorActions();
  const doc = docs[activeLang];
  const [tab, setTab] = useState<'html' | 'css'>('html');

  const handleChange = useMemo(
    () =>
      debounce((value?: string) => {
        if (value === undefined) return;
        if (tab === 'html') {
          updateHtml(activeLang, value);
        } else {
          updateCss(activeLang, value);
        }
      }, 300),
    [tab, updateHtml, updateCss, activeLang]
  );

  return (
    <div className="code-editor">
      <div className="code-editor__tabs">
        <button
          type="button"
          className={tab === 'html' ? 'active' : ''}
          onClick={() => setTab('html')}
        >
          HTML
        </button>
        <button
          type="button"
          className={tab === 'css' ? 'active' : ''}
          onClick={() => setTab('css')}
        >
          CSS
        </button>
      </div>
      <div className="code-editor__pane">
        <Editor
          height="100%"
          language={tab === 'html' ? 'html' : 'css'}
          value={tab === 'html' ? doc.html : doc.css}
          theme="vs"
          options={{
            minimap: { enabled: false },
            fontSize: 14,
            automaticLayout: true,
            scrollBeyondLastLine: false,
          }}
          onChange={handleChange}
        />
      </div>
    </div>
  );
};

export default CodeEditor;
