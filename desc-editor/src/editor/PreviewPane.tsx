import React, { useEffect, useMemo, useRef, useState } from 'react';
import debounce from 'lodash.debounce';
import { useEditorActions } from '../core/state';

const PreviewPane: React.FC = () => {
  const {
    state: { activeLang, docs },
  } = useEditorActions();
  const doc = docs[activeLang];
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const [auto, setAuto] = useState(true);

  const render = useMemo(
    () =>
      debounce((html: string, css: string) => {
        const iframe = iframeRef.current;
        if (!iframe) return;
        const doc = iframe.contentDocument;
        if (!doc) return;
        doc.open();
        doc.write(`<!DOCTYPE html><html><head><meta charset="utf-8" />` +
          `<style>body{font-family:'Inter',sans-serif;padding:24px;background:#fff;color:#111;} ${css}</style>` +
          `</head><body>${html}</body></html>`);
        doc.close();
      }, 500),
    []
  );

  useEffect(() => {
    if (auto) {
      render(doc.html, doc.css);
    }
  }, [doc.html, doc.css, auto, render]);

  return (
    <div className="preview-pane">
      <div className="preview-pane__toolbar">
        <label>
          <input type="checkbox" checked={auto} onChange={(e) => setAuto(e.target.checked)} /> Автооновлення
        </label>
        <button type="button" onClick={() => render(doc.html, doc.css)}>
          Оновити прев'ю
        </button>
      </div>
      <iframe ref={iframeRef} className="preview-pane__frame" title="Live preview" sandbox="allow-same-origin"></iframe>
    </div>
  );
};

export default PreviewPane;
