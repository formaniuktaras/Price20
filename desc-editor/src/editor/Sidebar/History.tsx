import React from 'react';
import dayjs from 'dayjs';
import { useEditorActions } from '../../core/state';

const History: React.FC = () => {
  const {
    state: { activeLang, docs },
    updateHtml,
    updateCss,
  } = useEditorActions();
  const doc = docs[activeLang];

  const restore = (index: number) => {
    const item = doc.history[index];
    if (!item) return;
    updateHtml(activeLang, item.html);
    updateCss(activeLang, item.css);
  };

  return (
    <div className="sidebar-section">
      <div className="sidebar-section__header">
        <h3>Історія</h3>
      </div>
      <ul className="history-list">
        {doc.history.map((entry, index) => (
          <li key={entry.ts}>
            <div>
              <strong>{dayjs(entry.ts).format('DD.MM HH:mm:ss')}</strong>
            </div>
            <button type="button" onClick={() => restore(index)}>
              Відкотити
            </button>
          </li>
        ))}
      </ul>
      {!doc.history.length && <p className="sidebar-hint">Ще немає збережених версій.</p>}
    </div>
  );
};

export default History;
