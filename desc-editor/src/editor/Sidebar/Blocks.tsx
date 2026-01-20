import React from 'react';
import { blocks, buildBlockHtml } from '../../core/blocks';
import { useEditorActions } from '../../core/state';

const Blocks: React.FC = () => {
  const {
    state: { activeLang, docs },
    updateHtml,
  } = useEditorActions();
  const doc = docs[activeLang];

  const insertBlock = (id: string) => {
    const block = blocks.find((item) => item.id === id);
    if (!block) return;
    const params: Record<string, string> = {};
    block.fields?.forEach((field) => {
      const value = window.prompt(field.label, field.placeholder ?? '') ?? field.placeholder ?? '';
      params[field.name] = value;
    });
    const html = doc.html + buildBlockHtml(block.id, params);
    updateHtml(activeLang, html);
  };

  return (
    <div className="sidebar-section">
      <div className="sidebar-section__header">
        <h3>Шаблони</h3>
      </div>
      <ul className="blocks-list">
        {blocks.map((block) => (
          <li key={block.id}>
            <div>
              <strong>{block.name}</strong>
              <p>{block.description}</p>
            </div>
            <button type="button" onClick={() => insertBlock(block.id)}>
              Вставити
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
};

export default Blocks;
