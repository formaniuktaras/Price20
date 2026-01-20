import React, { useRef } from 'react';
import { useEditorActions } from '../../core/state';

const Assets: React.FC = () => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const {
    state: { activeLang, docs },
    setAssets,
    updateHtml,
  } = useEditorActions();
  const doc = docs[activeLang];

  const handleFiles = async (files: FileList | null) => {
    if (!files) return;
    const items = await Promise.all(
      Array.from(files).map(async (file) => {
        const dataUrl = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result));
          reader.onerror = (err) => reject(err);
          reader.readAsDataURL(file);
        });
        return { name: file.name, dataUrl };
      })
    );
    setAssets(activeLang, [...doc.assets, ...items]);
  };

  const handleInsert = (assetUrl: string) => {
    const html = `${doc.html}\n<img src="${assetUrl}" alt="" />`;
    updateHtml(activeLang, html);
  };

  const handleRemove = (name: string) => {
    setAssets(
      activeLang,
      doc.assets.filter((asset) => asset.name !== name)
    );
  };

  return (
    <div className="sidebar-section">
      <div className="sidebar-section__header">
        <h3>Assets</h3>
        <button type="button" onClick={() => fileInputRef.current?.click()}>
          Додати
        </button>
        <input
          type="file"
          accept="image/*"
          ref={fileInputRef}
          hidden
          multiple
          onChange={(event) => handleFiles(event.target.files)}
        />
      </div>
      <ul className="assets-list">
        {doc.assets.map((asset) => (
          <li key={asset.name}>
            <span>{asset.name}</span>
            <div>
              <button type="button" onClick={() => handleInsert(asset.dataUrl)}>
                Вставити
              </button>
              <button type="button" onClick={() => handleRemove(asset.name)}>
                Видалити
              </button>
            </div>
          </li>
        ))}
      </ul>
      <p className="sidebar-hint">Перетягніть файли у це поле або натисніть «Додати».</p>
    </div>
  );
};

export default Assets;
