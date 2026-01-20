import React, { useEffect, useMemo, useRef } from 'react';
import { CKEditor } from '@ckeditor/ckeditor5-react';
import ClassicEditor from '@ckeditor/ckeditor5-build-classic';
import { useEditorActions } from '../core/state';

const VisualEditor: React.FC = () => {
  const {
    state: { activeLang, docs },
    updateHtml,
    setMode,
  } = useEditorActions();
  const doc = docs[activeLang];
  const lastHtml = useRef(doc.html);

  useEffect(() => {
    lastHtml.current = doc.html;
  }, [doc.html]);

  const config = useMemo(
    () => ({
      toolbar: [
        'heading',
        '|',
        'bold',
        'italic',
        'underline',
        'strikethrough',
        '|',
        'bulletedList',
        'numberedList',
        '|',
        'alignment',
        '|',
        'blockQuote',
        'link',
        'insertTable',
        'imageUpload',
        '|',
        'undo',
        'redo',
      ],
      table: {
        contentToolbar: [
          'tableColumn',
          'tableRow',
          'mergeTableCells',
          'tableProperties',
          'tableCellProperties',
        ],
      },
      heading: {
        options: [
          { model: 'paragraph', title: 'Paragraph', class: 'ck-heading_paragraph' },
          { model: 'heading1', view: 'h1', title: 'Заголовок H1', class: 'ck-heading_heading1' },
          { model: 'heading2', view: 'h2', title: 'Заголовок H2', class: 'ck-heading_heading2' },
          { model: 'heading3', view: 'h3', title: 'Заголовок H3', class: 'ck-heading_heading3' },
        ],
      },
      language: 'uk',
    }),
    []
  );

  return (
    <div className="visual-editor">
      <div className="visual-editor__toolbar">
        <button type="button" onClick={() => setMode('code')} className="visual-editor__toolbar-btn">
          Джерело
        </button>
      </div>
      <CKEditor
        editor={ClassicEditor}
        data={doc.html}
        config={config}
        onChange={(_, editor) => {
          const data = editor.getData();
          if (data !== lastHtml.current) {
            lastHtml.current = data;
            updateHtml(activeLang, data);
          }
        }}
      />
    </div>
  );
};

export default VisualEditor;
