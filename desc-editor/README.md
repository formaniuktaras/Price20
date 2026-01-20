# Desc Editor

Модуль «Редактор шаблону опису» для інтеграції у Prom.ua-подібні інтерфейси. Побудований на React + Vite з підтримкою двох режимів (Visual/Code) і Live Preview.

## Запуск демо

```bash
npm install
npm run dev
```

## Інтеграція з десктопним застосунком

`main.tsx` запускає редактор як окремий SPA, який очікує, що десктопний хост передасть початкові дані через `window.__DESC_INITIAL__`. Після збірки бандл вбудовується у Tkinter-додаток: він піднімає локальний HTTP-сервер, а редактор спілкується з ним через REST (`/api/session/<id>`). Якщо потрібно керувати станом вручну під час розробки, скористайтеся глобальним об'єктом `window.__DESC_EDITOR__`:

```ts
window.__DESC_EDITOR__?.setState({
  activeLang: 'uk',
  docs: {
    uk: { lang: 'uk', html: '<h1>...</h1>', css: '', assets: [] },
    ru: { lang: 'ru', html: '<h1>...</h1>', css: '', assets: [] },
    en: { lang: 'en', html: '<h1>...</h1>', css: '', assets: [] },
  },
});
```

Також доступні утиліти `getState`, `setLang`, `setMode` та `hostApi` (див. `src/core/host.ts`).

## Структура .desc.json

```json
{
  "lang": "uk",
  "html": "<h1>...</h1>",
  "css": "body { ... }",
  "assets": [{ "name": "hero.png", "dataUrl": "data:image/png;base64,..." }],
  "history": [
    { "ts": "2024-01-20T12:00:00.000Z", "html": "...", "css": "..." }
  ],
  "exportedAt": "2024-01-20T12:00:00.000Z"
}
```

## Гарячі клавіші

| Комбінація        | Дія                                   |
| ----------------- | ------------------------------------- |
| `Ctrl + S`        | Зберегти стан в localStorage          |
| `Ctrl + P`        | Відкрити прев'ю                       |
| `Ctrl + Shift + E`| Перемикання Code/Visual               |
| `Ctrl + /`        | Підказка щодо панелі «Шаблони/блоки»  |

## Основні можливості

- CKEditor 5 Classic у Visual Mode з кастомним тулбаром.
- Monaco Editor для HTML/CSS з двома вкладками.
- Live Preview через sandboxed iframe.
- Автозбереження до `localStorage` та експорт/імпорт `.desc.json`.
- Багатомовні вкладки (UA/RU/EN) з окремим станом.
- Панель Assets із drag&drop-файлами та вставкою в HTML.
- Готові блоки (hero, alert, списки, FAQ, таблиця) зі швидким налаштуванням.
- Історія версій (до 25 знімків) з можливістю відкату.
