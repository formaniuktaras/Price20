export type BlockId =
  | 'hero'
  | 'alert'
  | 'headings'
  | 'benefits'
  | 'faq'
  | 'spec-table';

export interface BlockDefinition {
  id: BlockId;
  name: string;
  description: string;
  template: (params: Record<string, string>) => string;
  fields?: Array<{ name: string; label: string; placeholder?: string }>;
}

const wrapSection = (content: string) => `\n<section class="desc-block">${content}</section>\n`;

export const blocks: BlockDefinition[] = [
  {
    id: 'hero',
    name: 'Hero банер',
    description: 'Верхній блок із великим заголовком і кнопкою',
    fields: [
      { name: 'brand', label: 'Бренд', placeholder: 'Apple' },
      { name: 'model', label: 'Модель', placeholder: 'iPhone 15' },
      { name: 'cta', label: 'Текст кнопки', placeholder: 'Придбати зараз' },
    ],
    template: ({ brand = '{{brand}}', model = '{{model}}', cta = 'Придбати' }) =>
      wrapSection(`
        <div class="hero-banner">
          <div class="hero-banner__content">
            <p class="hero-banner__tag">Захист</p>
            <h1>Захистіть свій <strong>${brand} ${model}</strong> від подряпин і падінь</h1>
            <p>Преміальний чохол з приємною текстурою, що зберігає стиль і комфорт користування.</p>
            <a class="hero-banner__cta" href="#order">${cta}</a>
          </div>
        </div>
      `),
  },
  {
    id: 'alert',
    name: 'Інфо-алерт',
    description: 'Кольоровий блок з інформацією',
    fields: [
      { name: 'text', label: 'Текст', placeholder: 'Вказуйте важливу інформацію про гарантію або доставку' },
    ],
    template: ({ text = 'Важливо: доставка безкоштовна при замовленні від 1500 грн' }) =>
      wrapSection(`
        <div class="promo-alert">
          <strong>Зверніть увагу!</strong>
          <p>${text}</p>
        </div>
      `),
  },
  {
    id: 'headings',
    name: 'Заголовки H1–H3',
    description: 'Набір заголовків із плейсхолдерами',
    template: () =>
      wrapSection(`
        <h1>Чохол для {{brand}} {{model}}</h1>
        <h2>Основні переваги</h2>
        <h3>Чому обирають нас</h3>
      `),
  },
  {
    id: 'benefits',
    name: 'Списки переваг',
    description: 'Три списки з емодзі',
    template: () =>
      wrapSection(`
        <div class="benefits-grid">
          <div>
            <h3>🛡️ Захист</h3>
            <ul>
              <li>Багатошаровий ударостійкий корпус</li>
              <li>Виступаючі борти для безпеки дисплея</li>
              <li>Антишокові вставки по кутах</li>
            </ul>
          </div>
          <div>
            <h3>👌 Зручність</h3>
            <ul>
              <li>Не ковзає в руці та не залишає відбитків</li>
              <li>Сумісний з бездротовою зарядкою</li>
              <li>Вирізи точно повторюють кнопки</li>
            </ul>
          </div>
          <div>
            <h3>✨ Вигляд</h3>
            <ul>
              <li>Лаконічний дизайн у стилі {{brand}}</li>
              <li>Матове покриття стійке до подряпин</li>
              <li>Доступно кілька трендових кольорів</li>
            </ul>
          </div>
        </div>
      `),
  },
  {
    id: 'faq',
    name: 'FAQ',
    description: 'Питання-відповіді у форматі details',
    template: () =>
      wrapSection(`
        <section class="faq">
          <h2>Часті запитання</h2>
          <details open>
            <summary>Чи підтримує чохол бездротову зарядку?</summary>
            <p>Так, чохол сумісний із більшістю зарядок стандарту Qi.</p>
          </details>
          <details>
            <summary>Які гарантійні умови?</summary>
            <p>Ми надаємо 12 місяців офіційної гарантії від виробника.</p>
          </details>
          <details>
            <summary>Як доглядати за чохлом?</summary>
            <p>Достатньо протирати його м'якою вологою серветкою без агресивних засобів.</p>
          </details>
        </section>
      `),
  },
  {
    id: 'spec-table',
    name: 'Таблиця характеристик',
    description: 'Таблиця з ключовими параметрами',
    template: () =>
      wrapSection(`
        <table class="spec-table">
          <thead>
            <tr>
              <th>Характеристика</th>
              <th>Значення</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Матеріал</td>
              <td>Термополіуретан + полікарбонат</td>
            </tr>
            <tr>
              <td>Сумісність</td>
              <td>{{brand}} {{model}}</td>
            </tr>
            <tr>
              <td>Гарантія</td>
              <td>12 місяців</td>
            </tr>
          </tbody>
        </table>
      `),
  },
];

export const buildBlockHtml = (id: BlockId, params: Record<string, string> = {}): string => {
  const block = blocks.find((b) => b.id === id);
  if (!block) {
    throw new Error(`Block ${id} is not defined`);
  }
  return block.template(params);
};
