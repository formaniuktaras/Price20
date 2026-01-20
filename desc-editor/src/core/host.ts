import { DescDoc, EditorState, Lang, toDescDoc } from './state';

const params = new URLSearchParams(window.location.search);
const session = params.get('session');

export const hostSession: string | null = session ? session : null;
export const isHosted = Boolean(hostSession);
export const hostApiBase = hostSession ? `/api/session/${hostSession}` : null;

export type HostStatePayload = {
  activeLang?: Lang;
  docs?: Partial<Record<Lang, DescDoc>>;
};

export const buildHostPayload = (state: EditorState): HostStatePayload => {
  const docs = (['uk', 'ru', 'en'] as const).reduce(
    (acc, lang) => {
      acc[lang] = toDescDoc(lang, state.docs[lang]);
      return acc;
    },
    {} as Record<Lang, DescDoc>,
  );
  return { activeLang: state.activeLang, docs };
};

export const sendStateToHost = async (state: EditorState): Promise<void> => {
  if (!hostApiBase) throw new Error('host API unavailable');
  const response = await fetch(`${hostApiBase}/save`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(buildHostPayload(state)),
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || 'Не вдалося зберегти дані у застосунок');
  }
};

export const loadHostState = async (): Promise<HostStatePayload | null> => {
  if (!hostApiBase) return null;
  const response = await fetch(`${hostApiBase}/state`, { cache: 'no-store' });
  if (!response.ok) {
    throw new Error('Не вдалося отримати початкові дані');
  }
  return (await response.json()) as HostStatePayload;
};
