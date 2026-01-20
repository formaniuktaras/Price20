import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { DescDoc } from './core/state';
import './styles/app.css';

declare global {
  interface Window {
    __DESC_INITIAL__?: Partial<Record<'uk' | 'ru' | 'en', DescDoc>>;
  }
}

const initial = window.__DESC_INITIAL__;

ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(
  <React.StrictMode>
    <App initial={initial} />
  </React.StrictMode>
);
