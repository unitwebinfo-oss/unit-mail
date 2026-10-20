import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { ComposeHost } from './components/ComposeModal';
import { initLang } from './i18n';
import './styles.css';

// Idioma antes do primeiro render — evita flash em pt nas janelas de composição.
void window.outlook.getLanguage?.().then((l) => initLang(l));

// Janelas de composição abrem com hash `#/compose/<id>` — mesma página,
// conteúdo diferente: só o compose em tela cheia, sem o app por trás.
const composeMatch = /^#\/compose\/(.+)$/.exec(window.location.hash);

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    {composeMatch ? <ComposeHost draftId={composeMatch[1]} /> : <App />}
  </React.StrictMode>,
);
