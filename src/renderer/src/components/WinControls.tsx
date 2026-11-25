/**
 * Botões min/max/close da barra de título — usados no Linux, onde a janela
 * abre com frame:false (o titleBarOverlay do Chromium desenhava uma cápsula
 * com sombra própria, que destoava do tema escuro). No Windows o overlay
 * nativo do SO continua desenhando os botões.
 */
export function WinControls() {
  const t = (s: string) => s; // tooltips curtos — sem catálogo de i18n aqui
  const go = (a: 'min' | 'max' | 'close') => window.outlook.winControl?.(a);
  return (
    <div className="win-btns">
      <button
        type="button"
        className="win-btn"
        title={t('Minimizar')}
        aria-label={t('Minimizar')}
        onClick={() => go('min')}
      >
        <svg width="11" height="11" viewBox="0 0 11 11">
          <rect x="0" y="5" width="11" height="1" fill="currentColor" />
        </svg>
      </button>
      <button
        type="button"
        className="win-btn"
        title={t('Maximizar')}
        aria-label={t('Maximizar')}
        onClick={() => go('max')}
      >
        <svg width="10" height="10" viewBox="0 0 10 10">
          <rect x="0.5" y="0.5" width="9" height="9" fill="none" stroke="currentColor" />
        </svg>
      </button>
      <button
        type="button"
        className="win-btn close"
        title={t('Fechar')}
        aria-label={t('Fechar')}
        onClick={() => go('close')}
      >
        <svg width="10" height="10" viewBox="0 0 10 10">
          <path
            d="M0.7 0.7 L9.3 9.3 M9.3 0.7 L0.7 9.3"
            stroke="currentColor"
            strokeWidth="1.1"
          />
        </svg>
      </button>
    </div>
  );
}
