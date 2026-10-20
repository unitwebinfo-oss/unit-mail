import { useEffect, useRef, useState } from 'react';
import { t, useLang } from '../i18n';

/** Fontes oferecidas na barra de formatação — conjunto seguro p/ email. */
const FONTS = [
  'Calibri',
  'Arial',
  'Segoe UI',
  'Tahoma',
  'Verdana',
  'Georgia',
  'Times New Roman',
  'Courier New',
];

/**
 * Editor de texto rico (WYSIWYG) — usado na edição de assinaturas.
 * Renderiza o HTML visualmente (não como código), com toolbar de formatação,
 * inserção de imagem embutida e redimensionamento DIRETO no corpo: clique na
 * imagem e arraste a alça do canto inferior direito (proporção mantida) —
 * estilo Word/Outlook, sem abrir diálogo.
 */
export function RichTextEditor({
  value,
  onChange,
  placeholder,
  minHeight = 320,
}: {
  value: string;
  onChange: (html: string) => void;
  placeholder?: string;
  minHeight?: number;
}) {
  useLang();
  const ref = useRef<HTMLDivElement>(null);
  const [selImg, setSelImg] = useState<HTMLImageElement | null>(null);
  const loaded = useRef(false);

  // Carrega o HTML inicial apenas uma vez por troca de documento —
  // depois disso o usuário edita visualmente e onChange propaga o HTML.
  useEffect(() => {
    if (ref.current && !loaded.current) {
      ref.current.innerHTML = value || '';
      loaded.current = true;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const exec = (cmd: string, val?: string) => {
    document.execCommand(cmd, false, val);
    ref.current?.focus();
    sync();
  };

  const sync = () => {
    if (ref.current) onChange(ref.current.innerHTML);
  };

  const insertImage = async () => {
    const dataUri = await window.outlook.pickImage?.();
    if (!dataUri) return;
    ref.current?.focus();
    document.execCommand('insertImage', false, dataUri);
    sync();
  };

  // window.prompt() não existe no Electron — campo inline para a URL do link.
  const [linkMode, setLinkMode] = useState(false);
  const [linkUrl, setLinkUrl] = useState('https://');
  const savedRange = useRef<Range | null>(null);

  const insertLink = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount) savedRange.current = sel.getRangeAt(0).cloneRange();
    setLinkUrl('https://');
    setLinkMode(true);
  };

  const applyLink = () => {
    if (linkUrl && linkUrl !== 'https://') {
      const sel = window.getSelection();
      if (sel && savedRange.current) {
        sel.removeAllRanges();
        sel.addRange(savedRange.current);
      }
      exec('createLink', linkUrl);
    }
    setLinkMode(false);
  };

  const markSel = (img: HTMLImageElement | null) => {
    ref.current?.querySelectorAll('img.rte-img-sel').forEach((el) => el.classList.remove('rte-img-sel'));
    img?.classList.add('rte-img-sel');
    setSelImg(img);
  };

  /**
   * Arraste de redimensionamento — começa na alça do canto (ou no hit-test do
   * próprio canto da imagem) e acompanha o mouse no documento inteiro, então
   * funciona mesmo quando o ponteiro sai do editor. Proporção sempre mantida.
   */
  const startResize = (img: HTMLImageElement, e: { clientX: number; clientY: number }) => {
    const r = img.getBoundingClientRect();
    const w0 = Math.max(8, Math.round(r.width));
    const h0 = Math.max(8, Math.round(r.height));
    const ratio =
      img.naturalWidth > 0 && img.naturalHeight > 0
        ? img.naturalWidth / img.naturalHeight
        : w0 / h0;
    const x0 = e.clientX;
    const move = (ev: MouseEvent) => {
      const w = Math.max(8, Math.round(w0 + (ev.clientX - x0)));
      img.style.width = `${w}px`;
      img.style.height = `${Math.max(8, Math.round(w / ratio))}px`;
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      sync();
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  };

  /** O clique foi no canto inferior direito da imagem (zona de arraste)? */
  const inGripZone = (img: HTMLImageElement, e: { clientX: number; clientY: number }) => {
    const r = img.getBoundingClientRect();
    return e.clientX > r.right - 16 && e.clientY > r.bottom - 16;
  };

  // mousedown na imagem: no canto = redimensionar; no resto = selecionar.
  // preventDefault impede a seleção nativa e o drag fantasma do browser;
  // clique fora dela desmarca.
  const onEditorMouseDown = (e: React.MouseEvent) => {
    const t = e.target as HTMLElement;
    if (t instanceof HTMLImageElement && ref.current?.contains(t)) {
      e.preventDefault();
      markSel(t);
      if (inGripZone(t, e)) startResize(t, e);
      return;
    }
    if (selImg) markSel(null);
  };

  // Cursor "redimensionar" aparece só sobre a zona da alça — dica visual.
  const onEditorMouseMove = (e: React.MouseEvent) => {
    const t = e.target as HTMLElement;
    if (t instanceof HTMLImageElement && ref.current?.contains(t)) {
      t.style.cursor = inGripZone(t, e) ? 'nwse-resize' : '';
    }
  };

  // Se a imagem selecionada sumir do documento (Delete/Backspace), limpa o estado.
  const liveSelImg = selImg && selImg.isConnected ? selImg : null;

  const removeImg = () => {
    if (!liveSelImg) return;
    liveSelImg.remove();
    markSel(null);
    sync();
  };

  return (
    <div className="rte">
      <div className="compose-toolbar" style={{ flexWrap: 'wrap' }}>
        <Tb title={t('Negrito')} onClick={() => exec('bold')}><b>B</b></Tb>
        <Tb title={t('Itálico')} onClick={() => exec('italic')}><i>I</i></Tb>
        <Tb title={t('Sublinhado')} onClick={() => exec('underline')}><u>U</u></Tb>
        <Tb title={t('Tachado')} onClick={() => exec('strikeThrough')}><s>S</s></Tb>
        <span className="tb-sep" />
        <select
          className="tb-fontsize tb-fontname"
          title={t('Fonte')}
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => {
            if (e.target.value) exec('fontName', e.target.value);
            e.target.value = '';
          }}
          defaultValue=""
        >
          <option value="" disabled>{t('Fonte')}</option>
          {FONTS.map((f) => (
            <option key={f} value={f} style={{ fontFamily: f }}>
              {f}
            </option>
          ))}
        </select>
        <select
          className="tb-fontsize"
          title={t('Tamanho da fonte')}
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => {
            if (e.target.value) exec('fontSize', e.target.value);
            e.target.value = '';
          }}
          defaultValue=""
        >
          <option value="" disabled>{t('Tamanho')}</option>
          <option value="2">{t('Pequena')}</option>
          <option value="3">{t('Normal')}</option>
          <option value="4">{t('Média')}</option>
          <option value="5">{t('Grande')}</option>
          <option value="6">{t('Muito grande')}</option>
        </select>
        <input
          type="color"
          className="tb-color"
          title={t('Cor do texto')}
          defaultValue="#000000"
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => exec('foreColor', e.target.value)}
        />
        <Tb title={t('Alinhar à esquerda')} onClick={() => exec('justifyLeft')}>⇤</Tb>
        <Tb title={t('Centralizar')} onClick={() => exec('justifyCenter')}>≡</Tb>
        <Tb title={t('Alinhar à direita')} onClick={() => exec('justifyRight')}>⇥</Tb>
        <span className="tb-sep" />
        <Tb title={t('Lista com marcadores')} onClick={() => exec('insertUnorderedList')}>•</Tb>
        <Tb title={t('Lista numerada')} onClick={() => exec('insertOrderedList')}>1.</Tb>
        <span className="tb-sep" />
        <Tb title={t('Inserir link')} onClick={insertLink}>🔗</Tb>
        <Tb title={t('Inserir imagem')} onClick={() => void insertImage()}>🖼</Tb>
        <Tb title={t('Limpar formatação')} onClick={() => exec('removeFormat')}>⌫</Tb>
      </div>

      {linkMode && (
        <div className="img-toolbar">
          <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>Link:</span>
          <input
            className="inp"
            style={{ flex: 1, minWidth: 200 }}
            autoFocus
            value={linkUrl}
            onChange={(e) => setLinkUrl(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') applyLink();
              if (e.key === 'Escape') setLinkMode(false);
            }}
            placeholder="https://…"
          />
          <button className="btn btn-primary" onClick={applyLink}>{t('Aplicar')}</button>
          <button className="btn btn-ghost" onClick={() => setLinkMode(false)}>{t('Cancelar')}</button>
        </div>
      )}

      {/* Ações da imagem selecionada — o tamanho se ajusta arrastando a alça */}
      {liveSelImg && (
        <div className="img-panel">
          <span className="img-panel-title">{t('Imagem selecionada')}</span>
          <span className="img-panel-hint">
            {t('Arraste a alça do canto inferior direito para redimensionar.')}
          </span>
          <PanelBtn className="danger" onClick={removeImg}>{t('Remover')}</PanelBtn>
        </div>
      )}

      <div
        ref={ref}
        className="rte-editor"
        style={{ minHeight }}
        contentEditable
        suppressContentEditableWarning
        data-placeholder={placeholder || ''}
        onInput={sync}
        onMouseDown={onEditorMouseDown}
        onMouseMove={onEditorMouseMove}
        onDragStart={(e) => {
          // imagem não pode ser arrastada de lugar — evita o drag fantasma do browser
          if (e.target instanceof HTMLImageElement) e.preventDefault();
        }}
        onBlur={sync}
      />

      {/* Alça de redimensionamento sobre o canto da imagem — fora do
          contentEditable para não entrar no HTML; position:fixed acompanha
          o rect da imagem e é recalculada em scroll/resize/reflows. */}
      {liveSelImg && (
        <ImgGrip img={liveSelImg} onGripDown={(e) => {
          e.preventDefault();
          startResize(liveSelImg, e);
        }} />
      )}
    </div>
  );
}

/**
 * Alça visual no canto inferior direito da imagem selecionada. Renderizada
 * fora do contentEditable (não pode entrar no HTML) com position:fixed —
 * a posição segue o bounding rect da imagem e é recalculada em scroll,
 * resize e num tick leve (reflows de fonte/layout movem a imagem).
 */
function ImgGrip({
  img,
  onGripDown,
}: {
  img: HTMLImageElement;
  onGripDown: (e: React.MouseEvent) => void;
}) {
  const [pos, setPos] = useState({ left: -999, top: -999 });
  useEffect(() => {
    const place = () => {
      const r = img.getBoundingClientRect();
      setPos({ left: r.right - 8, top: r.bottom - 8 });
    };
    place();
    window.addEventListener('scroll', place, true);
    window.addEventListener('resize', place);
    const tick = setInterval(place, 200);
    return () => {
      window.removeEventListener('scroll', place, true);
      window.removeEventListener('resize', place);
      clearInterval(tick);
    };
  }, [img]);
  return (
    <div
      className="rte-grip"
      style={{ left: pos.left, top: pos.top }}
      title={t('Arraste para redimensionar')}
      onMouseDown={onGripDown}
    />
  );
}

/** Botão do painel de imagem — não rouba o foco nem a seleção do editor. */
function PanelBtn({
  onClick,
  className = '',
  children,
}: {
  onClick: () => void;
  className?: string;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      className={`btn btn-ghost ${className}`}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}

function Tb({ title, onClick, children }: { title: string; onClick: () => void; children: React.ReactNode }) {
  return (
    <button type="button" title={title} className="tb-btn" onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      {children}
    </button>
  );
}
