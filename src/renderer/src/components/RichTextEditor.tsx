import { useEffect, useRef, useState } from 'react';

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
 * inserção de imagem embutida e edição de tamanho por diálogo dedicado
 * (pré-visualização, largura/altura, cadeado de proporção e controle
 * deslizante) — nada de alças de arraste sobre o conteúdo, que falhavam
 * conforme o layout/scrol do editor.
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
  const ref = useRef<HTMLDivElement>(null);
  const [selImg, setSelImg] = useState<HTMLImageElement | null>(null);
  const [imgDlgOpen, setImgDlgOpen] = useState(false);
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

  // mousedown na imagem: preventDefault impede a seleção nativa e o drag
  // fantasma do browser; clique fora dela desmarca.
  const onEditorMouseDown = (e: React.MouseEvent) => {
    const t = e.target as HTMLElement;
    if (t instanceof HTMLImageElement && ref.current?.contains(t)) {
      e.preventDefault();
      markSel(t);
      return;
    }
    if (selImg) markSel(null);
  };

  const onEditorDblClick = (e: React.MouseEvent) => {
    const t = e.target as HTMLElement;
    if (t instanceof HTMLImageElement && ref.current?.contains(t)) {
      e.preventDefault();
      markSel(t);
      setImgDlgOpen(true);
    }
  };

  // Se a imagem selecionada sumir do documento (Delete/Backspace), limpa o estado.
  const liveSelImg = selImg && selImg.isConnected ? selImg : null;

  const applyImgSize = (w: number, h: number) => {
    if (!liveSelImg) return;
    liveSelImg.style.width = `${w}px`;
    liveSelImg.style.height = `${h}px`;
    sync();
  };

  const removeImg = () => {
    if (!liveSelImg) return;
    liveSelImg.remove();
    markSel(null);
    setImgDlgOpen(false);
    sync();
  };

  return (
    <div className="rte">
      <div className="compose-toolbar" style={{ flexWrap: 'wrap' }}>
        <Tb title="Negrito" onClick={() => exec('bold')}><b>B</b></Tb>
        <Tb title="Itálico" onClick={() => exec('italic')}><i>I</i></Tb>
        <Tb title="Sublinhado" onClick={() => exec('underline')}><u>U</u></Tb>
        <Tb title="Tachado" onClick={() => exec('strikeThrough')}><s>S</s></Tb>
        <span className="tb-sep" />
        <select
          className="tb-fontsize tb-fontname"
          title="Fonte"
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => {
            if (e.target.value) exec('fontName', e.target.value);
            e.target.value = '';
          }}
          defaultValue=""
        >
          <option value="" disabled>Fonte</option>
          {FONTS.map((f) => (
            <option key={f} value={f} style={{ fontFamily: f }}>
              {f}
            </option>
          ))}
        </select>
        <select
          className="tb-fontsize"
          title="Tamanho da fonte"
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => {
            if (e.target.value) exec('fontSize', e.target.value);
            e.target.value = '';
          }}
          defaultValue=""
        >
          <option value="" disabled>Tamanho</option>
          <option value="2">Pequena</option>
          <option value="3">Normal</option>
          <option value="4">Média</option>
          <option value="5">Grande</option>
          <option value="6">Muito grande</option>
        </select>
        <input
          type="color"
          className="tb-color"
          title="Cor do texto"
          defaultValue="#000000"
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => exec('foreColor', e.target.value)}
        />
        <Tb title="Alinhar à esquerda" onClick={() => exec('justifyLeft')}>⇤</Tb>
        <Tb title="Centralizar" onClick={() => exec('justifyCenter')}>≡</Tb>
        <Tb title="Alinhar à direita" onClick={() => exec('justifyRight')}>⇥</Tb>
        <span className="tb-sep" />
        <Tb title="Lista com marcadores" onClick={() => exec('insertUnorderedList')}>•</Tb>
        <Tb title="Lista numerada" onClick={() => exec('insertOrderedList')}>1.</Tb>
        <span className="tb-sep" />
        <Tb title="Inserir link" onClick={insertLink}>🔗</Tb>
        <Tb title="Inserir imagem" onClick={() => void insertImage()}>🖼</Tb>
        <Tb title="Limpar formatação" onClick={() => exec('removeFormat')}>⌫</Tb>
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
          <button className="btn btn-primary" onClick={applyLink}>Aplicar</button>
          <button className="btn btn-ghost" onClick={() => setLinkMode(false)}>Cancelar</button>
        </div>
      )}

      {/* Ações da imagem selecionada — a edição de tamanho acontece no diálogo */}
      {liveSelImg && (
        <div className="img-panel">
          <span className="img-panel-title">Imagem selecionada</span>
          <PanelBtn onClick={() => setImgDlgOpen(true)}>Tamanho da imagem…</PanelBtn>
          <PanelBtn className="danger" onClick={removeImg}>Remover</PanelBtn>
          <span className="img-panel-hint">Duplo clique na imagem também abre o ajuste de tamanho.</span>
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
        onDoubleClick={onEditorDblClick}
        onDragStart={(e) => {
          // imagem não pode ser arrastada de lugar — evita o drag fantasma do browser
          if (e.target instanceof HTMLImageElement) e.preventDefault();
        }}
        onBlur={sync}
      />

      {imgDlgOpen && liveSelImg && (
        <ImageSizeDialog
          img={liveSelImg}
          onApply={applyImgSize}
          onRemove={removeImg}
          onClose={() => setImgDlgOpen(false)}
        />
      )}
    </div>
  );
}

/**
 * Diálogo "Tamanho da imagem" — pré-visualização ao vivo numa área com rolagem
 * própria, largura/altura independentes e cadeado de proporção. Nada é aplicado
 * à imagem real até clicar em Aplicar.
 */
function ImageSizeDialog({
  img,
  onApply,
  onRemove,
  onClose,
}: {
  img: HTMLImageElement;
  onApply: (w: number, h: number) => void;
  onRemove: () => void;
  onClose: () => void;
}) {
  const natW = img.naturalWidth || Math.round(img.getBoundingClientRect().width) || 100;
  const natH = img.naturalHeight || Math.round(img.getBoundingClientRect().height) || 100;
  const ratio = natW / natH;
  const cur = img.getBoundingClientRect();
  const [w, setW] = useState(Math.max(8, Math.round(cur.width) || natW));
  const [h, setH] = useState(Math.max(8, Math.round(cur.height) || natH));
  const [lock, setLock] = useState(true);
  const pct = Math.max(1, Math.round((w / natW) * 100));

  const setWidth = (v: number) => {
    const nv = Math.max(8, Math.round(v) || 8);
    setW(nv);
    if (lock) setH(Math.max(8, Math.round(nv / ratio)));
  };
  const setHeight = (v: number) => {
    const nv = Math.max(8, Math.round(v) || 8);
    setH(nv);
    if (lock) setW(Math.max(8, Math.round(nv * ratio)));
  };

  return (
    <div
      className="modal-overlay imgdlg-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="imgdlg" role="dialog" aria-label="Tamanho da imagem">
        <div className="imgdlg-head">
          <h2>Tamanho da imagem</h2>
          <button className="imgdlg-close" onClick={onClose} title="Fechar">✕</button>
        </div>

        <div className="imgdlg-preview">
          {/* clone da imagem na prévia — o original só muda ao aplicar */}
          <img src={img.src} alt="" draggable={false} style={{ width: w, height: h }} />
        </div>

        <div className="imgdlg-controls">
          <label className="img-field">
            Largura
            <input
              className="inp"
              type="number"
              min={8}
              value={w}
              onChange={(e) => setWidth(Number(e.target.value))}
            />
            px
          </label>
          <label className="img-field">
            Altura
            <input
              className="inp"
              type="number"
              min={8}
              value={h}
              onChange={(e) => setHeight(Number(e.target.value))}
            />
            px
          </label>
          <label className="img-check">
            <input type="checkbox" checked={lock} onChange={(e) => setLock(e.target.checked)} />
            Manter proporção
          </label>
        </div>

        <div className="imgdlg-controls" style={{ paddingTop: 0 }}>
          <input
            type="range"
            className="imgdlg-range"
            min={5}
            max={400}
            value={Math.min(400, pct)}
            onChange={(e) => setWidth((natW * Number(e.target.value)) / 100)}
            title="Escala em relação ao tamanho original"
          />
          <span className="imgdlg-info">
            {pct}% do original ({natW} × {natH} px)
          </span>
        </div>

        <div className="imgdlg-foot">
          <button className="btn btn-ghost danger" onClick={onRemove}>Remover imagem</button>
          <span style={{ flex: 1 }} />
          <button className="btn" onClick={onClose}>Cancelar</button>
          <button
            className="btn btn-primary"
            onClick={() => {
              onApply(Math.max(8, w), Math.max(8, h));
              onClose();
            }}
          >
            Aplicar
          </button>
        </div>
      </div>
    </div>
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
