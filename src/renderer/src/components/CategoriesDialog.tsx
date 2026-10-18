import React from 'react';
import { Icon } from '../icons';
import type { MailCategory } from '../../../shared/types';

/** Paleta de cores do Outlook (subconjunto das 25 cores de categoria). */
const PALETTE = [
  '#d13438', '#f7630c', '#eaa300', '#107c10', '#0078d4', '#5c2e91',
  '#b4009e', '#e3008c', '#00b7c3', '#038387', '#498205', '#8a8886',
];

interface Props {
  cats: MailCategory[];
  onSave: (cats: MailCategory[]) => void;
  onClose: () => void;
}

/** Diálogo "Todas as Categorias" — cria, renomeia, muda a cor e exclui
 *  categorias da lista-mestre, como o diálogo homônimo do Outlook. */
export function CategoriesDialog({ cats, onSave, onClose }: Props) {
  const [list, setList] = React.useState<MailCategory[]>(cats);
  /** Linha com a paleta de cores aberta. */
  const [colorEdit, setColorEdit] = React.useState<string | null>(null);

  const upd = (id: string, patch: Partial<MailCategory>) =>
    setList((l) => l.map((c) => (c.id === id ? { ...c, ...patch } : c)));

  const addCat = () => {
    const id = `cat-${Date.now().toString(36)}`;
    const used = new Set(list.map((c) => c.color));
    const free = PALETTE.find((p) => !used.has(p)) ?? PALETTE[list.length % PALETTE.length];
    setList((l) => [...l, { id, name: 'Nova categoria', color: free }]);
    setColorEdit(null);
  };

  const del = (id: string) => setList((l) => l.filter((c) => c.id !== id));

  return (
    <div
      className="modal-overlay"
      style={{ zIndex: 150 }}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="set-card cats-dlg"
        onClick={(e) => {
          e.stopPropagation();
          setColorEdit(null);
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <strong style={{ fontSize: 14 }}>
            <Icon name="flag" size={15} /> Todas as Categorias
          </strong>
          <button className="set-close" onClick={onClose} title="Fechar">
            <Icon name="close" size={15} />
          </button>
        </div>
        <div className="cats-list">
          {list.map((c) => (
            <div key={c.id}>
              <div className="cat-row">
                <span
                  className="cat-swatch"
                  style={{ background: c.color }}
                  title="Cor da categoria — clique para trocar"
                  onClick={(e) => {
                    e.stopPropagation();
                    setColorEdit(colorEdit === c.id ? null : c.id);
                  }}
                />
                <input
                  className="cat-name"
                  value={c.name}
                  onChange={(e) => upd(c.id, { name: e.target.value })}
                />
                <button
                  className="btn-ghost danger cat-del"
                  title="Excluir categoria"
                  onClick={() => del(c.id)}
                >
                  <Icon name="delete" size={13} />
                </button>
              </div>
              {colorEdit === c.id && (
                <div className="cat-palette" onClick={(e) => e.stopPropagation()}>
                  {PALETTE.map((p) => (
                    <button
                      key={p}
                      className={`pal-dot${p === c.color ? ' sel' : ''}`}
                      style={{ background: p }}
                      title={p}
                      onClick={() => {
                        upd(c.id, { color: p });
                        setColorEdit(null);
                      }}
                    />
                  ))}
                </div>
              )}
            </div>
          ))}
          {list.length === 0 && (
            <div className="cat-empty">Nenhuma categoria — clique em "Nova" para criar.</div>
          )}
        </div>
        <div className="cats-actions">
          <button className="btn" onClick={addCat}>
            <Icon name="new" size={13} /> Nova
          </button>
          <span style={{ flex: 1 }} />
          <button
            className="btn btn-primary"
            onClick={() =>
              onSave(list.filter((c) => c.name.trim()).map((c) => ({ ...c, name: c.name.trim() })))
            }
          >
            OK
          </button>
          <button className="btn" onClick={onClose}>
            Cancelar
          </button>
        </div>
      </div>
    </div>
  );
}
