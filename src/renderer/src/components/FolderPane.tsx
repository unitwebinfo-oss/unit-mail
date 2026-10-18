import React, { useState } from 'react';
import { Icon } from '../icons';
import { folderIconKey, containerClassForModule } from '../utils';
import { ITEM_DRAG_MIME } from './MessageList';
import { isDefaultFolderId } from '../../../shared/types';
import type { FolderNode, MailSource, ModuleId } from '../../../shared/types';

export interface ItemDragPayload {
  sourceId: string;
  folderId: string;
  index: number;
  /** Multi-seleção arrastada junta; ausente = só `index`. */
  indexes?: number[];
}

export interface FolderDragPayload {
  kind: 'folder';
  sourceId: string;
  folderId: string;
  /** Id do pai (string vazia = raiz da base). */
  parentId: string;
}

export const FOLDER_DRAG_MIME = 'application/x-outlook-folder';

/** Cabeçalho de base (arquivo de dados) arrastado para reordenar as bases. */
export interface SourceDragPayload {
  kind: 'source';
  sourceId: string;
}
export const SOURCE_DRAG_MIME = 'application/x-outlook-source';

interface SourceEntry {
  source: MailSource;
  tree: FolderNode[];
}

interface FolderPaneProps {
  entries: SourceEntry[];
  selected: { sourceId: string; folderId: string } | null;
  module: ModuleId;
  onSelect: (sourceId: string, folder: FolderNode) => void;
  onMoveItem?: (payload: ItemDragPayload, destSourceId: string, destFolder: FolderNode) => void;
  /** Pasta arrastada para dentro de outra pasta (ou para a raiz com newParentId=null). */
  onMoveFolder?: (sourceId: string, folderId: string, newParentId: string | null) => void;
  /** Pasta reordenada entre irmãs (position: soltar antes/depois do alvo). */
  onReorderFolder?: (
    sourceId: string,
    parentId: string,
    folderId: string,
    targetId: string,
    after: boolean,
  ) => void;
  /** Base reordenada entre as demais (soltar antes/depois do cabeçalho-alvo). */
  onReorderSource?: (dragId: string, targetId: string, after: boolean) => void;
  onCreateFolder?: (sourceId: string, parentId: string | null, name: string) => void;
  onRenameFolder?: (sourceId: string, folderId: string, name: string) => void;
  onDeleteFolder?: (sourceId: string, folderId: string) => void;
  onCopyFolder?: (sourceId: string, folderId: string) => void;
  onEmptyFolder?: (sourceId: string, folderId: string) => void;
  onMarkAllRead?: (sourceId: string, folderId: string) => void;
  /** Classificar subpastas de `parentId` ('' = raiz da base) em ordem alfabética. */
  onSortChildren?: (sourceId: string, parentId: string) => void;
  /** Mover a pasta uma posição para cima/baixo entre as irmãs. */
  onMoveFolderStep?: (sourceId: string, folderId: string, dir: 'up' | 'down') => void;
  /** Alterna a pasta nos Favoritos. */
  onToggleFavorite?: (sourceId: string, folderId: string) => void;
  /** Oculta/reexibe a pasta na árvore (não apaga conteúdo). */
  onSetFolderHidden?: (sourceId: string, folderId: string, hidden: boolean) => void;
  /** Ids "sourceId|folderId" atualmente nos Favoritos. */
  favorites?: Set<string>;
}

interface CtxState {
  x: number;
  y: number;
  sourceId: string;
  parentId: string;
  node: FolderNode;
  mode: 'menu' | 'new' | 'rename' | 'move';
}

function FolderItem({
  node,
  depth,
  sourceId,
  parentId,
  selected,
  showHidden,
  onSelect,
  onMoveItem,
  onMoveFolder,
  onReorderFolder,
  onCtx,
}: {
  node: FolderNode;
  depth: number;
  sourceId: string;
  parentId: string;
  selected: { sourceId: string; folderId: string } | null;
  /** true = pastas ocultas aparecem esmaecidas (permite reexibir) */
  showHidden: boolean;
  onSelect: (sourceId: string, f: FolderNode) => void;
  onMoveItem?: (payload: ItemDragPayload, destSourceId: string, destFolder: FolderNode) => void;
  onMoveFolder?: (sourceId: string, folderId: string, newParentId: string | null) => void;
  onReorderFolder?: (
    sourceId: string,
    parentId: string,
    folderId: string,
    targetId: string,
    after: boolean,
  ) => void;
  onCtx?: (e: React.MouseEvent, sourceId: string, parentId: string, node: FolderNode) => void;
}) {
  const expKey = `folder.expanded.${sourceId}.${node.id}`;
  const [expanded, setExpanded] = useState(() => {
    const v = localStorage.getItem(expKey);
    return v === null ? depth < 2 : v === '1';
  });
  const [dropHover, setDropHover] = useState(false);
  // Pasta ocultada pelo usuário — some da árvore a menos que "Mostrar
  // pastas ocultas" esteja ligado (aí aparece esmaecida para reexibir).
  if (node.hidden && !showHidden) return null;
  const hasChildren = node.children.length > 0;
  const icon = folderIconKey(node.name, node.containerClass);
  const isSelected = selected?.sourceId === sourceId && selected?.folderId === node.id;

  const handleDrop = (e: React.DragEvent) => {
    setDropHover(false);
    const types = e.dataTransfer.types;
    if (!types.includes(ITEM_DRAG_MIME) && !types.includes(FOLDER_DRAG_MIME)) return;
    e.preventDefault();
    e.stopPropagation();

    // Item de email arrastado sobre a pasta
    const itemRaw = e.dataTransfer.getData(ITEM_DRAG_MIME);
    if (itemRaw) {
      try {
        const payload = JSON.parse(itemRaw) as ItemDragPayload;
        if (payload.sourceId === sourceId && payload.folderId === node.id) return;
        onMoveItem?.(payload, sourceId, node);
      } catch {
        // payload inválido
      }
      return;
    }

    // Pasta arrastada: mesmo pai → reordena; pai diferente → move para dentro
    const folderRaw = e.dataTransfer.getData(FOLDER_DRAG_MIME);
    if (!folderRaw) return;
    try {
      const payload = JSON.parse(folderRaw) as FolderDragPayload;
      if (payload.folderId === node.id) return;
      const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
      const after = e.clientY > rect.top + rect.height / 2;
      if (payload.sourceId === sourceId && payload.parentId === parentId) {
        onReorderFolder?.(sourceId, parentId, payload.folderId, node.id, after);
      } else {
        onMoveFolder?.(payload.sourceId, payload.folderId, node.id);
      }
    } catch {
      // payload inválido
    }
  };

  return (
    <>
      <div
        className={`folder-item ${isSelected ? 'selected' : ''} ${dropHover ? 'drop-hover' : ''} ${node.hidden ? 'f-hidden' : ''}`}
        style={{ paddingLeft: 8 + depth * 14 }}
        onClick={() => onSelect(sourceId, node)}
        title={node.name}
        draggable
        onDragStart={(e) => {
          e.stopPropagation();
          e.dataTransfer.setData(
            FOLDER_DRAG_MIME,
            JSON.stringify({
              kind: 'folder',
              sourceId,
              folderId: node.id,
              parentId,
            } satisfies FolderDragPayload),
          );
          e.dataTransfer.effectAllowed = 'move';
        }}
        onDragOver={(e) => {
          if (
            e.dataTransfer.types.includes(ITEM_DRAG_MIME) ||
            e.dataTransfer.types.includes(FOLDER_DRAG_MIME)
          ) {
            e.preventDefault();
            e.stopPropagation();
            e.dataTransfer.dropEffect = 'move';
            setDropHover(true);
          }
        }}
        onDragLeave={() => setDropHover(false)}
        onDrop={handleDrop}
        onContextMenu={(e) => {
          e.preventDefault();
          e.stopPropagation();
          onCtx?.(e, sourceId, parentId, node);
        }}
      >
        <span
          className="chevron"
          onClick={(e) => {
            e.stopPropagation();
            if (hasChildren) {
              const next = !expanded;
              setExpanded(next);
              localStorage.setItem(expKey, next ? '1' : '0');
            }
          }}
        >
          {hasChildren ? (expanded ? '−' : '+') : null}
        </span>
        <span className="f-icon">
          <Icon name={icon} size={16} />
        </span>
        <span className="f-name">{node.name}</span>
        {node.unreadCount > 0 && <span className="unread-badge">{node.unreadCount}</span>}
      </div>
      {expanded &&
        node.children.map((c) => (
          <FolderItem
            key={c.id}
            node={c}
            depth={depth + 1}
            sourceId={sourceId}
            parentId={node.id}
            selected={selected}
            showHidden={showHidden}
            onSelect={onSelect}
            onMoveItem={onMoveItem}
            onMoveFolder={onMoveFolder}
            onReorderFolder={onReorderFolder}
            onCtx={onCtx}
          />
        ))}
    </>
  );
}

function filterTree(nodes: FolderNode[], cc: string): FolderNode[] {
  const out: FolderNode[] = [];
  for (const n of nodes) {
    const kids = filterTree(n.children, cc);
    if (n.containerClass === cc || kids.length > 0) {
      out.push({ ...n, children: kids });
    }
  }
  return out;
}

function findNode(nodes: FolderNode[], id: string): FolderNode | null {
  for (const n of nodes) {
    if (n.id === id) return n;
    const f = findNode(n.children, id);
    if (f) return f;
  }
  return null;
}

export function FolderPane({
  entries,
  selected,
  module,
  onSelect,
  onMoveItem,
  onMoveFolder,
  onReorderFolder,
  onReorderSource,
  onCreateFolder,
  onRenameFolder,
  onDeleteFolder,
  onCopyFolder,
  onEmptyFolder,
  onMarkAllRead,
  onSortChildren,
  onMoveFolderStep,
  onToggleFavorite,
  onSetFolderHidden,
  favorites,
}: FolderPaneProps) {
  const cc = containerClassForModule(module);
  const [ctx, setCtx] = useState<CtxState | null>(null);
  const [ctxValue, setCtxValue] = useState('');
  /** Bases com "mostrar pastas ocultas" ligado — id da fonte. */
  const [showHidden, setShowHidden] = useState<Set<string>>(new Set());
  /** Cabeçalho de base sob o cursor durante um drag — feedback visual. */
  const [srcDropHover, setSrcDropHover] = useState<string | null>(null);
  /** Expansão do cabeçalho de cada base — persistida como no Outlook. */
  const [srcCollapsed, setSrcCollapsed] = useState<Record<string, boolean>>({});
  const isSrcOpen = (id: string) => {
    if (id in srcCollapsed) return !srcCollapsed[id];
    return localStorage.getItem(`source.expanded.${id}`) !== '0';
  };
  const toggleSrc = (id: string) => {
    const open = isSrcOpen(id);
    localStorage.setItem(`source.expanded.${id}`, open ? '0' : '1');
    setSrcCollapsed((c) => ({ ...c, [id]: open }));
  };
  const toggleHiddenView = (sourceId: string) =>
    setShowHidden((s) => {
      const n = new Set(s);
      if (n.has(sourceId)) n.delete(sourceId);
      else n.add(sourceId);
      return n;
    });

  const visibleEntries = entries
    .map((e) => ({
      ...e,
      tree: module === 'mail' ? e.tree : e.source.kind === 'imap' ? [] : filterTree(e.tree, cc),
    }))
    .filter((e) => e.tree.length > 0 || (module === 'mail' && e.source.kind === 'imap'));

  const openCtx = (e: React.MouseEvent, sourceId: string, parentId: string, node: FolderNode) => {
    setCtx({ x: e.clientX, y: e.clientY, sourceId, parentId, node, mode: 'menu' });
    setCtxValue(node.name);
  };

  // Fecha o menu ao clicar fora / pressionar Esc
  React.useEffect(() => {
    if (!ctx) return;
    const close = () => setCtx(null);
    const key = (e: KeyboardEvent) => e.key === 'Escape' && close();
    window.addEventListener('click', close);
    window.addEventListener('keydown', key);
    return () => {
      window.removeEventListener('click', close);
      window.removeEventListener('keydown', key);
    };
  }, [ctx]);

  /** Pastas favoritas resolvidas (nó completo + base) na ordem salva. */
  const favNodes: Array<{ sourceId: string; node: FolderNode }> = [];
  if (favorites) {
    for (const key of favorites) {
      const sep = key.indexOf('|');
      const sid = key.slice(0, sep);
      const fid = key.slice(sep + 1);
      const entry = entries.find((e) => e.source.id === sid);
      const node = entry ? findNode(entry.tree, fid) : null;
      if (entry && node) favNodes.push({ sourceId: sid, node });
    }
  }

  /** Lista achatada de pastas da base do menu — destinos para "Mover Pasta". */
  const moveTargets: Array<{ id: string | null; label: string }> = [];
  if (ctx?.mode === 'move') {
    moveTargets.push({ id: null, label: '（raiz da base）' });
    const entry = entries.find((e) => e.source.id === ctx.sourceId);
    const flat = (nodes: FolderNode[], depth: number) => {
      for (const n of nodes) {
        if (n.id === ctx.node.id || n.id.startsWith(`${ctx.node.id}/`)) continue;
        moveTargets.push({ id: n.id, label: `${'　'.repeat(depth)}${n.name}` });
        flat(n.children, depth + 1);
      }
    };
    if (entry) flat(entry.tree, 0);
  }

  const menuStyle: React.CSSProperties = {
    position: 'fixed',
    left: Math.min(ctx?.x ?? 0, window.innerWidth - 240),
    top: Math.min(ctx?.y ?? 0, window.innerHeight - 380),
    zIndex: 120,
  };

  const item = (
    label: string,
    icon: string,
    onClick: () => void,
    disabled = false,
  ) => (
    <button
      key={label}
      disabled={disabled}
      onClick={() => {
        setCtx(null);
        onClick();
      }}
    >
      <Icon name={icon} size={14} /> {label}
    </button>
  );

  const isLocal = ctx?.node.id.startsWith('arc:') ?? false;
  // Pastas padrão (Caixa de Entrada, Itens Excluídos, Calendário…) são
  // protegidas — excluir/renomear/mover quebraria fluxos internos do app.
  const isProtected = ctx ? isDefaultFolderId(ctx.node.id) : false;
  const isFav = ctx ? favorites?.has(`${ctx.sourceId}|${ctx.node.id}`) : false;

  return (
    <div className="folder-pane">
      <div className="folder-tree">
        {favNodes.length > 0 && (
          <div style={{ marginBottom: 6 }}>
            <div className="folder-item fav-head">
              <span className="f-icon">
                <Icon name="flag" size={16} />
              </span>
              <span className="f-name">Favoritos</span>
            </div>
            {favNodes.map((f) => (
              <div
                key={`${f.sourceId}|${f.node.id}`}
                className={`folder-item ${
                  selected?.sourceId === f.sourceId && selected?.folderId === f.node.id
                    ? 'selected'
                    : ''
                }`}
                style={{ paddingLeft: 22 }}
                onClick={() => onSelect(f.sourceId, f.node)}
                title={f.node.name}
              >
                <span className="chevron" />
                <span className="f-icon">
                  <Icon name={folderIconKey(f.node.name, f.node.containerClass)} size={16} />
                </span>
                <span className="f-name">{f.node.name}</span>
                {f.node.unreadCount > 0 && <span className="unread-badge">{f.node.unreadCount}</span>}
              </div>
            ))}
          </div>
        )}
        {visibleEntries.length === 0 && <div className="empty-list">Nenhuma base aberta</div>}
        {visibleEntries.map((e) => (
          <div
            key={e.source.id}
            style={{ marginBottom: 6 }}
            // O bloco inteiro da base (cabeçalho + árvore) é alvo de drop de
            // outra base — soltar na metade de cima insere antes, embaixo depois.
            // Bases NUNCA viram subpasta de outra base: o drop só reordena.
            onDragOver={(ev) => {
              if (ev.dataTransfer.types.includes(SOURCE_DRAG_MIME)) {
                ev.preventDefault();
                ev.dataTransfer.dropEffect = 'move';
              }
            }}
            onDrop={(ev) => {
              const srcRaw = ev.dataTransfer.getData(SOURCE_DRAG_MIME);
              if (!srcRaw) return; // drop de pasta/item segue no cabeçalho abaixo
              ev.preventDefault();
              ev.stopPropagation();
              setSrcDropHover(null);
              try {
                const payload = JSON.parse(srcRaw) as SourceDragPayload;
                if (payload.sourceId === e.source.id) return;
                const rect = (ev.currentTarget as HTMLElement).getBoundingClientRect();
                const after = ev.clientY > rect.top + rect.height / 2;
                onReorderSource?.(payload.sourceId, e.source.id, after);
              } catch {
                // payload inválido
              }
            }}
          >
            <div
              className={`folder-item ${srcDropHover === e.source.id ? 'drop-hover' : ''}`}
              style={{ fontWeight: 700, cursor: 'default' }}
              draggable={!!onReorderSource}
              onDragStart={(ev) => {
                const json = JSON.stringify({
                  kind: 'source',
                  sourceId: e.source.id,
                } satisfies SourceDragPayload);
                ev.dataTransfer.setData(SOURCE_DRAG_MIME, json);
                // Alguns runtimes só iniciam o drag se houver um tipo padrão
                ev.dataTransfer.setData('text/plain', e.source.label);
                ev.dataTransfer.effectAllowed = 'move';
              }}
              onDragEnd={() => setSrcDropHover(null)}
              onDragEnter={(ev) => {
                if (ev.dataTransfer.types.includes(SOURCE_DRAG_MIME)) {
                  setSrcDropHover(e.source.id);
                }
              }}
              onDragOver={(ev) => {
                if (
                  ev.dataTransfer.types.includes(FOLDER_DRAG_MIME) ||
                  ev.dataTransfer.types.includes(SOURCE_DRAG_MIME)
                ) {
                  ev.preventDefault();
                  ev.stopPropagation();
                  ev.dataTransfer.dropEffect = 'move';
                }
              }}
              onDragLeave={() => setSrcDropHover((h) => (h === e.source.id ? null : h))}
              onDrop={(ev) => {
                setSrcDropHover(null);
                // Base arrastada sobre o cabeçalho → reordena (metade de cima
                // = antes, de baixo = depois). O wrapper também trata isso —
                // aqui o stopPropagation evita contar a posição duas vezes.
                const srcRaw = ev.dataTransfer.getData(SOURCE_DRAG_MIME);
                if (srcRaw) {
                  ev.preventDefault();
                  ev.stopPropagation();
                  try {
                    const payload = JSON.parse(srcRaw) as SourceDragPayload;
                    if (payload.sourceId === e.source.id) return;
                    const rect = (ev.currentTarget as HTMLElement).getBoundingClientRect();
                    const after = ev.clientY > rect.top + rect.height / 2;
                    onReorderSource?.(payload.sourceId, e.source.id, after);
                  } catch {
                    // payload inválido
                  }
                  return;
                }
                // Pasta arrastada sobre o cabeçalho → move para a raiz da base
                const raw = ev.dataTransfer.getData(FOLDER_DRAG_MIME);
                if (!raw) return;
                ev.preventDefault();
                try {
                  const payload = JSON.parse(raw) as FolderDragPayload;
                  if (payload.sourceId === e.source.id && payload.parentId !== '') {
                    onMoveFolder?.(e.source.id, payload.folderId, null);
                  }
                } catch {
                  // payload inválido
                }
              }}
            >
              <span
                className="chevron"
                style={{ cursor: 'pointer' }}
                title={isSrcOpen(e.source.id) ? 'Recolher base' : 'Expandir base'}
                onClick={(ev) => {
                  ev.stopPropagation();
                  toggleSrc(e.source.id);
                }}
              >
                {e.tree.length > 0 ? (isSrcOpen(e.source.id) ? '−' : '+') : null}
              </span>
              <span className="f-icon">
                <Icon name={e.source.kind === 'imap' ? 'envelope' : 'archive'} size={16} />
              </span>
              <span className="f-name" title={e.source.label}>
                {e.source.label}
              </span>
              {e.source.connecting && <span style={{ marginLeft: 6, fontSize: 10 }}>conectando…</span>}
              {e.source.error && (
                <span style={{ marginLeft: 6, fontSize: 10, color: '#c00' }} title={e.source.error}>
                  erro
                </span>
              )}
              {e.source.kind !== 'imap' && (
                <span
                  style={{
                    marginLeft: 'auto',
                    cursor: 'pointer',
                    color: showHidden.has(e.source.id) ? 'var(--blue)' : 'var(--text-dim)',
                  }}
                  title={
                    showHidden.has(e.source.id)
                      ? 'Ocultar pastas escondidas'
                      : 'Mostrar pastas ocultas'
                  }
                  onClick={(ev) => {
                    ev.stopPropagation();
                    toggleHiddenView(e.source.id);
                  }}
                >
                  <Icon name={showHidden.has(e.source.id) ? 'eyeOff' : 'eye'} size={13} />
                </span>
              )}
              {e.source.kind !== 'imap' && onCreateFolder && (
                <span
                  style={{ cursor: 'pointer', color: 'var(--text-dim)', marginLeft: 8 }}
                  title="Nova pasta nesta base"
                  onClick={(ev) => {
                    ev.stopPropagation();
                    setCtxValue('');
                    setCtx({
                      x: ev.clientX,
                      y: ev.clientY,
                      sourceId: e.source.id,
                      parentId: '',
                      node: {
                        id: '',
                        name: e.source.label,
                        containerClass: '',
                        contentCount: 0,
                        unreadCount: 0,
                        children: [],
                      },
                      mode: 'new',
                    });
                  }}
                >
                  <Icon name="new" size={13} />
                </span>
              )}
            </div>
            {isSrcOpen(e.source.id) &&
              e.tree.map((n) => (
              <FolderItem
                key={n.id}
                node={n}
                depth={1}
                sourceId={e.source.id}
                parentId=""
                selected={selected}
                showHidden={showHidden.has(e.source.id)}
                onSelect={onSelect}
                onMoveItem={onMoveItem}
                onMoveFolder={onMoveFolder}
                onReorderFolder={onReorderFolder}
                onCtx={openCtx}
              />
            ))}
          </div>
        ))}
      </div>

      {/* Menu de contexto da pasta (botão direito) — estilo Outlook */}
      {ctx && (
        <div
          className="ctx-menu"
          style={menuStyle}
          onClick={(e) => e.stopPropagation()}
          onContextMenu={(e) => e.preventDefault()}
        >
          {ctx.mode === 'menu' && (
            <>
              {item('Nova Pasta…', 'new', () => {
                // Campo vazio — "nova pasta" não deve pré-preencher o nome da pasta-pai
                setCtxValue('');
                setCtx({ ...ctx, mode: 'new' });
              })}
              {item('Renomear Pasta', 'draft', () => setCtx({ ...ctx, mode: 'rename' }), !isLocal || isProtected)}
              {item('Copiar Pasta', 'folder', () => onCopyFolder?.(ctx.sourceId, ctx.node.id), !isLocal)}
              {item('Mover Pasta…', 'open', () => setCtx({ ...ctx, mode: 'move' }), !isLocal || isProtected)}
              {/* A confirmação "Tem certeza?" é do próprio app (ConfirmDialog
                  no App) — window.confirm abriria um segundo diálogo nativo. */}
              {item('Excluir Pasta', 'delete', () =>
                onDeleteFolder?.(ctx.sourceId, ctx.node.id), !isLocal || isProtected)}
              {ctx.node.id !== '' &&
                item(
                  ctx.node.hidden ? 'Reexibir Pasta' : 'Ocultar Pasta',
                  ctx.node.hidden ? 'eye' : 'eyeOff',
                  () =>
                    onSetFolderHidden?.(ctx.sourceId, ctx.node.id, !ctx.node.hidden),
                  !onSetFolderHidden,
                )}
              <div className="ctx-sep" />
              {item('Marcar Todas como Lidas', 'envelope', () =>
                onMarkAllRead?.(ctx.sourceId, ctx.node.id),
              )}
              {item('Limpar Pasta', 'refresh', () =>
                onEmptyFolder?.(ctx.sourceId, ctx.node.id),
              )}
              {item('Excluir Tudo', 'trash', () =>
                onEmptyFolder?.(ctx.sourceId, ctx.node.id),
              )}
              <div className="ctx-sep" />
              {item(
                isFav ? 'Remover dos Favoritos' : 'Adicionar a Favoritos',
                'flag',
                () => onToggleFavorite?.(ctx.sourceId, ctx.node.id),
              )}
              {item('Classificar Subpastas A a Z', 'chevronD', () =>
                onSortChildren?.(ctx.sourceId, ctx.node.id),
              )}
              <div className="ctx-sep" />
              {item('Mover para cima', 'chevronR', () =>
                onMoveFolderStep?.(ctx.sourceId, ctx.node.id, 'up'),
              )}
              {item('Mover para Baixo', 'chevronD', () =>
                onMoveFolderStep?.(ctx.sourceId, ctx.node.id, 'down'),
              )}
            </>
          )}

          {(ctx.mode === 'new' || ctx.mode === 'rename') && (
            <div className="ctx-input">
              <label>{ctx.mode === 'new' ? 'Nome da nova pasta' : 'Novo nome'}</label>
              <input
                autoFocus
                value={ctxValue}
                onChange={(e) => setCtxValue(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && ctxValue.trim()) {
                    const v = ctxValue.trim();
                    setCtx(null);
                    if (ctx.mode === 'new') {
                      onCreateFolder?.(ctx.sourceId, ctx.node.id, v);
                    } else {
                      onRenameFolder?.(ctx.sourceId, ctx.node.id, v);
                    }
                  }
                  if (e.key === 'Escape') setCtx(null);
                }}
              />
              <div className="ctx-input-actions">
                <button onClick={() => setCtx(null)}>Cancelar</button>
                <button
                  className="primary"
                  onClick={() => {
                    if (!ctxValue.trim()) return;
                    const v = ctxValue.trim();
                    setCtx(null);
                    if (ctx.mode === 'new') {
                      onCreateFolder?.(ctx.sourceId, ctx.node.id, v);
                    } else {
                      onRenameFolder?.(ctx.sourceId, ctx.node.id, v);
                    }
                  }}
                >
                  OK
                </button>
              </div>
            </div>
          )}

          {ctx.mode === 'move' && (
            <div className="ctx-move">
              <div className="ctx-move-title">Mover "{ctx.node.name}" para:</div>
              {moveTargets.map((t) => (
                <button
                  key={t.id ?? '__root'}
                  onClick={() => {
                    setCtx(null);
                    onMoveFolder?.(ctx.sourceId, ctx.node.id, t.id);
                  }}
                >
                  {t.label}
                </button>
              ))}
              <div className="ctx-sep" />
              <button onClick={() => setCtx(null)}>Cancelar</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
