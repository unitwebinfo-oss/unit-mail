import React from 'react';
import { Icon } from '../icons';
import { fmtSmartDate } from '../utils';
import { normSearch } from '../../../shared/textNorm';
import { t, useLang } from '../i18n';
import type { ItemSummary, MailCategory } from '../../../shared/types';

export interface GlobalSearchHit {
  sourceId: string;
  sourceLabel: string;
  folderId: string;
  folderName: string;
  item: ItemSummary;
}

interface MessageListProps {
  items: ItemSummary[];
  total: number;
  selectedIndex: number | null;
  /** Índices multi-selecionados (Ctrl/Shift-clique) dentro da pasta atual. */
  selIndexes: ReadonlySet<number>;
  searchResults: GlobalSearchHit[] | null;
  /** Busca em andamento no backend — indicador discreto na barra (live search). */
  searchBusy?: boolean;
  /** Identificação da pasta atual (habilita arrastar itens para outra pasta). */
  dragRef?: { sourceId: string; folderId: string } | null;
  onSearch: (q: string) => void;
  onSelect: (item: ItemSummary, sourceId?: string, folderId?: string) => void;
  /** Duplo clique na linha → abre a mensagem em janela própria (como o Outlook). */
  onOpenItem?: (item: ItemSummary, sourceId?: string, folderId?: string) => void;
  /** Atualiza a multi-seleção; `primary` é o item clicado (mostrado no painel). */
  onSelChange: (indexes: number[], primary: number) => void;
  /** Tecla Delete sobre a lista → exclui os selecionados. */
  onDeleteKey?: () => void;
  /** Pastas de destino do submenu "Mover" para a base do item clicado. */
  moveTargets?: (sourceId: string) => Array<{ id: string; label: string }>;
  /** Ação do menu de contexto da mensagem (botão direito). */
  onCtxAction?: (action: string, ctx: MsgCtxPayload) => void;
  onLoadMore: () => void;
  hasMore: boolean;
  /** Carregando itens da pasta — overlay sutil sobre a lista. */
  loading?: boolean;
  /** Pastas de enviados (Itens Enviados/Caixa de Saída) não têm filtro "Não lidos". */
  noUnreadFilter?: boolean;
  /** Identidade da pasta listada (ex.: `sourceId|folderId`) — o filtro "Não lidos" recongela ao mudar. */
  listKey?: string;
  /** Recebe a ordem real exibida (agrupada/plana/filtrada) — o App usa para
   *  achar o "próximo" item visualmente ao excluir. */
  displayedRef?: { current: ItemSummary[] };
  /** Lista-mestre de categorias — selos coloridos nas linhas + submenu Categorizar. */
  categories?: MailCategory[];
}

export const ITEM_DRAG_MIME = 'application/x-outlook-item';

/** Dados da mensagem clicada com o botão direito. */
export interface MsgCtxPayload {
  sourceId: string;
  folderId: string;
  index: number;
  indexes: number[];
  isRead: boolean;
  /** Categorias do item clicado (para marcar os ✓ no submenu). */
  cats: string[];
  /** E-mails dos remetentes dos itens selecionados (Bloquear Remetente) —
   *  resolvidos aqui porque o App pode não ter os itens (ex.: busca global). */
  senders: string[];
}

interface MsgCtx extends MsgCtxPayload {
  x: number;
  y: number;
  mode: 'menu' | 'move' | 'cats';
}

/** Agrupamento por data como o Outlook: Hoje, Ontem, Esta semana, etc. */
function dateBucket(iso: string | null): string {
  if (!iso) return 'Sem data';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return 'Sem data';
  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const startOfDay = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const diff = startOfToday - startOfDay;
  if (diff <= 0) return 'Hoje';
  if (diff === 86400000) return 'Ontem';
  if (diff < 7 * 86400000) return 'Esta semana';
  if (diff < 14 * 86400000) return 'Semana passada';
  if (d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth()) return 'Este mês';
  if (d.getFullYear() === now.getFullYear()) {
    const label = d.toLocaleDateString('pt-BR', { month: 'long' });
    return label.charAt(0).toUpperCase() + label.slice(1);
  }
  return String(d.getFullYear());
}

const BUCKET_ORDER = [
  'Hoje',
  'Ontem',
  'Esta semana',
  'Semana passada',
  'Este mês',
  'Sem data',
];

function bucketRank(label: string): number {
  const i = BUCKET_ORDER.indexOf(label);
  return i === -1 ? BUCKET_ORDER.length : i;
}

/**
 * Realce dos termos pesquisados (estilo Outlook/Word): envolve cada
 * ocorrência em <mark>. A comparação usa o texto normalizado (minúsculas,
 * sem acentos) e mapeia de volta aos índices do texto original.
 */
function highlightText(text: string, tokens: string[]): React.ReactNode {
  if (!text || !tokens.length) return text;
  // Texto normalizado + mapa norm→original por caractere
  const normChars: string[] = [];
  const map: number[] = [];
  let i = 0;
  for (const ch of text) {
    for (const c of ch.toLowerCase().normalize('NFD')) {
      if (/\p{M}/u.test(c)) continue;
      normChars.push(c);
      map.push(i);
    }
    i += ch.length;
  }
  const norm = normChars.join('');
  const ranges: Array<[number, number]> = [];
  for (const t of tokens) {
    let from = 0;
    for (;;) {
      const j = norm.indexOf(t, from);
      if (j < 0) break;
      ranges.push([map[j], map[j + t.length - 1] + 1]);
      from = j + t.length;
    }
  }
  if (!ranges.length) return text;
  ranges.sort((a, b) => a[0] - b[0]);
  const merged: Array<[number, number]> = [];
  for (const [a, b] of ranges) {
    const last = merged[merged.length - 1];
    if (last && a <= last[1]) last[1] = Math.max(last[1], b);
    else merged.push([a, b]);
  }
  const out: React.ReactNode[] = [];
  let pos = 0;
  let k = 0;
  for (const [a, b] of merged) {
    if (a > pos) out.push(text.slice(pos, a));
    out.push(<mark key={k++}>{text.slice(a, b)}</mark>);
    pos = b;
  }
  if (pos < text.length) out.push(text.slice(pos));
  return out;
}

function fmtSize(size?: number | { toNumber?: () => number } | unknown): string {
  // size pode chegar como objeto Long do PST — coage para número primeiro
  const n =
    typeof size === 'number'
      ? size
      : typeof (size as { toNumber?: () => number })?.toNumber === 'function'
        ? (size as { toNumber: () => number }).toNumber()
        : Number(size);
  if (!Number.isFinite(n) || n <= 0) return '';
  if (n < 1024) return `${Math.round(n)} B`;
  const kb = n / 1024;
  if (kb < 1024) return `${Math.round(kb)} KB`;
  return `${(kb / 1024).toFixed(1)} MB`;
}

function Row({
  m,
  selected,
  dragPayload,
  onSelect,
  onOpen,
  onCtx,
  catById,
  hl,
}: {
  m: ItemSummary;
  selected: boolean;
  /** Índices arrastados juntos (quando a linha faz parte de multi-seleção). */
  dragPayload?: { sourceId: string; folderId: string; indexes: number[] } | null;
  onSelect: (e: React.MouseEvent) => void;
  onOpen?: () => void;
  onCtx?: (e: React.MouseEvent) => void;
  catById?: Map<string, MailCategory>;
  /** Termos da busca para realçar no remetente/assunto (já normalizados). */
  hl?: string[];
}) {
  return (
    <div
      className={`msg-line ${!m.isRead ? 'unread' : ''} ${selected ? 'selected' : ''}`}
      draggable={!!dragPayload}
      onDragStart={(e) => {
        if (!dragPayload) return;
        e.dataTransfer.setData(
          ITEM_DRAG_MIME,
          JSON.stringify({ ...dragPayload, index: m.index }),
        );
        e.dataTransfer.effectAllowed = 'move';
      }}
      onClick={onSelect}
      onDoubleClick={onOpen}
      onContextMenu={onCtx}
      title={`${m.senderName || m.senderEmailAddress} — ${m.subject}`}
    >
      <span className="ml-icons">
        {m.inviteMethod && <Icon name="meeting" size={15} color="#0f6cbd" />}
        {m.hasAttachments && <Icon name="clip" size={15} color="#ca5010" />}
        {m.importance === 2 && <Icon name="flag" size={15} color="#d13438" />}
        {(m.categories ?? []).map((id) => {
          const c = catById?.get(id);
          return c ? (
            <span
              key={id}
              className="cat-dot"
              style={{ background: c.color }}
              title={c.name}
            />
          ) : null;
        })}
      </span>
      <span className="ml-from">
        {highlightText(m.senderName || m.senderEmailAddress || '(desconhecido)', hl ?? [])}
      </span>
      <span className="ml-subject">{highlightText(m.subject, hl ?? [])}</span>
      <span className="ml-date">{fmtSmartDate(m.deliveryTime || m.sentTime)}</span>
      <span className="ml-size">{fmtSize(m.size)}</span>
    </div>
  );
}

type SortKey = 'from' | 'subject' | 'date' | 'size';

interface ColWidths {
  from: number;
  date: number;
  size: number;
}

const COLS_KEY = 'ml.columns';
function loadCols(): ColWidths {
  try {
    const s = JSON.parse(localStorage.getItem(COLS_KEY) || '{}');
    return {
      from: typeof s.from === 'number' ? s.from : 200,
      date: typeof s.date === 'number' ? s.date : 140,
      size: typeof s.size === 'number' ? s.size : 70,
    };
  } catch {
    return { from: 200, date: 140, size: 70 };
  }
}

export function MessageList({
  items,
  selectedIndex,
  selIndexes,
  searchResults,
  searchBusy,
  dragRef,
  onSearch,
  onSelect,
  onOpenItem,
  onSelChange,
  onDeleteKey,
  moveTargets,
  onCtxAction,
  onLoadMore,
  hasMore,
  loading,
  noUnreadFilter,
  listKey,
  displayedRef,
  categories,
}: MessageListProps) {
  useLang();
  const searching = searchResults !== null;
  // Texto da pesquisa — escopo do filtro de não-lidos durante a busca
  const [query, setQuery] = React.useState('');
  // Termos normalizados para realçar nas linhas de resultado
  const hlTokens = React.useMemo(
    () => normSearch(query).split(/\s+/).filter(Boolean),
    [query],
  );
  // Âncora do Shift-clique (posição na ordem exibida)
  const anchorRef = React.useRef<number | null>(null);
  const [msgCtx, setMsgCtx] = React.useState<MsgCtx | null>(null);
  const catById = React.useMemo(
    () => new Map((categories ?? []).map((c) => [c.id, c])),
    [categories],
  );
  // Ordenação por coluna (clique no cabeçalho, como no Outlook)
  const [sortKey, setSortKey] = React.useState<SortKey>('date');
  const [sortDesc, setSortDesc] = React.useState(true);
  // Larguras ajustáveis das colunas (arrastar a borda do cabeçalho)
  const [cols, setCols] = React.useState<ColWidths>(loadCols);
  // Filtro "Todas / Não lidos" — abas acima da lista, como no Outlook
  const [unreadOnly, setUnreadOnly] = React.useState(false);
  // Filtro por categoria — id da categoria, 'none' = sem categoria, null = todas
  const [catFilter, setCatFilter] = React.useState<string | null>(null);
  const paneRef = React.useRef<HTMLDivElement | null>(null);

  // ===== Virtualização da lista (como Outlook/Thunderbird) =====
  // Pastas grandes têm dezenas de milhares de itens — renderizar tudo no DOM
  // trava a rolagem. Só as linhas visíveis (+overscan) entram na tela.
  const listRef = React.useRef<HTMLDivElement | null>(null);
  const [viewport, setViewport] = React.useState({ top: 0, height: 600 });
  const VROW_H = 26; // .msg-line e .msg-group — alturas fixadas em 26px no CSS

  React.useEffect(() => {
    const el = listRef.current;
    if (!el) return;
    const measure = () => setViewport((v) => ({ ...v, height: el.clientHeight }));
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  React.useEffect(() => {
    localStorage.setItem(COLS_KEY, JSON.stringify(cols));
  }, [cols]);

  const startResize = (key: keyof ColWidths) => (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = cols[key];
    const onMove = (ev: MouseEvent) => {
      const w = Math.max(50, Math.min(600, startW + ev.clientX - startX));
      setCols((c) => ({ ...c, [key]: w }));
      paneRef.current?.style.setProperty(`--col-${key}`, `${w}px`);
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const clickSort = (key: SortKey) => {
    if (sortKey === key) setSortDesc((d) => !d);
    else {
      setSortKey(key);
      setSortDesc(key === 'date'); // data começa desc, texto asc
    }
  };

  // Ordena os itens carregados conforme a coluna ativa
  const sorted = React.useMemo(() => {
    const arr = [...items];
    const val = (m: ItemSummary): string | number => {
      switch (sortKey) {
        case 'from':
          return (m.senderName || m.senderEmailAddress || '').toLowerCase();
        case 'subject':
          return (m.subject || '').toLowerCase();
        case 'size':
          return m.size ?? 0;
        default:
          return m.deliveryTime || m.sentTime || '';
      }
    };
    arr.sort((a, b) => {
      const va = val(a);
      const vb = val(b);
      const c =
        typeof va === 'number' && typeof vb === 'number'
          ? va - vb
          : String(va).localeCompare(String(vb), 'pt-BR');
      return sortDesc ? -c : c;
    });
    return arr;
  }, [items, sortKey, sortDesc]);

  // Botão direito numa linha: se ela não está na seleção, vira a selecionada
  const openMsgCtx = (
    e: React.MouseEvent,
    m: ItemSummary,
    sourceId: string,
    folderId: string,
  ) => {
    e.preventDefault();
    e.stopPropagation();
    let indexes = selIndexes.has(m.index) && selIndexes.size > 1 ? [...selIndexes] : [m.index];
    if (!selIndexes.has(m.index)) {
      onSelChange([m.index], m.index);
      onSelect(m, sourceId, folderId);
      indexes = [m.index];
    }
    // Remetentes dos itens afetados — o clicado + os demais da multi-seleção
    // (buscados na lista atual; o item clicado sempre entra).
    const senders = new Set<string>();
    const addSender = (it?: ItemSummary) => {
      const em = it?.senderEmailAddress?.trim();
      if (em) senders.add(em.toLowerCase());
    };
    addSender(m);
    for (const ix of indexes) {
      if (ix !== m.index) addSender(items.find((it) => it.index === ix));
    }
    setMsgCtx({
      x: e.clientX,
      y: e.clientY,
      sourceId,
      folderId,
      index: m.index,
      indexes,
      isRead: m.isRead,
      cats: m.categories ?? [],
      senders: [...senders],
      mode: 'menu',
    });
  };

  // Fecha o menu ao clicar fora / Esc
  React.useEffect(() => {
    if (!msgCtx) return;
    const close = () => setMsgCtx(null);
    const key = (e: KeyboardEvent) => e.key === 'Escape' && close();
    window.addEventListener('click', close);
    window.addEventListener('keydown', key);
    return () => {
      window.removeEventListener('click', close);
      window.removeEventListener('keydown', key);
    };
  }, [msgCtx]);

  // Filtro Não lidos — como o Outlook: ao ativar, "congela" o conjunto de não-lidos.
  // Clicar numa mensagem (marcando-a como lida) NÃO a some da lista — ela só sai
  // ao voltar para "Todas", trocar de pasta ou refazer a busca. Itens não lidos
  // que chegam depois (paginação do scroll infinito) entram no conjunto.
  const unreadFrozen = React.useRef<Set<string>>(new Set());
  const frozenScope = React.useRef<string | null>(null);
  const wasUnreadOnly = React.useRef(false);
  const filterScope = searching ? `s|${query}` : `f|${listKey}`;
  const wantUnread = unreadOnly && !noUnreadFilter;
  if (wantUnread) {
    if (!wasUnreadOnly.current || frozenScope.current !== filterScope) {
      frozenScope.current = filterScope;
      unreadFrozen.current = new Set();
    }
    // Alimenta com os não-lidos atuais — itens novos entram, lidos continuam
    if (searching) {
      for (const r of searchResults!)
        if (!r.item.isRead) unreadFrozen.current.add(`${r.folderId}|${r.item.index}`);
    } else {
      for (const m of items) if (!m.isRead) unreadFrozen.current.add(String(m.index));
    }
  }
  wasUnreadOnly.current = wantUnread;

  // Filtro por categoria — aplicado sobre a lista ordenada e sobre os
  // resultados da pesquisa ('none' = itens sem nenhuma categoria).
  const catMatch = React.useCallback(
    (cats: string[] | undefined) =>
      catFilter === null
        ? true
        : catFilter === 'none'
          ? !(cats ?? []).length
          : (cats ?? []).includes(catFilter),
    [catFilter],
  );
  const visible = React.useMemo(
    () =>
      sorted.filter(
        (m) =>
          (!wantUnread || unreadFrozen.current.has(String(m.index))) && catMatch(m.categories),
      ),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sorted, wantUnread, catMatch],
  );
  const visibleSearch = React.useMemo(
    () =>
      searching
        ? searchResults!.filter(
            (r) =>
              (!wantUnread || unreadFrozen.current.has(`${r.folderId}|${r.item.index}`)) &&
              catMatch(r.item.categories),
          )
        : searchResults,
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [searchResults, searching, wantUnread, catMatch],
  );
  const unreadCount = React.useMemo(
    () => items.filter((m) => !m.isRead).length,
    [items],
  );

  // Filtro "Não lidos": o conjunto precisa cobrir a pasta inteira. Enquanto
  // houver itens não carregados, bombeia páginas — senão só os não-lidos das
  // primeiras páginas apareceriam. Quando o preenchimento de fundo já está
  // rodando, o onLoadMore é ignorado pelo guard do App (sem efeito colateral).
  React.useEffect(() => {
    if ((wantUnread || catFilter !== null) && !searching && hasMore && !loading) onLoadMore();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [wantUnread, catFilter, searching, hasMore, loading, items.length]);

  // Agrupa por faixa de data (como o Outlook agrupa a lista de mensagens).
  // A ordenação da coluna ativa vale dentro de cada faixa.
  // Memoizado: a lista re-renderiza a cada evento de rolagem; sem memo cada
  // render refazia faixas e datas (um Date por item) de toda a pasta.
  const { groupList, flat } = React.useMemo(() => {
    const groups = new Map<string, ItemSummary[]>();
    for (const m of visible) {
      const b = dateBucket(m.deliveryTime || m.sentTime);
      const g = groups.get(b) || [];
      g.push(m);
      groups.set(b, g);
    }
    // A ordem das faixas segue a ordenação por data: desc → recentes no topo,
    // asc → antigas no topo. Ordenando por outra coluna a lista fica plana
    // (sem cabeçalhos de data) — como o Outlook faz.
    const gl = [...groups.entries()].sort(
      (a, b) =>
        (sortDesc ? 1 : -1) * (bucketRank(a[0]) - bucketRank(b[0])),
    );
    // Ordem exibida (achatada) — base para seleção por intervalo
    return { groupList: gl, flat: gl.flatMap(([, rows]) => rows) };
  }, [visible, sortDesc]);

  // ===== Lista virtualizada: achata grupos+linhas numa única sequência e
  // renderiza só a janela visível (+overscan). É o que Outlook/Thunderbird
  // fazem — sem isso uma pasta com dezenas de milhares de mensagens gera
  // dezenas de milhares de nós DOM e a rolagem trava.
  type VRow =
    | { t: 'g'; label: string }
    | { t: 'r'; m: ItemSummary }
    | { t: 's'; r: GlobalSearchHit };
  const vrows = React.useMemo<VRow[]>(() => {
    if (searching) return visibleSearch!.map((r) => ({ t: 's' as const, r }));
    // Ordenação por outra coluna (De/Assunto/Tamanho) → lista plana, sem
    // cabeçalhos de faixa de data, para a ordenação valer de ponta a ponta.
    if (sortKey !== 'date') return visible.map((m) => ({ t: 'r' as const, m }));
    const out: VRow[] = [];
    for (const [bucket, rows] of groupList) {
      out.push({ t: 'g', label: bucket });
      for (const m of rows) out.push({ t: 'r', m });
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [groupList, visible, visibleSearch, searching, sortKey]);

  // Ordem real exibida (agrupada por data, plana ou filtrada por não-lidos) —
  // diferente de `items` quando ordenação/busca/agrupamento reorganizam a
  // tela. O App usa para achar o "próximo" item visualmente após excluir.
  if (displayedRef) {
    displayedRef.current = searching
      ? (visibleSearch ?? []).map((r) => r.item)
      : sortKey === 'date'
        ? flat
        : visible;
  }

  const OVERSCAN = 12;
  const totalV = vrows.length;
  const startIdx = Math.max(0, Math.floor(viewport.top / VROW_H) - OVERSCAN);
  const endIdx = Math.min(totalV, Math.ceil((viewport.top + viewport.height) / VROW_H) + OVERSCAN);
  const vslice = vrows.slice(startIdx, endIdx);

  const clickItem = (m: ItemSummary, e: React.MouseEvent) => {
    const pos = flat.indexOf(m);
    if (e.shiftKey && anchorRef.current !== null) {
      const a = Math.min(anchorRef.current, pos);
      const b = Math.max(anchorRef.current, pos);
      onSelChange(flat.slice(a, b + 1).map((x) => x.index), m.index);
      onSelect(m);
    } else if (e.ctrlKey || e.metaKey) {
      const next = new Set(selIndexes);
      if (next.has(m.index)) next.delete(m.index);
      else next.add(m.index);
      anchorRef.current = pos;
      onSelChange([...next], m.index);
      onSelect(m);
    } else {
      anchorRef.current = pos;
      onSelChange([m.index], m.index);
      onSelect(m);
    }
  };

  const hcell = (key: SortKey, label: string) => (
    <span className="hcell" onClick={() => clickSort(key)} title="Clique para ordenar">
      {label}
      {sortKey === key && <span className="sdir">{sortDesc ? '▼' : '▲'}</span>}
      {key !== 'subject' && (
        <span
          className="col-resize"
          onMouseDown={startResize(key as keyof ColWidths)}
          onClick={(e) => e.stopPropagation()}
          title="Arraste para ajustar a largura"
        />
      )}
    </span>
  );

  return (
    <div
      className="message-pane"
      ref={paneRef}
      style={
        {
          '--col-from': `${cols.from}px`,
          '--col-date': `${cols.date}px`,
          '--col-size': `${cols.size}px`,
        } as React.CSSProperties
      }
    >
      <div className="search-bar">
        <Icon name="search" size={14} />
        <input
          type="search"
          placeholder={t('Pesquisar emails (assunto, remetente, destinatário, prévia)…')}
          onChange={(e) => {
            setQuery(e.target.value);
            onSearch(e.target.value);
          }}
        />
        {searchBusy && <span className="scope busy">{t('pesquisando…')}</span>}
        {searching && !searchBusy && (
          <span className="scope">{visibleSearch!.length} {t('resultado(s)')}</span>
        )}
      </div>
      {/* Filtro Todas / Não lidos — como o Outlook 2016 (oculto em pastas de enviados) */}
      {(!noUnreadFilter || (categories?.length ?? 0) > 0) && (
        <div className="ml-filter">
          {!noUnreadFilter && (
            <>
              <button
                className={!unreadOnly ? 'active' : ''}
                onClick={() => setUnreadOnly(false)}
              >
                {t('Todas')}
              </button>
              <button
                className={unreadOnly ? 'active' : ''}
                onClick={() => setUnreadOnly(true)}
              >
                {t('Não lidos')}{unreadCount > 0 ? ` ${unreadCount}` : ''}
              </button>
            </>
          )}
          {(categories?.length ?? 0) > 0 && (
            <select
              className={`cat-filter${catFilter !== null ? ' active' : ''}`}
              value={catFilter ?? ''}
              onChange={(e) => setCatFilter(e.target.value || null)}
              title={t('Filtrar por categoria')}
            >
              <option value="">{t('Todas as categorias')}</option>
              {categories!.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.name}
                </option>
              ))}
              <option value="none">{t('Sem categoria')}</option>
            </select>
          )}
          {catFilter !== null && (
            <button
              className="cat-clear"
              title={t('Limpar filtro de categoria')}
              onClick={() => setCatFilter(null)}
            >
              <Icon name="close" size={11} />
            </button>
          )}
        </div>
      )}
      <div className="msg-header">
        <span className="ml-icons" />
        {hcell('from', 'De')}
        {hcell('subject', 'Assunto')}
        {hcell('date', 'Recebido em')}
        {hcell('size', 'Tamanho')}
      </div>
      <div
        className="message-list"
        tabIndex={0}
        ref={listRef}
        onScroll={(e) => {
          const el = e.currentTarget;
          setViewport((v) => (v.top === el.scrollTop ? v : { ...v, top: el.scrollTop }));
          // Scroll infinito: perto do fim → carrega a próxima página
          if (searching || loading || !hasMore) return;
          if (el.scrollTop + el.clientHeight >= el.scrollHeight - 220) onLoadMore();
        }}
        onKeyDown={(e) => {
          if (!searching && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
            e.preventDefault();
            onSelChange(flat.map((x) => x.index), selectedIndex ?? flat[0]?.index ?? -1);
          } else if (e.key === 'Delete' && selIndexes.size > 0) {
            e.preventDefault();
            onDeleteKey?.();
          } else if (
            !searching &&
            !e.ctrlKey &&
            !e.metaKey &&
            !e.altKey &&
            (e.key === 'ArrowDown' || e.key === 'ArrowUp')
          ) {
            // Navegação por setas como no Outlook: muda a seleção (abrindo a
            // mensagem no painel) e rola a lista até a linha — o scroll
            // virtualizado não segue o foco sozinho.
            if (!flat.length) return;
            e.preventDefault();
            const dir = e.key === 'ArrowDown' ? 1 : -1;
            const cur = flat.findIndex((m) => m.index === selectedIndex);
            const pos =
              cur < 0
                ? dir > 0
                  ? 0
                  : flat.length - 1
                : Math.min(flat.length - 1, Math.max(0, cur + dir));
            const m = flat[pos];
            if (!m) return;
            if (e.shiftKey) {
              // Shift+seta estende a seleção a partir da âncora
              if (anchorRef.current === null) anchorRef.current = cur < 0 ? pos : cur;
              const a = Math.min(anchorRef.current, pos);
              const b = Math.max(anchorRef.current, pos);
              onSelChange(flat.slice(a, b + 1).map((x) => x.index), m.index);
            } else {
              anchorRef.current = pos;
              onSelChange([m.index], m.index);
            }
            onSelect(m);
            // Rola até a linha alvo se ela estiver fora da viewport —
            // posição real = índice na sequência virtualizada × altura.
            const row = vrows.findIndex((v) => v.t === 'r' && v.m === m);
            const el = listRef.current;
            if (row >= 0 && el) {
              const y = row * VROW_H;
              if (y < el.scrollTop) el.scrollTop = y;
              else if (y + VROW_H > el.scrollTop + el.clientHeight) {
                el.scrollTop = y + VROW_H - el.clientHeight;
              }
            }
          }
        }}
      >
        {/* Espaçadores mantêm a altura total do scroll — só a janela
            visível (vslice) existe de verdade no DOM */}
        <div style={{ height: startIdx * VROW_H }} />
        {vslice.map((v) =>
          v.t === 'g' ? (
            <div className="msg-group" key={`g|${v.label}`}>
              <Icon name="chevronD" size={11} />
              <span>{v.label}</span>
            </div>
          ) : v.t === 'r' ? (
            <Row
              key={v.m.index}
              m={v.m}
              selected={selIndexes.has(v.m.index)}
              catById={catById}
              dragPayload={
                dragRef
                  ? {
                      ...dragRef,
                      indexes:
                        selIndexes.size > 1 && selIndexes.has(v.m.index)
                          ? [...selIndexes]
                          : [v.m.index],
                    }
                  : null
              }
              onSelect={(e) => clickItem(v.m, e)}
              onOpen={
                onOpenItem
                  ? () => onOpenItem(v.m, dragRef?.sourceId, dragRef?.folderId)
                  : undefined
              }
              onCtx={
                dragRef
                  ? (e) => openMsgCtx(e, v.m, dragRef.sourceId, dragRef.folderId)
                  : undefined
              }
            />
          ) : (
            <Row
              key={`${v.r.sourceId}:${v.r.folderId}:${v.r.item.index}`}
              m={v.r.item}
              selected={selectedIndex === v.r.item.index}
              catById={catById}
              dragPayload={null}
              onSelect={() => onSelect(v.r.item, v.r.sourceId, v.r.folderId)}
              onOpen={
                onOpenItem
                  ? () => onOpenItem(v.r.item, v.r.sourceId, v.r.folderId)
                  : undefined
              }
              onCtx={(e) => openMsgCtx(e, v.r.item, v.r.sourceId, v.r.folderId)}
              hl={hlTokens}
            />
          ),
        )}
        <div style={{ height: (totalV - endIdx) * VROW_H }} />
        {loading && (
          <div className="list-loading">
            <div className="spinner" />
          </div>
        )}
        {!searching && !loading && items.length === 0 && (
          <div className="empty-list">{t('Nenhum item nesta pasta')}</div>
        )}
        {!searching && !loading && items.length > 0 && visible.length === 0 && (
          <div className="empty-list">{t('Nenhuma mensagem não lida')}</div>
        )}
        {searching && visibleSearch!.length === 0 && (
          <div className="empty-list">{t('Nenhum resultado')}</div>
        )}
      </div>
      {/* Rodapé só aparece quando há algo útil a dizer — contagem de itens
          recalculava a cada pasta clicada e o "Carregando…" piscava com a
          paginação em segundo plano. */}
      {(searching || selIndexes.size > 1) && (
        <div className="list-footer">
          <span>
            {searching
              ? `${visibleSearch!.length} ${t('resultado(s)')}`
              : `${selIndexes.size} ${t('selecionado(s)')}`}
          </span>
        </div>
      )}

      {/* Menu de contexto da mensagem (botão direito) — estilo Outlook */}
      {msgCtx && (
        <div
          className="ctx-menu"
          style={{
            position: 'fixed',
            left: Math.min(msgCtx.x, window.innerWidth - 250),
            top: Math.max(8, Math.min(msgCtx.y, window.innerHeight - 560)),
            zIndex: 130,
          }}
          onClick={(e) => e.stopPropagation()}
          onContextMenu={(e) => e.preventDefault()}
        >
          {msgCtx.mode === 'menu' &&
            (() => {
              const act = (action: string) => () => {
                if (action === '__move__' || action === '__cats__') {
                  setMsgCtx({ ...msgCtx, mode: action === '__move__' ? 'move' : 'cats' });
                  return;
                }
                setMsgCtx(null);
                onCtxAction?.(action, msgCtx);
              };
              const it = (
                label: string,
                icon: string,
                action: string,
                disabled = false,
              ) => (
                <button key={label} disabled={disabled} onClick={act(action)}>
                  <Icon name={icon} size={14} /> {t(label)}
                </button>
              );
              return (
                <>
                  {it('Responder', 'reply', 'reply')}
                  {it('Responder a Todos', 'reply', 'replyAll')}
                  {it('Encaminhar', 'forward', 'forward')}
                  <div className="ctx-sep" />
                  {it(
                    msgCtx.isRead ? 'Marcar como Não Lida' : 'Marcar como Lida',
                    'envelope',
                    msgCtx.isRead ? 'markUnread' : 'markRead',
                  )}
                  {it('Mover…', 'open', '__move__')}
                  {it('Regras…', 'sendReceive', 'rules')}
                  <div className="ctx-sep" />
                  {it('Ignorar', 'close', 'ignore', true)}
                  {it('Lixo Eletrônico', 'junk', 'junk')}
                  {it('Bloquear Remetente', 'close', 'blockSender')}
                  {it('Arquivar', 'archive', 'archive')}
                  {it('Excluir', 'delete', 'delete')}
                  <div className="ctx-sep" />
                  {it('Categorizar', 'flag', '__cats__')}
                  {it('Localizar Relacionado', 'search', 'findRelated', true)}
                </>
              );
            })()}

          {msgCtx.mode === 'move' && (
            <div className="ctx-move">
              <div className="ctx-move-title">
                {t('Mover {alvo} para:', {
                  alvo:
                    msgCtx.indexes.length > 1
                      ? t('{n} itens', { n: msgCtx.indexes.length })
                      : t('item'),
                })}
              </div>
              {(moveTargets?.(msgCtx.sourceId) ?? [])
                .filter((t) => t.id !== msgCtx.folderId)
                .map((t) => (
                  <button
                    key={t.id}
                    onClick={() => {
                      setMsgCtx(null);
                      onCtxAction?.(`move:${t.id}`, msgCtx);
                    }}
                  >
                    {t.label}
                  </button>
                ))}
              <div className="ctx-sep" />
              <button onClick={() => setMsgCtx(null)}>{t('Cancelar')}</button>
            </div>
          )}

          {/* Submenu Categorizar — como o Outlook: clique alterna a marca;
              vários selecionados aplicam a todos. */}
          {msgCtx.mode === 'cats' && (
            <div className="ctx-move">
              <div className="ctx-move-title">
                {t('Categorizar {alvo}:', {
                  alvo:
                    msgCtx.indexes.length > 1
                      ? t('{n} itens', { n: msgCtx.indexes.length })
                      : t('item'),
                })}
              </div>
              {(categories ?? []).map((c) => {
                const on = msgCtx.cats.includes(c.id);
                return (
                  <button
                    key={c.id}
                    onClick={() => {
                      setMsgCtx(null);
                      onCtxAction?.(`cat:${c.id}`, msgCtx);
                    }}
                  >
                    <span className="cat-dot" style={{ background: c.color }} />
                    <span style={{ flex: 1 }}>{c.name}</span>
                    {on && <Icon name="check" size={12} />}
                  </button>
                );
              })}
              <div className="ctx-sep" />
              <button
                disabled={!msgCtx.cats.length}
                onClick={() => {
                  setMsgCtx(null);
                  onCtxAction?.('catClear', msgCtx);
                }}
              >
                {t('Limpar Todas as Categorias')}
              </button>
              <button
                onClick={() => {
                  setMsgCtx(null);
                  onCtxAction?.('catsManage', msgCtx);
                }}
              >
                {t('Todas as Categorias…')}
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
