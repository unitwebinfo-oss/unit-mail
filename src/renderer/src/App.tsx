import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Ribbon, type ReadingPaneLayout } from './components/Ribbon';
import { Splitter } from './components/Splitter';
import { FolderPane, type ItemDragPayload } from './components/FolderPane';
import { MessageList, type GlobalSearchHit } from './components/MessageList';
import { ReadingPane } from './components/ReadingPane';
import { ModuleBar } from './components/ModuleBar';
import { AccountsSettings, type SettingsTab } from './components/AccountsSettings';
import { OnboardingWizard } from './components/OnboardingWizard';
import type { ComposeDraft } from './components/ComposeModal';
import { NewItemDialog } from './components/NewItemDialog';
import { SyncProgressDialog } from './components/SyncProgressDialog';
import { ConfirmDialog } from './components/ConfirmDialog';
import { CategoriesDialog } from './components/CategoriesDialog';
import { ReminderDialog, type ReminderEvent } from './components/ReminderDialog';
import { CalendarModule } from './modules/CalendarModule';
import { ContactsModule } from './modules/ContactsModule';
import { TasksModule } from './modules/TasksModule';
import { Icon } from './icons';
import { findFirstFolder, containerClassForModule, applyTheme, arcFolderId, playReminderSound } from './utils';
import { isDefaultFolderId } from '../../shared/types';
import { folderKindOf, folderName as stdFolderName, localizeFolderName } from '../../shared/folderNames';
import { initLang, locale, t, useLang } from './i18n';
import type {
  AccountConfig,
  FolderNode,
  ItemDetail,
  ItemSummary,
  MailCategory,
  MailSource,
  ModuleId,
  NewSpecialItem,
  SearchResult,
  SyncProgress,
} from '../../shared/types';

// Página menor = primeira abertura de pasta mais rápida. O pst-extractor lê
// cada item do zero na 1ª vez que é tocado (sem cache de disco) — pedir 500
// itens de uma pasta grande (ex.: 10 mil msgs) custava ~2s só na 1ª página.
// 100 é visível instantaneamente e o scroll infinito busca o resto sob demanda.
const PAGE = 100;

/** Rascunhos abrem numa janela do SO separada — como no Outlook. */
function openDraft(d: ComposeDraft): void {
  void window.outlook.openComposeWindow(d);
}

interface Selection {
  sourceId: string;
  folder: FolderNode;
}

function findNode(nodes: FolderNode[], id: string): FolderNode | null {
  for (const n of nodes) {
    if (n.id === id) return n;
    const f = findNode(n.children, id);
    if (f) return f;
  }
  return null;
}

/** Decrementa/incrementa o contador de não-lidos de uma pasta sem refazer a árvore inteira. */
function patchUnread(nodes: FolderNode[], id: string, delta: number): FolderNode[] {
  return nodes.map((n) => {
    if (n.id === id) return { ...n, unreadCount: Math.max(0, (n.unreadCount || 0) + delta) };
    if (n.children.length) {
      const children = patchUnread(n.children, id, delta);
      if (children !== n.children) return { ...n, children };
    }
    return n;
  });
}

/**
 * Ordem personalizada das bases (arquivos de dados) no painel — mesma
 * persistência de folderOrders, sob a chave especial 'sources'.
 */
function applySourceOrder(sources: MailSource[], orders: Record<string, string[]>): MailSource[] {
  const order = orders['sources'];
  if (!order) return sources;
  return [...sources].sort((a, b) => {
    const ia = order.indexOf(a.id);
    const ib = order.indexOf(b.id);
    return (ia === -1 ? Number.MAX_SAFE_INTEGER : ia) - (ib === -1 ? Number.MAX_SAFE_INTEGER : ib);
  });
}

/** Aplica a ordem de pastas personalizada (persistida) a cada nível da árvore. */
function applyFolderOrder(
  nodes: FolderNode[],
  sourceId: string,
  parentId: string,
  orders: Record<string, string[]>,
): FolderNode[] {
  const order = orders[`${sourceId}|${parentId}`];
  const sorted = order
    ? [...nodes].sort((a, b) => {
        const ia = order.indexOf(a.id);
        const ib = order.indexOf(b.id);
        return (ia === -1 ? Number.MAX_SAFE_INTEGER : ia) - (ib === -1 ? Number.MAX_SAFE_INTEGER : ib);
      })
    : nodes;
  return sorted.map((n) => ({ ...n, children: applyFolderOrder(n.children, sourceId, n.id, orders) }));
}

export default function App() {
  const lang = useLang();
  const [sources, setSources] = useState<MailSource[]>([]);
  const [folderTrees, setFolderTrees] = useState<Record<string, FolderNode[]>>({});
  const [accounts, setAccounts] = useState<AccountConfig[]>([]);
  const [module, setModule] = useState<ModuleId>('mail');
  const [selected, setSelected] = useState<Selection | null>(null);
  const [items, setItems] = useState<ItemSummary[]>([]);
  /** Índices multi-selecionados na pasta atual (Ctrl/Shift-clique na lista). */
  const [selIndexes, setSelIndexes] = useState<Set<number>>(new Set());
  const [itemCount, setItemCount] = useState(0);
  const [detail, setDetail] = useState<ItemDetail | null>(null);
  const [detailRef, setDetailRef] = useState<{ sourceId: string; folderId: string } | null>(null);
  /** Spinner leve dentro do painel de leitura (não bloqueia a tela). */
  const [detailLoading, setDetailLoading] = useState(false);
  /**
   * Cache de detalhes por mensagem — clicar de novo num item já lido é
   * instantâneo (o getItem de PST re-ler o arquivo todo seria pesado).
   * Chaveado por stableId (não por índice): o índice posicional desloca
   * quando a pasta recebe/perde itens e a entrada cacheada abriria a
   * mensagem errada.
   */
  const detailCache = useRef(new Map<string, ItemDetail>());
  const detailKey = (sid: string, fid: string, idx: number, stableId?: string) =>
    `${sid}|${fid}|${stableId ?? `i${idx}`}`;
  const [searchResults, setSearchResults] = useState<GlobalSearchHit[] | null>(null);
  const [loading, setLoading] = useState('');
  /** Loading leve da lista de mensagens (troca de pasta) — overlay só na lista. */
  const [listLoading, setListLoading] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  // Tarefas do Enviar/Receber em andamento (barra de status + diálogo)
  const [syncTasks, setSyncTasks] = useState<SyncProgress[]>([]);
  /** Mensagem aberta em janela própria (duplo clique na lista). */
  const [openMsg, setOpenMsg] = useState<{
    detail: ItemDetail;
    sourceId: string;
    folderId: string;
  } | null>(null);
  const [showSyncDialog, setShowSyncDialog] = useState(false);
  // Confirmação "Tem certeza?" antes de ações destrutivas
  const [confirm, setConfirm] = useState<{
    title: string;
    message: string;
    confirmLabel?: string;
    onOk: () => void;
  } | null>(null);
  const [folderOrders, setFolderOrders] = useState<Record<string, string[]>>({});
  /** Windows usa barra de título integrada (titleBarOverlay) — faixa azul do app. */
  const [winChrome, setWinChrome] = useState(false);
  /** Janela de mensagem maximizada (duplo clique num email). */
  const [msgMax, setMsgMax] = useState(false);
  /** Release mais nova no GitHub — exibida nas Configurações. */
  const [showSettings, setShowSettings] = useState(false);
  const [settingsInitTab, setSettingsInitTab] = useState<SettingsTab>('email');
  const [showOnboarding, setShowOnboarding] = useState(false);
  /** Pasta de dados configurada ausente no boot (ex.: drive removível não
   *  montado) — o main não restaura nada; este modal cobre a janela. */
  const [profileMissing, setProfileMissing] = useState<string | null>(null);

  /** Diálogo "Novos Itens" — compromisso, contato ou tarefa. */
  const [newItemKind, setNewItemKind] = useState<'appt' | 'contact' | 'task' | null>(null);
  /** Item especial em edição (compromisso/contato/tarefa criado nesta base). */
  const [reminders, setReminders] = useState<ReminderEvent[]>([]);
  const [editingItem, setEditingItem] = useState<{
    sourceId: string;
    folderId: string;
    index: number;
    stableId?: string;
    item: NewSpecialItem;
  } | null>(null);
  /** Base onde itens especiais são criados — sempre a principal (arquivo de dados padrão). */
  const [apptSourceId, setApptSourceId] = useState('');
  /** Caminho do arquivo de dados principal — cacheado para decisões síncronas. */
  const defaultStoreRef = useRef('');
  const [loadImages, setLoadImages] = useState(false);
  /** Lista-mestre de categorias coloridas (Categorizar). */
  const [categories, setCategories] = useState<MailCategory[]>([]);
  /** Diálogo "Todas as Categorias". */
  const [showCatsDlg, setShowCatsDlg] = useState(false);
  const [rpLayout, setRpLayout] = useState<ReadingPaneLayout>(
    () => (localStorage.getItem('rpLayout') as ReadingPaneLayout) || 'right',
  );
  // Larguras/alturas dos painéis redimensionáveis (persistidas)
  const [folderW, setFolderW] = useState(() => Number(localStorage.getItem('pane.folderW')) || 230);
  const [msgW, setMsgW] = useState(() => Number(localStorage.getItem('pane.msgW')) || 380);
  const [msgH, setMsgH] = useState(() => Number(localStorage.getItem('pane.msgH')) || 320);
  const dragBase = useRef(0);
  const searchTimer = useRef<ReturnType<typeof setTimeout>>();
  /** Módulo atual acessível dentro de callbacks (para persistir a seleção por módulo). */
  const moduleRef = useRef(module);
  moduleRef.current = module;
  /** Pasta selecionada acessível dentro de listeners de eventos do main. */
  const selectedRef = useRef(selected);
  selectedRef.current = selected;
  const itemsRef = useRef(items);
  itemsRef.current = items;
  const detailRefState = useRef<{ detail: ItemDetail | null; ref: typeof detailRef }>({
    detail,
    ref: detailRef,
  });
  detailRefState.current = { detail, ref: detailRef };

  const paneResize = (
    key: string,
    setter: (v: number) => void,
    base: number,
    delta: number,
    lo: number,
    hi: number,
  ) => {
    const v = Math.min(hi, Math.max(lo, base + delta));
    setter(v);
    localStorage.setItem(key, String(v));
  };

  const setLayout = useCallback((l: ReadingPaneLayout) => {
    setRpLayout(l);
    localStorage.setItem('rpLayout', l);
  }, []);

  const refreshImagePref = useCallback(() => {
    window.outlook.getLoadImages?.().then((v) => setLoadImages(!!v));
  }, []);

  const [settingsEditId, setSettingsEditId] = useState<string | undefined>(undefined);
  const [settingsAction, setSettingsAction] = useState<'importPst' | 'importTb' | undefined>(undefined);
  const openSettings = useCallback(
    (tab: SettingsTab = 'email', accountId?: string, action?: 'importPst' | 'importTb') => {
      setSettingsInitTab(tab);
      setSettingsEditId(accountId);
      setSettingsAction(action);
      setShowSettings(true);
    },
    [],
  );

  const refreshAccounts = useCallback(() => {
    window.outlook.listAccounts().then(setAccounts);
    window.outlook.getFolderOrders?.().then((o) => setFolderOrders(o || {}));
    window.outlook.getDefaultStore?.().then((p) => {
      defaultStoreRef.current = p || '';
    });
  }, []);

  const refreshSources = useCallback(async () => {
    const list = await window.outlook.listSources();
    setSources(list);
    const trees: Record<string, FolderNode[]> = {};
    await Promise.all(
      list.map(async (s) => {
        try {
          trees[s.id] = await window.outlook.getFolderTree(s.id);
        } catch {
          trees[s.id] = [];
        }
      }),
    );
    setFolderTrees(trees);
    return { list, trees };
  }, []);

  /** Sequência de carga — trocar de pasta aborta o carregamento em segundo
   *  plano da pasta anterior. */
  const loadSeq = useRef(0);
  /**
   * Cache de itens por pasta — como o Outlook: voltar a uma pasta já
   * visitada mostra a lista instantaneamente e revalida em background
   * (se chegou mensagem nova, ela entra sem esperar).
   */
  const itemsCache = useRef(new Map<string, ItemSummary[]>());
  /** Pasta dona dos itens atualmente em tela — evita gravar itens da pasta
   *  antiga sob a chave da nova durante a troca. */
  const itemsKey = useRef('');
  /** `items` cobre a pasta inteira — o efeito-espelho só grava no itemsCache
   *  quando completa, senão listas parciais envenenariam o cache. */
  const itemsComplete = useRef(false);
  /** Ordem real exibida na lista (agrupada/ordenada/filtrada) — alimentada
   *  pela MessageList; usada para achar o "próximo" item visual ao excluir. */
  const displayedItemsRef = useRef<{ key: string; list: ItemSummary[] }>({
    key: '',
    list: [],
  });

  /**
   * Baixa a pasta inteira em blocos alimentando a lista e o cache.
   * `seq` deve ser o loadSeq corrente — aborta se outro load assumir.
   * Usada tanto por loadFolder quanto pelo refresh silencioso pós-sync.
   */
  const fillFolder = useCallback(
    async (sourceId: string, folder: FolderNode, count: number, seq: number) => {
      const cacheKey = `${sourceId}|${folder.id}`;
      // Primeira página instantânea — a lista já aparece e responde.
      const first = await window.outlook.getItems(sourceId, folder.id, 0, PAGE);
      if (seq !== loadSeq.current) return;
      // A pasta recarregou: os índices podem ter mudado (mensagem nova no
      // topo desloca todos). O detailCache é chaveado por stableId, então
      // sobrevive ao deslocamento — só entradas de itens que sumiram ficam
      // órfãs e saem pelo limite do cache.
      for (const k of [...detailCache.current.keys()]) {
        if (k.startsWith(`${cacheKey}|i`)) detailCache.current.delete(k);
      }
      setItems(first);
      setItemCount(count);
      itemsKey.current = cacheKey;
      itemsComplete.current = false;
      setListLoading(false);
      // O resto da pasta chega em segundo plano, em blocos — ordenar/
      // pesquisar/filtrar passam a cobrir tudo assim que terminar.
      // loadingMore bloqueia o scroll infinito enquanto isso roda.
      loadingMore.current = true;
      try {
        // Dedupe por stableId: uma entrega no meio da paginação desloca o
        // espaço de índices e a página seguinte repetiria mensagens já
        // exibidas (a lista duplicada depois apontava para o item errado).
        const seen = new Set(first.map((i) => i.stableId ?? `i${i.index}`));
        let offset = first.length;
        const acc = [...first];
        while (offset < count) {
          const more = await window.outlook.getItems(sourceId, folder.id, offset, 500);
          if (seq !== loadSeq.current || more.length === 0) break;
          offset += more.length;
          const fresh = more.filter((i) => {
            const k = i.stableId ?? `i${i.index}`;
            if (seen.has(k)) return false;
            seen.add(k);
            return true;
          });
          acc.push(...fresh);
          if (fresh.length) setItems((prev) => [...prev, ...fresh]);
        }
        // Só grava no cache se o carregamento completou — cache parcial
        // mostraria a pasta "pela metade" na próxima visita.
        if (seq === loadSeq.current && offset >= count) {
          itemsComplete.current = true;
          // LRU simples: no máximo 8 pastas grandes em memória
          if (itemsCache.current.size >= 8 && !itemsCache.current.has(cacheKey)) {
            const oldest = itemsCache.current.keys().next().value;
            if (oldest) itemsCache.current.delete(oldest);
          }
          itemsCache.current.set(cacheKey, acc);
        }
      } finally {
        loadingMore.current = false;
      }
    },
    [],
  );

  const loadFolder = useCallback(async (sourceId: string, folder: FolderNode, silent = false) => {
    // Se já é a pasta selecionada e os itens já estão carregados, ignora —
    // evita o flash/reload ao clicar de novo na mesma pasta
    setSelected((cur) =>
      cur?.sourceId === sourceId && cur.folder.id === folder.id ? cur : { sourceId, folder },
    );
    if (!silent) {
      // Navegação real de pasta: limpa seleção/detalhe. No reload SILENCIOSO
      // (mesma pasta, pós-exclusão/sync) preserva — era ele que apagava a
      // seleção recém-feita e causava a "piscada" ao deletar após trocar
      // de pasta (qualquer divergência de contagem derrubava o sucessor).
      setDetail(null);
      setDetailRef(null);
      setSelIndexes(new Set());
      setSearchResults(null);
    }
    detailCache.current.clear();
    // Lembra a última pasta aberta por módulo (como o Outlook)
    localStorage.setItem(`sel.${moduleRef.current}`, `${sourceId}|${folder.id}`);
    // Loading inline na lista — mantém os itens anteriores até os novos chegarem
    const cacheKey = `${sourceId}|${folder.id}`;
    const cached = itemsCache.current.get(cacheKey);
    if (cached) {
      // Acerto de cache — mostra tudo na hora (sem spinner). O cache só
      // guarda cargas completas (itemsComplete), então já vale.
      setItems(cached);
      setItemCount(cached.length);
      itemsKey.current = cacheKey;
      itemsComplete.current = true;
    } else {
      // Reload silencioso (pós-exclusão, por ex.): mantém a lista atual na
      // tela em vez de cobri-la com o overlay de carregamento.
      if (!silent) setListLoading(true);
    }
    const seq = ++loadSeq.current;
    try {
      const count = await window.outlook.getItemCount(sourceId, folder.id);
      if (seq !== loadSeq.current) return;
      // Como o Outlook: se a contagem não mudou, o cache continua válido —
      // nem re-baixa os itens. Mutações desta sessão (lido/excluído/movido)
      // já foram refletidas no cache pelo efeito de sincronização abaixo.
      if (cached && count === cached.length) {
        setItemCount(count);
        return;
      }
      await fillFolder(sourceId, folder, count, seq);
    } catch (e) {
      if (seq === loadSeq.current) setError(t('Erro ao ler pasta: {msg}', { msg: (e as Error).message }));
    } finally {
      if (seq === loadSeq.current) setListLoading(false);
    }
  }, [fillFolder]);

  /**
   * Revalida a pasta aberta depois que o Enviar/Receber entregou itens —
   * versão silenciosa de loadFolder: não limpa seleção, mensagem aberta
   * nem pesquisa. O count-check torna a verificação barata quando nada
   * mudou nesta pasta (ex.: a regra entregou em outra).
   */
  const refreshOpenFolder = useCallback(async () => {
    const sel = selectedRef.current;
    if (!sel) return;
    const cacheKey = `${sel.sourceId}|${sel.folder.id}`;
    const seqAtStart = loadSeq.current;
    try {
      const count = await window.outlook.getItemCount(sel.sourceId, sel.folder.id);
      // Se uma navegação começou durante a contagem, não atrapalha —
      // aborta antes de assumir o seq (não há await entre o check e o bump)
      const cur = selectedRef.current;
      if (
        loadSeq.current !== seqAtStart ||
        cur?.sourceId !== sel.sourceId ||
        cur.folder.id !== sel.folder.id
      ) {
        return;
      }
      const cached = itemsCache.current.get(cacheKey);
      if (count === (cached?.length ?? -1)) return;
      const seq = ++loadSeq.current;
      await fillFolder(sel.sourceId, sel.folder, count, seq);
    } catch {
      // Refresh de fundo — falha silenciosa, a próxima sync tenta de novo
    }
  }, [fillFolder]);

  /**
   * Sincroniza o cache com a lista em tela — mutações desta sessão
   * (marcar lido, excluir, mover) atualizam a entrada da pasta sem
   * precisar re-baixar. `itemsKey` garante que os itens são gravados
   * sob a pasta que os gerou, não a recém-selecionada.
   */
  useEffect(() => {
    // Só espelha cargas completas — durante o preenchimento em blocos `items`
    // é parcial e gravaria uma lista pela metade em cima do cache válido.
    if (itemsKey.current && itemsComplete.current) {
      itemsCache.current.set(itemsKey.current, items);
    }
  }, [items]);

  const selectDefaultFolder = useCallback(
    (m: ModuleId, list: MailSource[], trees: Record<string, FolderNode[]>) => {
      const cc = containerClassForModule(m);
      // Restaura a última pasta aberta neste módulo (persistida por sourceId
      // estável — continua válida entre reinícios do app)
      const saved = localStorage.getItem(`sel.${m}`);
      if (saved) {
        const sep = saved.indexOf('|');
        const sid = saved.slice(0, sep);
        const fid = saved.slice(sep + 1);
        const node = findNode(trees[sid] || [], fid);
        const visible =
          node && (m === 'mail' || node.containerClass === cc);
        if (visible && list.some((s) => s.id === sid)) {
          return loadFolder(sid, node);
        }
      }
      // Calendário/Pessoas/Tarefas operam sempre no arquivo de dados principal —
      // ele vem primeiro na procura da pasta padrão do módulo.
      const def = m === 'mail' ? '' : defaultStoreRef.current;
      const ordered = def
        ? [...list].sort((a, b) => (a.filePath === def ? -1 : b.filePath === def ? 1 : 0))
        : list;
      for (const s of ordered) {
        const tree = trees[s.id] || [];
        if (m === 'mail' && s.kind === 'imap' && tree.length > 0) return loadFolder(s.id, tree[0]);
        const first = findFirstFolder(tree, cc);
        if (first) return loadFolder(s.id, first);
      }
      setSelected(null);
      setItems([]);
      setItemCount(0);
      return undefined;
    },
    [loadFolder],
  );

  const bootstrap = useCallback(async () => {
    setLoading(t('Carregando bases…'));
    try {
      const { list, trees } = await refreshSources();
      refreshAccounts();
      await selectDefaultFolder('mail', list, trees);
    } finally {
      setLoading('');
    }
  }, [refreshSources, refreshAccounts, selectDefaultFolder]);

  useEffect(() => {
    bootstrap();
    refreshImagePref();
    window.outlook.getCategories?.().then(setCategories).catch(() => {});
    // No Windows o app desenha a faixa de título azul (botões do SO por cima)
    // Barra azul integrada onde o titleBarOverlay existe (Windows e Linux);
    // no macOS a moldura nativa continua.
    window.outlook.getAppInfo?.().then((i) => setWinChrome(i.platform !== 'darwin'));
    // Verificação de atualização — feita dentro das Configurações.
    // Idioma salvo — decide t() e os nomes de exibição das pastas padrão
    window.outlook.getLanguage?.().then((l) => initLang(l));
    // Tema: aplica o salvo e reage quando o SO alterna claro/escuro
    window.outlook.getTheme?.().then((t) => applyTheme(t));
    // Primeiro paint real → o main só mostra a janela depois disso
    // (double-RAF = depois do frame ser composto).
    requestAnimationFrame(() =>
      requestAnimationFrame(() => window.outlook.appReady?.()),
    );
    const mq = window.matchMedia('(prefers-color-scheme: dark)');
    const onMq = () => {
      window.outlook.getTheme?.().then((t) => {
        if (t === 'system') applyTheme('system');
      });
    };
    mq.addEventListener('change', onMq);
    // Primeira execução → jornada de configuração inicial
    window.outlook.getSetupComplete?.().then((done) => {
      if (!done) setShowOnboarding(true);
    });
    // Pasta de dados configurada num caminho que não existe agora (drive
    // removível não montado na abertura, pasta renomeada/apagada). Nada é
    // restaurado nem sincronizado nesse estado — só o aviso bloqueante.
    window.outlook.getProfileStatus?.().then((ps) => {
      if (ps?.hasPointer && !ps.available) setProfileMissing(ps.target);
    });
    const unsubProfile = window.outlook.onProfileMissing?.((dir) => setProfileMissing(dir));
    // Bases restauradas em 2º plano pelo main — cada uma notifica ao
    // ficar pronta; um refresh por leva recarrega a árvore e ela aparece
    // na hora (antes, o app esperava todas abrirem na splash).
    let srcTimer: ReturnType<typeof setTimeout> | undefined;
    const unsubSrc = window.outlook.onSourcesChanged?.(() => {
      clearTimeout(srcTimer);
      srcTimer = setTimeout(async () => {
        const { list, trees } = await refreshSources();
        // Se ainda não havia pasta selecionada (nenhuma base aberta no
        // boot), seleciona a padrão agora que as bases chegaram.
        if (!selectedRef.current) {
          selectDefaultFolder('mail', list, trees);
        }
      }, 250);
    });
    // Enviar/Receber automático do main → aviso + refresh das pastas
    const unsub = window.outlook.onSyncDone?.((results) => {
      // Sem toast de "mensagens recebidas"/"nenhuma encontrada" — só erro avisa;
      // o progresso aparece no indicador Enviar/Receber e nas badges das pastas.
      const errs = results.filter((r) => r.error).map((r) => r.error).join('; ');
      if (errs) setError(t('Erro no envio/recebimento: {errs}', { errs }));
      refreshSources();
      // Emails novos (entregues localmente ou detectados no servidor IMAP)
      // → revalida a pasta aberta sem precisar sair e voltar nela.
      if (results.some((r) => r.archived > 0 || (r.newOnServer ?? 0) > 0)) refreshOpenFolder();
      // Limpa o indicador da barra de status alguns segundos após o fim
      setTimeout(() => {
        setSyncTasks((prev) =>
          prev.every((t) => t.status !== 'processing') ? [] : prev,
        );
      }, 5000);
    });
    // Progresso das tarefas de Enviar/Receber (barra de status + diálogo)
    const unsubProg = window.outlook.onSyncProgress?.((ev) => {
      setSyncTasks((prev) => {
        // O diálogo mostra só progresso por conta — o detalhe por mensagem
        // (ev.item) é descartado aqui para não acumular lista em memória.
        const i = prev.findIndex((t) => t.taskId === ev.taskId);
        const merged: SyncProgress = { ...ev, item: undefined, items: undefined };
        if (i < 0) return [...prev, merged];
        const next = [...prev];
        next[i] = merged;
        return next;
      });
    });
    // Envio em segundo plano: o modal fecha ao clicar em Enviar — o main
    // move a mensagem da Caixa de Saída para Itens Enviados ao concluir
    const unsubSend = window.outlook.onSendResult?.((r) => {
      refreshSources();
      // O envio move a mensagem de Caixa de Saída → Itens Enviados; o
      // refresh silencioso atualiza qualquer uma delas que estiver aberta
      refreshOpenFolder();
      if (r.ok) {
        setNotice(t('"{subj}" enviada para {folder}.', { subj: r.subject || t('Mensagem'), folder: stdFolderName('sent', lang) }));
        setTimeout(() => setNotice(''), 6000);
      } else {
        setError(
          t('Falha ao enviar "{subj}": {err} — a mensagem ficou em {folder}.', { subj: r.subject || t('mensagem'), err: r.error || '', folder: stdFolderName('outbox', lang) }),
        );
      }
    });
    // Link mailto: clicado num email → abre rascunho já preenchido
    const unsubMailto = window.outlook.onMailto?.((url) => {
      try {
        const m = /^mailto:([^?]*)\??(.*)$/i.exec(url);
        if (!m) return;
        const q = new URLSearchParams(m[2] || '');
        openDraft({
          accountId: defaultAccountId(),
          kind: 'new',
          to: decodeURIComponent(m[1] || ''),
          cc: q.get('cc') ?? undefined,
          bcc: q.get('bcc') ?? undefined,
          subject: q.get('subject') ?? '',
          bodyText: q.get('body') ?? '',
        });
      } catch {
        /* mailto malformado — ignora */
      }
    });
    // Lotes da busca em andamento (live search) — hits por pasta
    const unsubSearchBatch = window.outlook.onSearchBatch?.((sourceId, hits) =>
      pushSearchHits(sourceId, hits),
    );
    // Lembrete de compromisso → janela de lembretes estilo Outlook + som
    const unsubCal = window.outlook.onCalendarReminder?.((ev) => {
      playReminderSound();
      setReminders((prev) =>
        prev.some((r) => r.key === ev.key) ? prev : [...prev, ev],
      );
    });
    // Envio concluído numa janela de composição separada → recarrega listas
    const unsubSent2 = window.outlook.onComposeSent?.(() => void refreshOpenFolder());
    return () => {
      clearTimeout(srcTimer);
      unsub?.();
      unsubProfile?.();
      unsubSrc?.();
      unsubProg?.();
      unsubCal?.();
      unsubSend?.();
      unsubMailto?.();
      unsubSearchBatch?.();
      unsubSent2?.();
      mq.removeEventListener('change', onMq);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // O toast de erro some sozinho após alguns segundos (como o do Outlook)
  useEffect(() => {
    if (!error) return;
    const t = setTimeout(() => setError(''), 10000);
    return () => clearTimeout(t);
  }, [error]);

  const closeSource = useCallback(
    async (sourceId: string) => {
      await window.outlook.closeSource(sourceId);
      if (selected?.sourceId === sourceId) {
        setSelected(null);
        setItems([]);
        setDetail(null);
      }
      await refreshSources();
    },
    [selected, refreshSources],
  );

  const switchModule = useCallback(
    async (m: ModuleId) => {
      setModule(m);
      setDetail(null);
      setSearchResults(null);
      await selectDefaultFolder(m, sources, folderTrees);
    },
    [sources, folderTrees, selectDefaultFolder],
  );

  /**
   * Atualiza o badge de não-lidos da pasta na hora, só com aritmética local —
   * refazer a árvore inteira (mescla PST+sombra, recalcula todas as pastas)
   * a cada mensagem lida era a maior causa de lentidão/loading constante.
   * O estado de leitura já foi persistido pelo backend em getItem().
   */
  const decrementFolderUnread = useCallback((sid: string, fid: string) => {
    setFolderTrees((t) => (t[sid] ? { ...t, [sid]: patchUnread(t[sid], fid, -1) } : t));
  }, []);

  /** Sequência de seleção — uma leitura lenta (PST) não pode sobrescrever
   *  uma seleção mais nova que já resolveu. */
  const selSeq = useRef(0);

  const selectItem = useCallback(
    async (item: ItemSummary, sourceId?: string, folderId?: string) => {
      const sid = sourceId ?? selected?.sourceId;
      const fid = folderId ?? selected?.folder.id;
      if (sid === undefined || fid === undefined) return;
      const seq = ++selSeq.current;
      setDetailRef({ sourceId: sid, folderId: fid });
      const key = detailKey(sid, fid, item.index, item.stableId);
      const cached = detailCache.current.get(key);
      if (cached) {
        if (seq === selSeq.current) setDetail(cached);
      } else {
        // Spinner SÓ no painel de leitura — o overlay de tela inteira comia o
        // segundo clique do duplo clique e piscava a cada seleção.
        setDetailLoading(true);
        try {
          // stableId resolve a posição ATUAL do item — um índice velho
          // (pasta recebeu itens depois da listagem) não abre a mensagem errada.
          const d = await window.outlook.getItem(sid, fid, item.index, item.stableId);
          // Só cacheia se a seleção ainda é a atual — um getItem atrasado
          // (pré-seleção do sucessor na exclusão) resolveria depois de outra
          // seleção assumir e envenenaria o painel.
          if (seq === selSeq.current) {
            detailCache.current.set(key, d);
            if (detailCache.current.size > 400) {
              const oldest = detailCache.current.keys().next().value;
              if (oldest !== undefined) detailCache.current.delete(oldest);
            }
            setDetail(d);
          }
        } catch (e) {
          if (seq === selSeq.current)
            setError(t('Erro ao ler item: {msg}', { msg: (e as Error).message }));
        } finally {
          if (seq === selSeq.current) setDetailLoading(false);
        }
      }
      // Abrir marca como lida (backend já persistiu) — reflete na lista na
      // hora e no badge da pasta em background (rebuild da árvore é pesado).
      // Só para a seleção atual: chamada atrasada decrementaria o badge e
      // marcaria lido o item ERRADO (o índice já remapeou para outro).
      if (!item.isRead && seq === selSeq.current) {
        const same = (it: ItemSummary) =>
          item.stableId ? it.stableId === item.stableId : it.index === item.index;
        setItems((prev) => prev.map((it) => (same(it) ? { ...it, isRead: true } : it)));
        decrementFolderUnread(sid, fid);
      }
    },
    [selected, decrementFolderUnread],
  );

  /** Multi-seleção da lista (Ctrl/Shift-clique, Ctrl+A). */
  const onSelChange = useCallback((indexes: number[], _primary: number) => {
    setSelIndexes(new Set(indexes));
  }, []);



  /** Evita chamadas duplicadas do scroll infinito enquanto a página carrega. */
  const loadingMore = useRef(false);
  const loadMore = useCallback(async () => {
    if (!selected || loadingMore.current) return;
    loadingMore.current = true;
    // Spinner só no rodapé da lista — o overlay de tela inteira travava o
    // scroll e dava a impressão de loading constante.
    setListLoading(true);
    try {
      const more = await window.outlook.getItems(selected.sourceId, selected.folder.id, items.length, PAGE);
      setItems((prev) => [...prev, ...more]);
      // Última página — a lista passa a cobrir a pasta toda
      if (items.length + more.length >= itemCount) itemsComplete.current = true;
    } finally {
      loadingMore.current = false;
      setListLoading(false);
    }
  }, [selected, items.length, itemCount]);

  /** Sequência da pesquisa — resultados de buscas anteriores são descartados. */
  const searchSeq = useRef(0);
  const [searchBusy, setSearchBusy] = useState(false);
  /** Hits da busca em curso — lotes por pasta (evento) e retorno final se
   *  juntam aqui, com dedupe por chave. */
  const liveSearch = useRef<{
    seq: number;
    hits: GlobalSearchHit[];
    seen: Set<string>;
  } | null>(null);
  /** Rótulo da fonte pelo id — usado pelo handler de lotes (evento, sem closure). */
  const sourcesRef = useRef<MailSource[]>([]);
  sourcesRef.current = sources;

  /** Acrescenta hits à busca viva e publica na lista (ordenação por data). */
  const pushSearchHits = useCallback((sourceId: string, res: SearchResult[]) => {
    const live = liveSearch.current;
    if (!live || live.seq !== searchSeq.current) return;
    const label = sourcesRef.current.find((s) => s.id === sourceId)?.label ?? sourceId;
    let added = false;
    for (const r of res) {
      const k = `${sourceId}|${r.folderId}|${r.item.index}`;
      if (live.seen.has(k)) continue; // lote já emitido ao vivo — não duplica
      live.seen.add(k);
      live.hits.push({
        sourceId,
        sourceLabel: label,
        folderId: r.folderId,
        folderName: r.folderName,
        item: r.item,
      });
      added = true;
    }
    if (!added) return;
    // Recentes primeiro — igual à ordenação da caixa de entrada
    live.hits.sort((a, b) =>
      (b.item.deliveryTime || b.item.sentTime || '').localeCompare(
        a.item.deliveryTime || a.item.sentTime || '',
      ),
    );
    setSearchResults([...live.hits]);
  }, []);
  const onSearch = useCallback(
    (q: string) => {
      clearTimeout(searchTimer.current);
      const seq = ++searchSeq.current;
      const query = q.trim();
      if (!query) {
        liveSearch.current = null;
        setSearchResults(null);
        setSearchBusy(false);
        return;
      }
      // Live search: a lista vira resultado vazio já e os hits chegam por
      // pasta/fonte conforme o backend emite — sem overlay de tela inteira.
      setSearchResults([]);
      setSearchBusy(true);
      searchTimer.current = setTimeout(() => {
        liveSearch.current = { seq, hits: [], seen: new Set() };
        // Escopo = pasta selecionada (como o "Pasta Atual" do Outlook) —
        // outras pastas e outras bases não entram no resultado
        if (!selected) {
          setSearchBusy(false);
          return;
        }
        const { sourceId, folder } = selected;
        Promise.race([
          window.outlook.search(sourceId, query, undefined, folder.id),
          new Promise<SearchResult[]>((_, rej) =>
            setTimeout(() => rej(new Error('timeout')), 30000),
          ),
        ])
          .then((res) => pushSearchHits(sourceId, res))
          .catch(() => {
            // fonte indisponível/lenta para busca — ignora
          })
          .finally(() => {
            if (seq !== searchSeq.current) return;
            setSearchBusy(false);
            liveSearch.current = null;
          });
      }, 250);
    },
    [selected, pushSearchHits],
  );

  const saveAttachment = useCallback(
    async (attachIndex: number) => {
      if (!detail || !detailRef) return;
      try {
        await window.outlook.saveAttachment(
          detailRef.sourceId,
          detailRef.folderId,
          detail.index,
          attachIndex,
          detail.stableId,
        );
      } catch (e) {
        setError(t('Erro ao salvar anexo: {msg}', { msg: (e as Error).message }));
      }
    },
    [detail, detailRef],
  );

  /** Abre anexo sem salvar (cópia temporária, aplicativo padrão do SO). */
  const openAttachment = useCallback(
    async (attachIndex: number) => {
      if (!detail || !detailRef) return;
      try {
        await window.outlook.openAttachment(
          detailRef.sourceId,
          detailRef.folderId,
          detail.index,
          attachIndex,
          detail.stableId,
        );
      } catch (e) {
        setError(t('Erro ao abrir anexo: {msg}', { msg: (e as Error).message }));
      }
    },
    [detail, detailRef],
  );

  /** "Salvar todos" — grava todos os anexos numa pasta escolhida. */
  const saveAllAttachments = useCallback(async () => {
    if (!detail || !detailRef) return;
    try {
      const res = await window.outlook.saveAllAttachments(
        detailRef.sourceId,
        detailRef.folderId,
        detail.index,
        detail.stableId,
      );
      if (res && res.saved < res.total) {
        setError(t('{n} de {total} anexos salvos em "{dir}"', { n: res.saved, total: res.total, dir: res.dir }));
      }
    } catch (e) {
      setError(t('Erro ao salvar anexos: {msg}', { msg: (e as Error).message }));
    }
  }, [detail, detailRef]);

  /** Salvar anexo da mensagem aberta em janela própria. */
  const saveOpenMsgAttachment = useCallback(
    async (attachIndex: number) => {
      if (!openMsg) return;
      try {
        await window.outlook.saveAttachment(
          openMsg.sourceId,
          openMsg.folderId,
          openMsg.detail.index,
          attachIndex,
          openMsg.detail.stableId,
        );
      } catch (e) {
        setError(t('Erro ao salvar anexo: {msg}', { msg: (e as Error).message }));
      }
    },
    [openMsg],
  );

  /** Abrir anexo sem salvar — janela própria. */
  const openOpenMsgAttachment = useCallback(
    async (attachIndex: number) => {
      if (!openMsg) return;
      try {
        await window.outlook.openAttachment(
          openMsg.sourceId,
          openMsg.folderId,
          openMsg.detail.index,
          attachIndex,
          openMsg.detail.stableId,
        );
      } catch (e) {
        setError(t('Erro ao abrir anexo: {msg}', { msg: (e as Error).message }));
      }
    },
    [openMsg],
  );

  /** "Salvar todos" — janela própria. */
  const saveAllOpenMsgAttachments = useCallback(async () => {
    if (!openMsg) return;
    try {
      const res = await window.outlook.saveAllAttachments(
        openMsg.sourceId,
        openMsg.folderId,
        openMsg.detail.index,
        openMsg.detail.stableId,
      );
      if (res && res.saved < res.total) {
        setError(t('{n} de {total} anexos salvos em "{dir}"', { n: res.saved, total: res.total, dir: res.dir }));
      }
    } catch (e) {
      setError(t('Erro ao salvar anexos: {msg}', { msg: (e as Error).message }));
    }
  }, [openMsg]);

  /** Duplo clique numa mensagem → abre em janela própria (modal grande).
   *  Na pasta Rascunhos o duplo clique reabre a janela de composição — como
   *  no Outlook, que volta a editar o rascunho em vez de só lê-lo. */
  const openMessageWindow = useCallback(
    async (item: ItemSummary, sourceId?: string, folderId?: string) => {
      const sid = sourceId ?? selected?.sourceId;
      const fid = folderId ?? selected?.folder.id;
      if (sid === undefined || fid === undefined) return;
      const node = findNode(folderTrees[sid] || [], fid) ?? selected?.folder;
      const arcName = fid.startsWith('arc:') ? fid.slice(4).split('/').pop() || '' : '';
      const isDrafts =
        folderKindOf(node?.name || '') === 'drafts' ||
        folderKindOf(node?.archiveName || '') === 'drafts' ||
        folderKindOf(arcName) === 'drafts';
      if (isDrafts) {
        try {
          const payload = await window.outlook.openDraft(sid, fid, item.index, item.stableId);
          if (payload) {
            void selectItem(item, sid, fid); // marca na lista também
            openDraft({ bodyText: '', ...payload } as ComposeDraft);
            return;
          }
        } catch {
          // falha ao montar o rascunho → abre como mensagem normal
        }
      }
      // Seleciona na lista também (marca lida + mostra no painel)
      void selectItem(item, sid, fid);
      const key = detailKey(sid, fid, item.index, item.stableId);
      const cached = detailCache.current.get(key);
      if (cached) {
        setOpenMsg({ detail: cached, sourceId: sid, folderId: fid });
        return;
      }
      try {
        const d = await window.outlook.getItem(sid, fid, item.index, item.stableId);
        detailCache.current.set(key, d);
        setOpenMsg({ detail: d, sourceId: sid, folderId: fid });
      } catch (e) {
        setError(t('Erro ao abrir mensagem: {msg}', { msg: (e as Error).message }));
      }
    },
    [selected, selectItem, folderTrees],
  );

  // Lembrete: abrir o compromisso no Calendário
  const openReminderItem = useCallback(async (r: ReminderEvent) => {
    const tree = folderTrees[r.sourceId];
    const node = tree ? findNode(tree, r.folderId) : null;
    if (!node) return;
    setModule('calendar');
    await loadFolder(r.sourceId, node);
    const d = await window.outlook.getItem(r.sourceId, r.folderId, r.index, r.stableId);
    if (d) setDetail(d);
    dismissReminderKey(r.key);
    void window.outlook.snoozeReminder?.(r.key, 0);
  }, [folderTrees, loadFolder]);

  const dismissReminderKey = useCallback((key: string) => {
    setReminders((prev) => prev.filter((x) => x.key !== key));
    // minutes=0 → "dispensar" persistente no main (não volta após reabrir)
    void window.outlook.snoozeReminder?.(key, 0);
  }, []);

  const dismissAllReminders = useCallback(() => {
    for (const r of reminders) void window.outlook.snoozeReminder?.(r.key, 0);
    setReminders([]);
  }, [reminders]);

  const snoozeReminders = useCallback((minutes: number) => {
    for (const r of reminders) void window.outlook.snoozeReminder?.(r.key, minutes);
    setReminders([]);
    setNotice(t('Lembrete adiado por {n} min', { n: minutes }));
  }, [reminders]);

  const refresh = useCallback(async () => {
    if (selected) await loadFolder(selected.sourceId, selected.folder);
  }, [selected, loadFolder]);

  const onAccountsChanged = useCallback(async () => {
    await refreshSources();
    refreshAccounts();
  }, [refreshSources, refreshAccounts]);

  // Enviar/Receber roda em segundo plano: sem overlay de tela inteira. O
  // progresso aparece na barra de status; a pasta aberta é revalidada em
  // silêncio pelo evento sync:done (sem limpar seleção nem mensagem aberta).
  const sendReceive = useCallback(async () => {
    setSyncTasks([]); // nova rodada → zera a lista de tarefas do diálogo
    try {
      const results = await window.outlook.syncAll();
      const errs = results.filter((r) => r.error).map((r) => r.error).join('; ');
      // Sem aviso de "nenhuma mensagem encontrada" — o indicador de tarefas já
      // mostra o progresso; aqui só aparece quando algo deu errado.
      if (errs) setError(t('Sincronização parcial: {errs}', { errs }));
      // recarrega árvores: novas pastas locais de arquivamento podem ter surgido
      await refreshSources();
    } catch (e) {
      setError(t('Erro no envio/recebimento: {msg}', { msg: (e as Error).message }));
    }
  }, [refreshSources]);

  const moveItem = useCallback(
    async (payload: ItemDragPayload, destSourceId: string, destFolder: FolderNode) => {
      const indexes = (payload.indexes ?? [payload.index]).slice().sort((a, b) => b - a);
      const stIds = stableIdsOf(indexes);
      setLoading(
        indexes.length > 1
          ? `Movendo ${indexes.length} itens para "${destFolder.name}"…`
          : `Movendo para "${destFolder.name}"…`,
      );
      try {
        let copied = 0;
        let moved = 0;
        for (const [i, ix] of indexes.entries()) {
          const res = await window.outlook.moveItem(
            payload.sourceId,
            payload.folderId,
            ix,
            destSourceId,
            destFolder.id,
            stIds[i],
          );
          if (res === 'copiado') copied++;
          else moved++;
        }
        if (copied > 0) {
          setError(
            t('{n} item(ns) copiado(s) — o arquivo PST é somente leitura, o original foi preservado.', { n: copied }),
          );
        } else {
          setError('');
        }
        if (
          detailRef?.sourceId === payload.sourceId &&
          detailRef?.folderId === payload.folderId &&
          detail &&
          indexes.includes(detail.index)
        ) {
          setDetail(null);
        }
        await refreshSources();
        if (selected) await loadFolder(selected.sourceId, selected.folder);
      } catch (e) {
        setError(t('Erro ao mover: {msg}', { msg: (e as Error).message }));
      } finally {
        setLoading('');
      }
    },
    [selected, loadFolder, refreshSources, detail, detailRef],
  );

  /**
   * Índices-alvo das ações em lote: a multi-seleção da pasta aberta quando o
   * item em detalhe faz parte dela; senão só o item aberto (ex.: resultado
   * de busca em outra pasta).
   */
  const targetIndexes = useCallback((): number[] => {
    if (
      detail &&
      detailRef &&
      selected &&
      detailRef.sourceId === selected.sourceId &&
      detailRef.folderId === selected.folder.id &&
      selIndexes.size > 1
    ) {
      return [...selIndexes].sort((a, b) => b - a);
    }
    return detail ? [detail.index] : [];
  }, [detail, detailRef, selected, selIndexes]);

  /**
   * stableIds paralelos a uma lista de índices — resolvem o item pela
   * identidade estável mesmo que a pasta tenha recebido/perdido itens
   * depois da listagem (índice posicional desloca). Itens fora da lista
   * em memória ficam undefined → o backend cai no índice, como antes.
   */
  const stableIdsOf = useCallback((indexes: number[]): (string | undefined)[] => {
    const byIdx = new Map(itemsRef.current.map((m) => [m.index, m.stableId]));
    const d = detailRefState.current.detail;
    return indexes.map((ix) => byIdx.get(ix) ?? (d?.index === ix ? d.stableId : undefined));
  }, []);

  /** Move os itens selecionados para uma pasta por nome (Lixo eletrônico, Arquivo Morto…). */
  const moveItemTo = useCallback(
    async (folderName: string) => {
      if (!detail || !detailRef) return;
      const indexes = targetIndexes();
      if (indexes.length === 0) return;
      // Pasta padrão (lixo, arquivo…) → id da pasta EXISTENTE na base, em
      // qualquer idioma; só cai no nome localizado quando ainda não existe.
      const kind = folderKindOf(folderName);
      const tree = folderTrees[detailRef.sourceId] || [];
      const dstId = kind ? arcFolderId(tree, kind, lang) : `arc:${folderName}`;
      const dstLabel = kind ? stdFolderName(kind, lang) : folderName;
      setLoading(
        indexes.length > 1
          ? t('Movendo {n} itens para "{pasta}"…', { n: indexes.length, pasta: dstLabel })
          : t('Movendo para "{pasta}"…', { pasta: dstLabel }),
      );
      try {
        let copied = 0;
        const stIds = stableIdsOf(indexes);
        for (const [i, ix] of indexes.entries()) {
          const res = await window.outlook.moveItem(
            detailRef.sourceId,
            detailRef.folderId,
            ix,
            detailRef.sourceId,
            dstId,
            stIds[i],
          );
          if (res === 'copiado') copied++;
        }
        if (copied > 0) {
          setError(
            t('{n} item(ns) copiado(s) — o arquivo PST é somente leitura, o original foi preservado.', { n: copied }),
          );
        } else {
          setError('');
        }
        setDetail(null);
        await refreshSources();
        if (selected) await loadFolder(selected.sourceId, selected.folder);
      } catch (e) {
        setError(t('Erro ao mover: {msg}', { msg: (e as Error).message }));
      } finally {
        setLoading('');
      }
    },
    [detail, detailRef, selected, selIndexes, loadFolder, refreshSources, targetIndexes, folderTrees, lang],
  );

  /**
   * Exclui os itens selecionados — locais removem de verdade (em ordem
   * decrescente de índice, p/ não deslocar os demais), itens do PST são
   * copiados para Itens Excluídos.
   */
  /** A pasta aberta já é a lixeira? (id arc: ou nome da pasta do PST) */
  const isTrashFolder = (folderId: string, node?: FolderNode | null) =>
    (folderId.startsWith('arc:') && folderKindOf(folderId.slice(4).split('/').pop() || '') === 'deleted') ||
    folderKindOf(node?.name || '') === 'deleted' ||
    folderKindOf(node?.archiveName || '') === 'deleted' ||
    /itens exclu|deleted items|itens elimin|lixeira/i.test(
      `${folderId} ${node?.name ?? ''} ${node?.archiveName ?? ''}`,
    );

  /** Exclusões em sequência são serializadas: enquanto uma ainda corre,
   *  a próxima só é enfileirada. Sem isso, um segundo Delete rápido usava
   *  o índice do item JÁ APAGADO (detail.index ainda não tinha virado o
   *  sucessor) e excluía a mensagem errada — e, quando falhava, a seleção
   *  "travava" e os deletes seguintes não faziam mais nada. */
  const delBusy = useRef(false);
  const delAgain = useRef(false);
  const doDeleteRef = useRef<() => void>(() => {});

  /** Exclui os itens de `detailRef` — chamado após a confirmação. */
  const doDeleteItem = useCallback(async () => {
    doDeleteRef.current = () => void doDeleteItem();
    if (delBusy.current) {
      delAgain.current = true;
      return;
    }
    if (!detail || !detailRef) return;
    const indexes = targetIndexes();
    if (indexes.length === 0) return;
    delBusy.current = true;
    // Na lixeira não há "mover para Itens Excluídos" — item local apaga de
    // vez; item gravado no PST não pode ser removido (somente leitura).
    const inTrash = isTrashFolder(
      detailRef.folderId,
      findNode(folderTrees[detailRef.sourceId] || [], detailRef.folderId) ?? selected?.folder,
    );
    // Otimista: os itens saem da lista na hora — sem overlay de tela inteira.
    // A recarga em segundo plano (abaixo) corrige índices e badges.
    const inOpenFolder =
      selected &&
      detailRef.sourceId === selected.sourceId &&
      detailRef.folderId === selected.folder.id;
    // Posição (na ordem exibida) do item que vai assumir o lugar do excluído
    // + fingerprint dele — depois do reload o localizamos pelo fingerprint e
    // selecionamos com o índice NOVO. Não dá para usar o index como posição:
    // em pastas PST mescladas o index = arcCount + posição bruta, e itens
    // "excluídos" ficam ocultos mantendo buracos na numeração — usar o index
    // como posição pularia itens a cada exclusão.
    let nxtItem: ItemSummary | null = null;
    let remaining: ItemSummary[] = [];
    if (inOpenFolder) {
      const gone = new Set(indexes);
      // Sucessor pela ORDEM EXIBIDA (agrupada/ordenada/filtrada) — `items` tem
      // ordem de backend (locais depois PST), que difere da tela e fazia a
      // seleção pular itens, principalmente nas divisórias de data. O ref só
      // vale se a chave confere com ESTA pasta — durante a troca de pasta ele
      // ainda pode refletir a lista anterior e o sucessor sairia errado.
      const dRef = displayedItemsRef.current;
      const disp =
        dRef.list.length && dRef.key === `${detailRef.sourceId}|${detailRef.folderId}`
          ? dRef.list
          : items;
      const nextPos = disp.findIndex((it) => gone.has(it.index));
      const remainingDisp = disp.filter((it) => !gone.has(it.index));
      nxtItem =
        (nextPos >= 0 ? remainingDisp[Math.min(nextPos, remainingDisp.length - 1)] : null) ??
        remainingDisp[0] ??
        null;
      remaining = items.filter((it) => !gone.has(it.index));
      setItems(remaining);
      setItemCount((c) => Math.max(0, c - indexes.length));
      if (nxtItem) {
        // PRÉ-SELECIONA o sucessor já: os índices atuais ainda valem em disco
        // (o IPC ainda não compactou), então o painel troca direto para a
        // mensagem de baixo — sem tela branca ("piscada") entre o delete e
        // a seleção, e o usuário pode ir deletando em sequência.
        setSelIndexes(new Set([nxtItem.index]));
        void selectItem(nxtItem, detailRef.sourceId, detailRef.folderId);
      } else {
        setSelIndexes(new Set());
        setDetail(null);
      }
    } else {
      setDetail(null);
    }
    try {
      // Fora da lixeira TUDO vai para "Itens Excluídos" — local move o .eml,
      // item do PST é copiado e ocultado (o .pst é somente leitura).
      // Na lixeira: apaga o .eml local e oculta o item do PST.
      // Em lote (uma chamada IPC): estados em disco gravados uma vez no fim —
      // centenas de itens de uma vez em vez de uma gravação por item.
      const stIds = stableIdsOf(indexes);
      if (inTrash) {
        await window.outlook.deleteItems(detailRef.sourceId, detailRef.folderId, indexes, stIds);
      } else {
        await window.outlook.moveItems(
          detailRef.sourceId,
          detailRef.folderId,
          indexes,
          detailRef.sourceId,
          arcFolderId(folderTrees[detailRef.sourceId] || [], 'deleted', lang),
          stIds,
        );
      }
      setError('');
    } catch (e) {
      setError(t('Erro ao excluir: {msg}', { msg: (e as Error).message }));
    } finally {
      void refreshSources();
      if (selected && inOpenFolder) {
        const sid = detailRef.sourceId;
        const fid = detailRef.folderId;
        // Re-indexação LOCAL, sem reload da pasta (era o "loading" a cada
        // exclusão): arquivos .eml compactam as posições; itens do PST
        // mantêm a posição bruta (buracos dos ocultos) e só deslocam junto
        // com o novo arcCount da parte local.
        const oldArc = items.filter((m) => m.local).length;
        const newArc = remaining.filter((m) => m.local).length;
        let li = 0;
        // Mapa índice antigo → novo: o sucessor escolhido na ordem exibida é
        // encontrado com exatidão aqui — sem fingerprint (duplicatas de
        // assunto/remetente/data confundiam o find) e sem usar posição da
        // ordem visual como posição na ordem de backend (com as divisórias
        // de data misturando .eml e PST, caía no item errado).
        const newIdx = new Map<number, number>();
        const reindexed = remaining.map((m) => {
          const ni = m.local ? li++ : newArc + (m.index - oldArc);
          newIdx.set(m.index, ni);
          return { ...m, index: ni };
        });
        detailCache.current.clear(); // índices mudaram — cache por índice é inválido
        setItems(reindexed);
        // Seleciona o sucessor na hora — o detalhe já estava sendo exibido
        // pela pré-seleção; esta chamada só "firma" o índice novo e a ordem
        // é garantida pelo selSeq dentro de selectItem.
        const next =
          (nxtItem
            ? reindexed.find((it) => it.index === newIdx.get(nxtItem.index))
            : undefined) ?? reindexed[0];
        if (next) {
          setSelIndexes(new Set([next.index]));
          // Aguarda o detalhe resolver — só assim detail.index vira o índice
          // NOVO do sucessor e a exclusão enfileirada (delAgain) mira certo.
          await selectItem(next, sid, fid);
        }
        // Conferida de segurança: se a contagem TOTAL divergir do esperado
        // (ex.: algo entregou mensagem no meio da exclusão), um reload
        // silencioso corrige os índices sem piscar a lista. Compara com o
        // total esperado — `reindexed` é só a página carregada, não a pasta
        // inteira (comparar com ela disparava reload a cada exclusão e,
        // se o usuário já tivesse trocado de pasta, sobrescrevia a view).
        const expectedTotal = Math.max(0, itemCount - indexes.length);
        void window.outlook
          .getItemCount(sid, fid)
          .then((cnt) => {
            if (cnt === expectedTotal) return;
            itemsCache.current.delete(`${sid}|${fid}`);
            const cur = selectedRef.current;
            if (cur && cur.sourceId === sid && cur.folder.id === fid) {
              void loadFolder(sid, cur.folder, true);
            }
          })
          .catch(() => {});
      }
    }
    delBusy.current = false;
    if (delAgain.current) {
      delAgain.current = false;
      setTimeout(() => void doDeleteRef.current(), 0);
    }
  }, [detail, detailRef, selected, items, itemCount, loadFolder, refreshSources, targetIndexes, selectItem, folderTrees, lang]);

  /**
   * Exclui a(s) mensagem(ns) selecionada(s). Fora da lixeira move direto para
   * "Itens Excluídos" (recuperável — sem pergunta, como o Outlook); dentro da
   * lixeira a exclusão é permanente — aí sim pede confirmação.
   */
  const deleteItem = useCallback(() => {
    if (!detail || !detailRef) return;
    const n = targetIndexes().length;
    if (n === 0) return;
    const inTrash = isTrashFolder(
      detailRef.folderId,
      findNode(folderTrees[detailRef.sourceId] || [], detailRef.folderId) ?? selected?.folder,
    );
    if (!inTrash) {
      void doDeleteItem();
      return;
    }
    setConfirm({
      title: 'Excluir mensagem',
      message:
        n > 1
          ? `Excluir permanentemente ${n} itens? Eles não poderão ser recuperados.`
          : 'Excluir permanentemente este item? Ele não poderá ser recuperado.',
      onOk: () => {
        setConfirm(null);
        void doDeleteItem();
      },
    });
  }, [detail, detailRef, selected, targetIndexes, doDeleteItem, folderTrees]);
  const moveFolder = useCallback(
    async (sourceId: string, folderId: string, newParentId: string | null, destSourceId?: string) => {
      setLoading(t('Movendo pasta…'));
      try {
        const trees = await window.outlook.moveFolder(sourceId, folderId, newParentId, destSourceId);
        // O retorno traz a árvore da fonte de origem E da de destino
        // (mover entre bases altera as duas).
        setFolderTrees((t) => ({ ...t, ...trees }));
        // A pasta selecionada mudou de id — limpa a seleção se ela estava dentro do que foi movido
        if (
          selected &&
          selected.sourceId === sourceId &&
          (selected.folder.id === folderId || selected.folder.id.startsWith(`${folderId}/`))
        ) {
          setSelected(null);
          setItems([]);
          setDetail(null);
        }
      } catch (e) {
        setError(t('Erro ao mover pasta: {msg}', { msg: (e as Error).message }));
      } finally {
        setLoading('');
      }
    },
    [selected],
  );

  /** Reordena pastas irmãs (persistido — funciona inclusive em PSTs somente leitura). */
  const reorderFolder = useCallback(
    async (sourceId: string, parentId: string, folderId: string, targetId: string, after: boolean) => {
      const rawTree = folderTrees[sourceId] || [];
      const parentNode = parentId === '' ? null : findNode(rawTree, parentId);
      const siblings = applyFolderOrder(parentNode?.children || rawTree, sourceId, parentId, folderOrders);
      const ids = siblings.map((n) => n.id).filter((id) => id !== folderId);
      let ti = ids.indexOf(targetId);
      if (ti === -1) ti = ids.length;
      ids.splice(after ? ti + 1 : ti, 0, folderId);
      const key = `${sourceId}|${parentId}`;
      setFolderOrders((o) => ({ ...o, [key]: ids }));
      await window.outlook.setFolderOrder(key, ids).catch(() => {});
    },
    [folderTrees, folderOrders],
  );

  /** Reordena as bases (arquivos de dados) arrastando o cabeçalho — persistido. */
  const reorderSource = useCallback(
    async (dragId: string, targetId: string, after: boolean) => {
      const ids = applySourceOrder(sources, folderOrders)
        .map((s) => s.id)
        .filter((id) => id !== dragId);
      let ti = ids.indexOf(targetId);
      if (ti === -1) ti = ids.length;
      ids.splice(after ? ti + 1 : ti, 0, dragId);
      setFolderOrders((o) => ({ ...o, sources: ids }));
      await window.outlook.setFolderOrder('sources', ids).catch(() => {});
    },
    [sources, folderOrders],
  );

  /** Nó pai + irmãs (na ordem exibida) de uma pasta. parent null = raiz. */
  const siblingsOf = useCallback(
    (sourceId: string, folderId: string): { parentId: string; siblings: FolderNode[] } | null => {
      const raw = folderTrees[sourceId] || [];
      const find = (nodes: FolderNode[], parentId: string): { parentId: string; siblings: FolderNode[] } | null => {
        const ordered = applyFolderOrder(nodes, sourceId, parentId, folderOrders);
        if (ordered.some((n) => n.id === folderId)) return { parentId, siblings: ordered };
        for (const n of ordered) {
          const r = find(n.children, n.id);
          if (r) return r;
        }
        return null;
      };
      return find(raw, '');
    },
    [folderTrees, folderOrders],
  );

  /** Mover pasta uma posição para cima/baixo entre as irmãs. */
  const moveFolderStep = useCallback(
    async (sourceId: string, folderId: string, dir: 'up' | 'down') => {
      const ctx = siblingsOf(sourceId, folderId);
      if (!ctx) return;
      const ids = ctx.siblings.map((n) => n.id);
      const i = ids.indexOf(folderId);
      const j = dir === 'up' ? i - 1 : i + 1;
      if (i === -1 || j < 0 || j >= ids.length) return;
      [ids[i], ids[j]] = [ids[j], ids[i]];
      const key = `${sourceId}|${ctx.parentId}`;
      setFolderOrders((o) => ({ ...o, [key]: ids }));
      await window.outlook.setFolderOrder(key, ids).catch(() => {});
    },
    [siblingsOf],
  );

  /** Classifica as subpastas de uma pasta em ordem alfabética. */
  const sortChildren = useCallback(
    async (sourceId: string, parentId: string) => {
      const raw = folderTrees[sourceId] || [];
      const parent = parentId === '' ? null : findNode(raw, parentId);
      const kids = applyFolderOrder(parent ? parent.children : raw, sourceId, parentId, folderOrders);
      const ids = [...kids]
        .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'))
        .map((n) => n.id);
      const key = `${sourceId}|${parentId}`;
      setFolderOrders((o) => ({ ...o, [key]: ids }));
      await window.outlook.setFolderOrder(key, ids).catch(() => {});
    },
    [folderTrees, folderOrders],
  );

  /** Favoritos — "sourceId|folderId" persistido no localStorage. */
  const [favorites, setFavorites] = useState<Set<string>>(() => {
    try {
      return new Set(JSON.parse(localStorage.getItem('folder.favorites') || '[]'));
    } catch {
      return new Set();
    }
  });
  const toggleFavorite = useCallback((sourceId: string, folderId: string) => {
    setFavorites((prev) => {
      const next = new Set(prev);
      const key = `${sourceId}|${folderId}`;
      if (next.has(key)) next.delete(key);
      else next.add(key);
      localStorage.setItem('folder.favorites', JSON.stringify([...next]));
      return next;
    });
  }, []);

  /** Atualiza a árvore de uma base e limpa seleção se a pasta afetada sumiu. */
  const applyTree = useCallback(
    (sourceId: string, tree: FolderNode[], removedId?: string) => {
      setFolderTrees((t) => ({ ...t, [sourceId]: tree }));
      if (
        removedId &&
        selected?.sourceId === sourceId &&
        (selected.folder.id === removedId || selected.folder.id.startsWith(`${removedId}/`))
      ) {
        setSelected(null);
        setItems([]);
        setDetail(null);
      }
    },
    [selected],
  );

  const createFolder = useCallback(
    async (sourceId: string, parentId: string | null, name: string) => {
      try {
        applyTree(sourceId, await window.outlook.createFolder(sourceId, name, parentId));
      } catch (e) {
        setError(t('Erro ao criar pasta: {msg}', { msg: (e as Error).message }));
      }
    },
    [applyTree],
  );

  /** ItemDetail → NewSpecialItem — preenche o diálogo em modo edição. */
  const detailToNewItem = useCallback((d: ItemDetail): NewSpecialItem => {
    const toLocal = (iso?: string | null, dateOnly?: boolean): string => {
      if (!iso) return '';
      const dt = new Date(iso);
      if (isNaN(dt.getTime())) return '';
      const p = (n: number) => String(n).padStart(2, '0');
      const date = `${dt.getFullYear()}-${p(dt.getMonth() + 1)}-${p(dt.getDate())}`;
      return dateOnly ? date : `${date}T${p(dt.getHours())}:${p(dt.getMinutes())}`;
    };
    if (d.messageClass.startsWith('IPM.Appointment') || d.messageClass.startsWith('IPM.Schedule.Meeting')) {
      return {
        kind: 'appt',
        subject: d.subject,
        body: d.body || '',
        start: toLocal(d.startTime, d.allDay),
        end: toLocal(d.endTime, d.allDay),
        location: d.location || '',
        attendees: d.requiredAttendees || d.allAttendees || '',
        organizer: d.organizer || '',
        reminderMin: d.reminderSet ? (d.reminderMinutes ?? 15) : null,
        allDay: !!d.allDay,
      };
    }
    if (d.messageClass === 'IPM.Contact' && d.contact) {
      const c = d.contact;
      return {
        kind: 'contact',
        subject: c.displayName || d.subject,
        body: c.notes || '',
        contactName: c.displayName,
        givenName: c.givenName,
        surname: c.surname,
        company: c.company,
        jobTitle: c.jobTitle,
        dept: c.department,
        email: c.email1Address || c.email1,
        email2: c.email2Address || c.email2,
        phone: c.businessPhone,
        mobile: c.mobilePhone,
        birthday: toLocal(c.birthday, true),
      };
    }
    const tk = d.task;
    return {
      kind: 'task',
      subject: d.subject,
      body: tk?.body || d.body || '',
      startDate: toLocal(tk?.startDate, true),
      dueDate: toLocal(tk?.dueDate, true),
      percentComplete: tk?.percentComplete ?? 0,
      isComplete: !!tk?.isComplete,
      owner: tk?.owner || '',
    };
  }, []);

  /** O que a área de edição recebe — o diálogo chama isso ao salvar a edição. */
  const updateSpecialItem = useCallback(
    async (
      ref: { sourceId: string; folderId: string; index: number; stableId?: string },
      item: NewSpecialItem,
    ) => {
      const res = await window.outlook.updateSpecialItem(
        ref.sourceId,
        ref.folderId,
        ref.index,
        item,
        ref.stableId,
      );
      applyTree(ref.sourceId, res.tree);
      setEditingItem(null);
      // Detalhe em cache ficou velho — recarrega do disco.
      detailCache.current.delete(detailKey(ref.sourceId, ref.folderId, ref.index, ref.stableId));
      try {
        const fresh = await window.outlook.getItem(ref.sourceId, ref.folderId, ref.index, ref.stableId);
        setDetail(fresh);
      } catch {
        setDetail(null);
      }
      if (res.invite) {
        if (res.invite.sent > 0) {
          setNotice(t('Convite atualizado enviado para {n} participante(s).', { n: res.invite.sent }));
        } else if (res.invite.noAccount) {
          setError(
            t('Compromisso salvo, mas nenhuma conta de email está configurada para enviar os convites.'),
          );
        } else {
          setError(
            t('Compromisso salvo, mas falha ao enviar convites: {msg}', {
              msg: res.invite.error || t('erro desconhecido'),
            }),
          );
        }
      }
    },
    [applyTree, detailKey],
  );

  /** Abre o diálogo de edição — ReadingPane e inspetor do calendário chamam. */
  const openItemEditor = useCallback(
    (
      ref: { sourceId: string; folderId: string; index: number; stableId?: string },
      d: ItemDetail,
    ) => {
      setEditingItem({ ...ref, item: detailToNewItem(d) });
    },
    [detailToNewItem],
  );

  /** Botão Editar do item aberto — só itens .eml criados nesta base (X-UM-*). */
  const editDetailCb =
    detail?.local && detailRef && /IPM\.(Appointment|Contact|Task)/.test(detail.messageClass)
      ? () =>
          openItemEditor(
            {
              sourceId: detailRef.sourceId,
              folderId: detailRef.folderId,
              index: detail.index,
              stableId: detail.stableId,
            },
            detail,
          )
      : undefined;

  /** Cria compromisso/contato/tarefa e abre a pasta do item na base escolhida. */
  const saveSpecialItem = useCallback(
    async (sourceId: string, item: NewSpecialItem) => {
      const res = await window.outlook.createSpecialItem(sourceId, item);
      const tree = res.tree;
      applyTree(sourceId, tree);
      setNewItemKind(null);
      // Convites de reunião — informa o resultado do envio iMIP.
      if (res.invite) {
        if (res.invite.sent > 0) {
          setNotice(t('Convite enviado para {n} participante(s).', { n: res.invite.sent }));
        } else if (res.invite.noAccount) {
          setError(
            t('Compromisso salvo, mas nenhuma conta de email está configurada para enviar os convites.'),
          );
        } else {
          setError(
            t('Compromisso salvo, mas falha ao enviar convites: {msg}', {
              msg: res.invite.error || t('erro desconhecido'),
            }),
          );
        }
      }
      // Abre a pasta do item criado (Calendário/Contatos/Tarefas) e o módulo certo
      const kind =
        item.kind === 'appt' ? 'calendar' : item.kind === 'contact' ? 'contacts' : 'tasks';
      const mod: ModuleId =
        item.kind === 'appt' ? 'calendar' : item.kind === 'contact' ? 'people' : 'tasks';
      const cc = containerClassForModule(mod);
      setModule(mod);
      const find = (ns: FolderNode[]): FolderNode | null => {
        for (const n of ns) {
          if (folderKindOf(n.name) === kind || n.containerClass === cc) return n;
          const f = find(n.children);
          if (f) return f;
        }
        return null;
      };
      const folder = find(tree);
      setDetail(null);
      if (folder) await loadFolder(sourceId, folder);
    },
    [applyTree, loadFolder],
  );

  const renameFolder = useCallback(
    async (sourceId: string, folderId: string, name: string) => {
      try {
        applyTree(sourceId, await window.outlook.renameFolder(sourceId, folderId, name), folderId);
      } catch (e) {
        setError(t('Erro ao renomear: {msg}', { msg: (e as Error).message }));
      }
    },
    [applyTree],
  );

  /** Exclui a pasta de fato — chamado após confirmação. */
  const doDeleteFolder = useCallback(
    async (sourceId: string, folderId: string) => {
      try {
        applyTree(sourceId, await window.outlook.deleteFolder(sourceId, folderId), folderId);
      } catch (e) {
        setError(t('Erro ao excluir pasta: {msg}', { msg: (e as Error).message }));
      }
    },
    [applyTree],
  );

  const deleteFolder = useCallback(
    (sourceId: string, folderId: string) => {
      // Pastas padrão são necessárias ao app (ex.: exclusão usa "Itens
      // Excluídos") — bloqueia aqui e não só no menu de contexto.
      if (isDefaultFolderId(folderId)) {
        setError(t('"{folder}" é uma pasta padrão e não pode ser excluída.', { folder: folderId.slice(4) }));
        return;
      }
      const node = findNode(folderTrees[sourceId] || [], folderId);
      // Conteúdo que será apagado — o aviso fica concreto, como no diálogo
      // de exclusão do Outlook
      const countItems = (n: FolderNode): number =>
        (n.contentCount || 0) + n.children.reduce((acc, c) => acc + countItems(c), 0);
      const countFolders = (n: FolderNode): number =>
        n.children.length + n.children.reduce((acc, c) => acc + countFolders(c), 0);
      const nItems = node ? countItems(node) : 0;
      const nSubs = node ? countFolders(node) : 0;
      const details = [
        nItems > 0 ? `${nItems} ${nItems === 1 ? 'item' : 'itens'}` : '',
        nSubs > 0 ? `${nSubs} subpasta${nSubs === 1 ? '' : 's'}` : '',
      ]
        .filter(Boolean)
        .join(' e ');
      setConfirm({
        title: 'Excluir pasta',
        message:
          `Excluir a pasta "${node?.name || folderId}"` +
          (details ? ` e todo o seu conteúdo (${details})` : '') +
          '? Esta ação não pode ser desfeita.',
        onOk: () => {
          setConfirm(null);
          void doDeleteFolder(sourceId, folderId);
        },
      });
    },
    [doDeleteFolder, folderTrees],
  );

  /** Oculta/reexibe uma pasta — não apaga nada, só some da árvore. */
  const setFolderHidden = useCallback(
    async (sourceId: string, folderId: string, hidden: boolean) => {
      try {
        applyTree(sourceId, await window.outlook.setFolderHidden(sourceId, folderId, hidden));
      } catch (e) {
        setError(t('Erro ao ocultar pasta: {msg}', { msg: (e as Error).message }));
      }
    },
    [applyTree],
  );

  const copyFolder = useCallback(
    async (sourceId: string, folderId: string) => {
      try {
        applyTree(sourceId, await window.outlook.copyFolder(sourceId, folderId));
      } catch (e) {
        setError(t('Erro ao copiar pasta: {msg}', { msg: (e as Error).message }));
      }
    },
    [applyTree],
  );

  /** Esvazia a pasta de fato — chamado após confirmação. */
  const doEmptyFolder = useCallback(
    async (sourceId: string, folderId: string) => {
      try {
        const removed = await window.outlook.emptyFolder(sourceId, folderId);
        applyTree(sourceId, await window.outlook.getFolderTree(sourceId));
        if (selected?.sourceId === sourceId && selected.folder.id === folderId) {
          setDetail(null);
          if (removed > 0) await loadFolder(sourceId, selected.folder);
        }
        setNotice(
          removed > 0
            ? t('{n} item(ns) removido(s) da pasta.', { n: removed })
            : t('A pasta já está vazia.'),
        );
        setTimeout(() => setNotice(''), 7000);
      } catch (e) {
        setError(t('Erro ao limpar pasta: {msg}', { msg: (e as Error).message }));
      }
    },
    [applyTree, selected, loadFolder],
  );

  const emptyFolder = useCallback(
    (sourceId: string, folderId: string) => {
      const node = findNode(folderTrees[sourceId] || [], folderId);
      setConfirm({
        title: 'Esvaziar pasta',
        message: `Remover todos os itens de "${node?.name || folderId}"? Esta ação não pode ser desfeita.`,
        confirmLabel: 'Esvaziar',
        onOk: () => {
          setConfirm(null);
          void doEmptyFolder(sourceId, folderId);
        },
      });
    },
    [doEmptyFolder, folderTrees],
  );

  const markAllRead = useCallback(
    async (sourceId: string, folderId: string) => {
      try {
        await window.outlook.markAllRead(sourceId, folderId);
        // O count-check do loadFolder ignora mudanças de isRead (a contagem
        // é a mesma) — reflete "lido" nos itens em memória e no cache da
        // pasta em vez de re-baixar a lista inteira.
        const readAll = (list: ItemSummary[]) => list.map((it) => ({ ...it, isRead: true }));
        const ck = `${sourceId}|${folderId}`;
        const hit = itemsCache.current.get(ck);
        if (hit) itemsCache.current.set(ck, readAll(hit));
        if (selected?.sourceId === sourceId && selected.folder.id === folderId) {
          setItems((prev) => readAll(prev));
        }
        applyTree(sourceId, await window.outlook.getFolderTree(sourceId));
      } catch (e) {
        setError(t('Erro ao marcar como lidas: {msg}', { msg: (e as Error).message }));
      }
    },
    [applyTree, selected],
  );

  /** "Executar Regras Agora" — aplica regras habilitadas aos itens da pasta aberta. */
  /** Exporta contas+regras+assinaturas para um arquivo de backup JSON. */
  const exportBackup = useCallback(async () => {
    try {
      const path = await window.outlook.exportBackup();
      if (path) {
        setNotice(t('Backup salvo em: {path}', { path }));
        setTimeout(() => setNotice(''), 9000);
      }
    } catch (e) {
      setError(t('Erro ao exportar: {msg}', { msg: (e as Error).message }));
    }
  }, []);

  /** Importa backup de configurações e avisa se alguma conta ficou sem base. */
  const importBackup = useCallback(async () => {
    try {
      const r = await window.outlook.importBackup();
      if (!r) return;
      const parts = [
        `${r.accounts} conta(s)`,
        `${r.rules} regra(s)`,
        `${r.signatures} assinatura(s)`,
      ];
      setNotice(t('Backup importado: {parts}.{warn}', { parts: parts.join(', '), warn: r.warnings.length ? ` ${t('Atenção: {w}', { w: r.warnings.join(' ') })}` : '' }));
      setTimeout(() => setNotice(''), r.warnings.length ? 20000 : 9000);
      await refreshSources();
      refreshAccounts();
    } catch (e) {
      setError(t('Erro ao importar backup: {msg}', { msg: (e as Error).message }));
    }
  }, [refreshSources, refreshAccounts]);

  /** Salva a mensagem aberta como .eml (Arquivo → Salvar como). */
  const saveItemAs = useCallback(async () => {
    if (!detailRef || !detail) return;
    try {
      const path = await window.outlook.saveItemAs(
        detailRef.sourceId,
        detailRef.folderId,
        detail.index,
        detail.stableId,
      );
      if (path) {
        setNotice(t('Mensagem salva em: {path}', { path }));
        setTimeout(() => setNotice(''), 9000);
      }
    } catch (e) {
      setError(t('Erro ao salvar: {msg}', { msg: (e as Error).message }));
    }
  }, [detail, detailRef]);

  /** Imprime a mensagem aberta (Arquivo → Imprimir). */
  const printItem = useCallback(async () => {
    if (!detailRef || !detail) return;
    try {
      await window.outlook.printItem(
        detailRef.sourceId,
        detailRef.folderId,
        detail.index,
        detail.stableId,
      );
    } catch (e) {
      setError(t('Erro ao imprimir: {msg}', { msg: (e as Error).message }));
    }
  }, [detail, detailRef]);

  /** Imprime a mensagem aberta na janela separada. */
  const printOpenMsg = useCallback(async () => {
    if (!openMsg) return;
    try {
      await window.outlook.printItem(
        openMsg.sourceId,
        openMsg.folderId,
        openMsg.detail.index,
        openMsg.detail.stableId,
      );
    } catch (e) {
      setError(t('Erro ao imprimir: {msg}', { msg: (e as Error).message }));
    }
  }, [openMsg]);

  /** Compromisso/contato/tarefa novo vai SEMPRE para o arquivo de dados principal. */
  const openNewItem = useCallback(
    async (kind: 'appt' | 'contact' | 'task') => {
      const main = await window.outlook.getDefaultStore().catch(() => '');
      defaultStoreRef.current = main || '';
      const mainSrc = main ? sources.find((s) => s.filePath === main) : undefined;
      if (main && !mainSrc) {
        setError(
          t('O arquivo de dados principal não está aberto — abra-o em Configurações → Arquivos de Dados.'),
        );
      }
      setApptSourceId(mainSrc?.id || '');
      setNewItemKind(kind);
    },
    [sources],
  );
  const openNewAppt = useCallback(() => openNewItem('appt'), [openNewItem]);

  const runRules = useCallback(async () => {
    if (!selected) return;
    setLoading(t('Executando regras na pasta…'));
    try {
      const r = await window.outlook.applyRulesToFolder(selected.sourceId, selected.folder.id);
      setError(
        r.applied > 0
          ? t('{n} mensagem(ns) movida(s) pelas regras.', { n: r.applied })
          : t('Nenhuma mensagem correspondeu às regras.'),
      );
      await refreshSources();
      await loadFolder(selected.sourceId, selected.folder);
    } catch (e) {
      setError(t('Erro ao executar regras: {msg}', { msg: (e as Error).message }));
    } finally {
      setLoading('');
    }
  }, [selected, loadFolder, refreshSources]);

  const syncOne = useCallback(
    async (accountId: string) => {
      setSyncTasks([]);
      try {
        const r = await window.outlook.syncAccount(accountId);
        if (r.error) setError(t('Erro: {err}', { err: r.error }));
        await refreshSources();
      } catch (e) {
        setError(t('Erro ao sincronizar: {msg}', { msg: (e as Error).message }));
      }
    },
    [refreshSources],
  );

  const defaultAccountId = useCallback(() => {
    // A conta DONA da base aberta é o jeito mais confiável: no POP3 a
    // mensagem mora numa base local que está vinculada a uma conta
    // (deliveryPath). Responder da caixa do @agenciaunit sai pelo
    // @agenciaunit — como o Outlook faz.
    const src = sources.find((s) => s.id === detailRef?.sourceId);
    if (src?.filePath) {
      const norm = (p: string) => p.replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();
      const bound = accounts.find(
        (a) => a.deliveryPath && norm(a.deliveryPath) === norm(src.filePath!),
      );
      if (bound) return bound.id;
    }
    if (detailRef?.sourceId.startsWith('acct:')) {
      const id = detailRef.sourceId.slice('acct:'.length);
      if (accounts.some((a) => a.id === id)) return id;
    }
    // Fallback para base sem conta vinculada (principal compartilhada):
    // remetente nosso (itens enviados) → aquela conta; senão a conta cujo
    // e-mail aparece no Para/Cc — a que recebeu a mensagem.
    if (detail) {
      const sender = (detail.senderEmailAddress || '').toLowerCase();
      const bySender = accounts.find(
        (a) => a.email && sender.includes(a.email.toLowerCase()),
      );
      if (bySender) return bySender.id;
      const targets = `${detail.displayTo || ''} ${detail.displayCC || ''}`.toLowerCase();
      const match = accounts.find(
        (a) => a.email && targets.includes(a.email.toLowerCase()),
      );
      if (match) return match.id;
    }
    return accounts[0]?.id || '';
  }, [detailRef, detail, accounts, sources]);

  /** Responde convite de reunião (Aceitar / Recusar / Propor novo horário).
   *  Feedback via toast inline — o envio SMTP é rápido e o overlay de tela
   *  cheia bloqueava a UI à toa. */
  const inviteBusy = useRef(false);
  const respondInvite = useCallback(
    async (response: 'accept' | 'decline' | 'counter', newStart?: string, newEnd?: string) => {
      if (!detail || !detailRef || inviteBusy.current) return;
      inviteBusy.current = true;
      setNotice(t('Enviando resposta ao organizador…'));
      try {
        const res = await window.outlook.respondInvite({
          sourceId: detailRef.sourceId,
          folderId: detailRef.folderId,
          index: detail.index,
          stableId: detail.stableId,
          accountId: defaultAccountId(),
          response,
          newStart,
          newEnd,
        });
        setError('');
        setNotice(
          response === 'accept'
            ? res?.apptSaved === false
              ? t('Resposta enviada, mas o compromisso não pôde ser registrado no Calendário — verifique a base de dados principal.')
              : t('Convite aceito — compromisso adicionado ao Calendário.')
            : response === 'decline'
              ? t('Convite recusado — resposta enviada ao organizador.')
              : t('Nova proposta de horário enviada ao organizador.'),
        );
        setTimeout(() => setNotice(''), 9000);
        // Convite respondido sai da listagem — o main já o moveu para Itens
        // Excluídos (opção padrão do Outlook "excluir após responder"). O
        // compromisso continua no Calendário/lembretes.
        const inOpen =
          selected &&
          detailRef.sourceId === selected.sourceId &&
          detailRef.folderId === selected.folder.id;
        if (inOpen) {
          const gone = detail.index;
          // Ordem exibida só se o ref ainda é desta pasta — durante a
          // troca de pasta ele pode refletir a lista anterior.
          const dRef = displayedItemsRef.current;
          const disp =
            dRef.list.length && dRef.key === `${detailRef.sourceId}|${detailRef.folderId}`
              ? dRef.list
              : items;
          const pos = disp.findIndex((it) => it.index === gone);
          const remainingDisp = disp.filter((it) => it.index !== gone);
          const nxtItem =
            (pos >= 0 ? remainingDisp[Math.min(pos, remainingDisp.length - 1)] : null) ??
            remainingDisp[0] ??
            null;
          const remaining = items.filter((it) => it.index !== gone);
          // Reindexação local — o .eml já saiu do disco, posições compactadas
          const oldArc = items.filter((m) => m.local).length;
          const newArc = remaining.filter((m) => m.local).length;
          let li = 0;
          // Mapa índice antigo → novo — sucessor determinístico.
          const newIdx = new Map<number, number>();
          const reindexed = remaining.map((m) => {
            const ni = m.local ? li++ : newArc + (m.index - oldArc);
            newIdx.set(m.index, ni);
            return { ...m, index: ni };
          });
          detailCache.current.clear();
          setItems(reindexed);
          setItemCount((c) => Math.max(0, c - 1));
          const next =
            (nxtItem
              ? reindexed.find((it) => it.index === newIdx.get(nxtItem.index))
              : undefined) ?? reindexed[0];
          if (next) {
            setSelIndexes(new Set([next.index]));
            void selectItem(next, detailRef.sourceId, detailRef.folderId);
          } else {
            setSelIndexes(new Set());
            setDetail(null);
          }
        }
        void refreshSources();
      } catch (e) {
        setError(t('Erro ao responder convite: {msg}', { msg: (e as Error).message }));
      } finally {
        inviteBusy.current = false;
      }
    },
    [detail, detailRef, defaultAccountId, refreshSources, selected, items, selectItem],
  );

  const openCompose = useCallback(() => {
    openDraft({ accountId: defaultAccountId(), kind: 'new', to: '', subject: '', bodyText: '' });
  }, [defaultAccountId]);

  const esc = (s: string) =>
    s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  /** Bloco de mensagem original estilo Outlook (usa o HTML real quando existe). */
  const quoteHtml = (d: ItemDetail) => {
    const orig = d.bodyHTML || `<div>${esc(d.body || '').replace(/\n/g, '<br>')}</div>`;
    return (
      `<br><div class="quote-block" style="border-top:1px solid #b5c4df;padding-top:10px;margin-top:6px">` +
      `<div style="font-family:'Segoe UI',sans-serif;font-size:12px;color:#1a1a1a">` +
      `<b>De:</b> ${esc(d.senderName || d.senderEmailAddress)} &lt;${esc(d.senderEmailAddress)}&gt;<br>` +
      `<b>Enviada em:</b> ${d.sentTime ? new Date(d.sentTime).toLocaleString('pt-BR') : ''}<br>` +
      `<b>Para:</b> ${esc(d.displayTo)}<br>` +
      `<b>Assunto:</b> ${esc(d.subject)}<br><br></div>${orig}</div>`
    );
  };

  /** Monta draft de resposta a partir de qualquer ItemDetail (painel ou janela). */
  const replyDraftOf = (
    d: ItemDetail,
    all: boolean,
    ref?: { sourceId: string; folderId: string; index: number; stableId?: string },
  ): ComposeDraft => {
    const ours = new Set(
      accounts.map((a) => (a.email || '').toLowerCase()).filter(Boolean),
    );
    const sender = (d.senderEmailAddress || '').toLowerCase();
    // Mensagem enviada por uma conta NOSSA (Itens Enviados, teste entre
    // nossas contas): responder "para o remetente" seria responder para
    // nós mesmos — como o Outlook, a resposta vai aos destinatários
    // originais.
    const selfSent = ours.has(sender);
    const addrs = (s: string) =>
      s.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
    const recips = addrs(`${d.displayTo || ''}; ${d.displayCC || ''}`).filter(
      (e) => !ours.has(e.toLowerCase()),
    );
    // Mensagem enviada por nós: responde aos destinatários originais. Se
    // TODOS eram nossos (teste entre nossas contas), volta ao remetente —
    // a mensagem chegou como recebida na outra caixa, então "responder"
    // é para quem enviou, senão o Para ficava vazio.
    const to =
      selfSent && recips.length
        ? recips.join('; ')
        : d.senderEmailAddress || d.senderName;
    // Responder a todos: Cc = destinatários originais que não somos nós nem
    // o remetente (ele já está no Para).
    const cc =
      all && !selfSent
        ? recips.filter((e) => e.toLowerCase() !== sender).join('; ')
        : '';
    return {
      accountId: defaultAccountId(),
      kind: 'reply',
      to,
      cc,
      subject: d.subject.toLowerCase().startsWith('re:') ? d.subject : `Re: ${d.subject}`,
      bodyText: '',
      quoteHtml: quoteHtml(d),
      replyRef: ref,
    };
  };
  const forwardDraftOf = (
    d: ItemDetail,
    ref?: { sourceId: string; folderId: string; index: number; stableId?: string },
  ): ComposeDraft => ({
    accountId: defaultAccountId(),
    kind: 'forward',
    to: '',
    subject: d.subject.toLowerCase().startsWith('fwd:') ? d.subject : `Fwd: ${d.subject}`,
    bodyText: '',
    quoteHtml: quoteHtml(d),
    replyRef: ref,
  });

  const openReply = useCallback(
    (all: boolean) => {
      if (!detail) return;
      const ref = detailRef
        ? { ...detailRef, index: detail.index, stableId: detail.stableId }
        : undefined;
      openDraft(replyDraftOf(detail, all, ref));
    },
    [detail, detailRef, defaultAccountId],
  );

  const openForward = useCallback(() => {
    if (!detail) return;
    const ref = detailRef
      ? { ...detailRef, index: detail.index, stableId: detail.stableId }
      : undefined;
    openDraft(forwardDraftOf(detail, ref));
  }, [detail, detailRef, defaultAccountId]);

  /** Move índices de uma pasta para outra (na mesma base) — usado pelo ctx menu. */
  const moveIndexesTo = useCallback(
    async (
      sourceId: string,
      folderId: string,
      indexes: number[],
      destFolderId: string,
      destSourceId = sourceId,
      stableIds?: (string | undefined)[],
    ) => {
      const inOpenFolder =
        selected && selected.sourceId === sourceId && selected.folder.id === folderId;
      // Otimista — mesmo fluxo da exclusão: os itens saem da lista na hora,
      // sem overlay de tela inteira, e a seleção vai para a mensagem seguinte.
      let nxtItem: ItemSummary | null = null;
      let remaining: ItemSummary[] = [];
      if (inOpenFolder) {
        const gone = new Set(indexes);
        const dRef = displayedItemsRef.current;
        const disp =
          dRef.list.length && dRef.key === `${sourceId}|${folderId}` ? dRef.list : items;
        const nextPos = disp.findIndex((it) => gone.has(it.index));
        const remainingDisp = disp.filter((it) => !gone.has(it.index));
        nxtItem =
          (nextPos >= 0 ? remainingDisp[Math.min(nextPos, remainingDisp.length - 1)] : null) ??
          remainingDisp[0] ??
          null;
        remaining = items.filter((it) => !gone.has(it.index));
        setItems(remaining);
        setItemCount((c) => Math.max(0, c - indexes.length));
        if (nxtItem) {
          // Pré-seleção do sucessor — mesmo fluxo da exclusão: o painel
          // troca direto para a mensagem seguinte, sem piscar.
          setSelIndexes(new Set([nxtItem.index]));
          void selectItem(nxtItem, sourceId, folderId);
        } else {
          setSelIndexes(new Set());
          if (detail && detailRef?.sourceId === sourceId && detailRef.folderId === folderId && gone.has(detail.index)) {
            setDetail(null);
          }
        }
      }
      try {
        // Em lote: uma chamada IPC, estados em disco gravados uma vez no fim
        await window.outlook.moveItems(
          sourceId,
          folderId,
          indexes,
          destSourceId,
          destFolderId,
          stableIds ?? stableIdsOf(indexes),
        );
        setError('');
      } catch (e) {
        setError(t('Erro ao mover: {msg}', { msg: (e as Error).message }));
        // Falhou — recarrega em silêncio para restaurar a lista real.
        if (selected) await loadFolder(selected.sourceId, selected.folder, true);
        return;
      }
      await refreshSources();
      if (inOpenFolder && selected) {
        // Re-indexação local (arquivos .eml compactam posições; itens PST
        // mantêm a posição bruta deslocada pelo novo arcCount) — sem reload.
        const oldArc = items.filter((m) => m.local).length;
        const newArc = remaining.filter((m) => m.local).length;
        let li = 0;
        // Mapa índice antigo → novo — sucessor determinístico.
        const newIdx = new Map<number, number>();
        const reindexed = remaining.map((m) => {
          const ni = m.local ? li++ : newArc + (m.index - oldArc);
          newIdx.set(m.index, ni);
          return { ...m, index: ni };
        });
        detailCache.current.clear();
        setItems(reindexed);
        const next =
          (nxtItem
            ? reindexed.find((it) => it.index === newIdx.get(nxtItem.index))
            : undefined) ?? reindexed[0];
        if (next) {
          setSelIndexes(new Set([next.index]));
          void selectItem(next, sourceId, folderId);
        }
        void window.outlook
          .getItemCount(sourceId, folderId)
          .then((cnt) => {
            if (cnt === reindexed.length) return;
            itemsCache.current.delete(`${sourceId}|${folderId}`);
            void loadFolder(sourceId, selected.folder, true);
          })
          .catch(() => {});
      } else if (selected) {
        // Item de outra pasta (resultado de busca) — só revalida a aberta.
        await loadFolder(selected.sourceId, selected.folder, true);
      }
    },
    [selected, items, detail, detailRef, loadFolder, refreshSources, selectItem, stableIdsOf],
  );

  /**
   * Aplica/remove uma categoria nos itens — semântica do Outlook: se TODOS
   * os selecionados já têm a marca, remove de todos; senão adiciona nos que
   * faltam. catId null limpa todas as categorias. Reflete na lista e no
   * painel sem recarregar a pasta.
   */
  const applyCategories = useCallback(
    async (sourceId: string, folderId: string, indexes: number[], catId: string | null) => {
      const api = window.outlook;
      const isCur =
        !!selected && selected.sourceId === sourceId && selected.folder.id === folderId;
      const inList = new Map(itemsRef.current.map((m) => [m.index, m]));
      const stIds = indexes.map(
        (ix) => inList.get(ix)?.stableId ?? stableIdsOf([ix])[0],
      );
      const current = new Map<number, string[]>();
      for (const [i, ix] of indexes.entries()) {
        const m = isCur ? inList.get(ix) : undefined;
        current.set(
          ix,
          m?.categories ?? (await api.getItemCategories(sourceId, folderId, ix, stIds[i])),
        );
      }
      // Agrupa por conjunto resultante — uma chamada IPC por grupo (lote)
      const groups = new Map<string, { cats: string[]; idxs: number[] }>();
      const put = (ix: number, cats: string[]) => {
        const key = cats.join('');
        const g = groups.get(key) ?? { cats, idxs: [] };
        g.idxs.push(ix);
        groups.set(key, g);
      };
      if (catId === null) {
        for (const ix of indexes) put(ix, []);
      } else {
        const allHave = indexes.every((ix) => (current.get(ix) ?? []).includes(catId));
        for (const ix of indexes) {
          const cur = current.get(ix) ?? [];
          put(ix, allHave ? cur.filter((c) => c !== catId) : [...new Set([...cur, catId])]);
        }
      }
      await Promise.all(
        [...groups.values()].map((g) =>
          api.setItemCategories(
            sourceId,
            folderId,
            g.idxs,
            g.cats,
            g.idxs.map((ix) => stIds[indexes.indexOf(ix)]),
          ),
        ),
      );
      if (!isCur) return;
      const nextByIdx = new Map<number, string[]>();
      for (const g of groups.values()) for (const ix of g.idxs) nextByIdx.set(ix, g.cats);
      setItems((prev) =>
        prev.map((m) =>
          nextByIdx.has(m.index)
            ? { ...m, categories: nextByIdx.get(m.index)!.length ? nextByIdx.get(m.index) : undefined }
            : m,
        ),
      );
      const d = detailRefState.current;
      if (d.detail && nextByIdx.has(d.detail.index)) {
        const cats = nextByIdx.get(d.detail.index)!;
        setDetail({ ...d.detail, categories: cats.length ? cats : undefined });
        if (d.ref) {
          detailCache.current.set(
            detailKey(d.ref.sourceId, d.ref.folderId, d.detail.index, d.detail.stableId),
            { ...d.detail, categories: cats.length ? cats : undefined },
          );
        }
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [selected, stableIdsOf],
  );

  /** Ações do menu de contexto da mensagem (botão direito). */
  const msgCtxAction = useCallback(
    async (
      action: string,
      ctx: {
        sourceId: string;
        folderId: string;
        index: number;
        stableId?: string;
        indexes: number[];
        stableIds?: (string | undefined)[];
        isRead: boolean;
        senders?: string[];
      },
    ) => {
      const indexes = ctx.indexes.length ? ctx.indexes : [ctx.index];
      const stIds = indexes.map((ix) => {
        const fromCtx = ix === ctx.index ? ctx.stableId : ctx.stableIds?.[ctx.indexes.indexOf(ix)];
        return fromCtx ?? stableIdsOf([ix])[0];
      });
      try {
        if (action === 'reply' || action === 'replyAll' || action === 'forward') {
          const d = await window.outlook.getItem(ctx.sourceId, ctx.folderId, ctx.index, ctx.stableId);
          const ref = {
            sourceId: ctx.sourceId,
            folderId: ctx.folderId,
            index: ctx.index,
            stableId: ctx.stableId,
          };
          if (action === 'forward') {
            openDraft({
              accountId: defaultAccountId(),
              kind: 'forward',
              to: '',
              subject: d.subject.toLowerCase().startsWith('fwd:') ? d.subject : `Fwd: ${d.subject}`,
              bodyText: '',
              quoteHtml: quoteHtml(d),
              replyRef: ref,
            });
          } else {
            openDraft({
              accountId: defaultAccountId(),
              kind: 'reply',
              to: d.senderEmailAddress || d.senderName,
              cc: action === 'replyAll' ? d.displayCC : '',
              subject: d.subject.toLowerCase().startsWith('re:') ? d.subject : `Re: ${d.subject}`,
              bodyText: '',
              quoteHtml: quoteHtml(d),
              replyRef: ref,
            });
          }
        } else if (action === 'markRead' || action === 'markUnread') {
          const read = action === 'markRead';
          // setItemRead cobre .eml locais e itens do PST (override lógico) —
          // não há gate de "somente leitura" para estado de leitura.
          for (const [i, ix] of indexes.entries()) {
            await window.outlook.setItemRead(ctx.sourceId, ctx.folderId, ix, read, stIds[i]);
          }
          const ixSet = new Set(indexes);
          // loadFolder não re-baixaria (a contagem não muda) — reflete o
          // novo isRead direto nos itens e no cache da pasta.
          const upd = (list: ItemSummary[]) =>
            list.map((it) => (ixSet.has(it.index) ? { ...it, isRead: read } : it));
          const ck = `${ctx.sourceId}|${ctx.folderId}`;
          const hit = itemsCache.current.get(ck);
          if (hit) itemsCache.current.set(ck, upd(hit));
          if (selected?.sourceId === ctx.sourceId && selected.folder.id === ctx.folderId) {
            setItems((prev) => upd(prev));
          }
          applyTree(ctx.sourceId, await window.outlook.getFolderTree(ctx.sourceId));
        } else if (action === 'junk') {
          const tree = folderTrees[ctx.sourceId] || [];
          await moveIndexesTo(
            ctx.sourceId,
            ctx.folderId,
            indexes,
            arcFolderId(tree, 'junk', lang),
            undefined,
            stIds,
          );
        } else if (action === 'blockSender') {
          // Bloqueia o(s) remetente(s) e joga a(s) mensagem(ns) no Lixo —
          // as próximas entregas caem direto lá (filtro no processo principal).
          // ctx.senders já vem resolvido pelo MessageList (funciona também em
          // resultados de busca, onde itemsRef não cobre o item).
          const addrs = new Set<string>((ctx.senders ?? []).map((a) => a.toLowerCase()));
          if (!addrs.size) {
            setError(t('Não foi possível identificar o endereço do remetente.'));
          } else {
            for (const a of addrs) await window.outlook.blockedAdd(a);
            await moveIndexesTo(
              ctx.sourceId,
              ctx.folderId,
              indexes,
              arcFolderId(folderTrees[ctx.sourceId] || [], 'junk', lang),
              undefined,
              stIds,
            );
            setNotice(
              addrs.size === 1
                ? t('"{a}" bloqueado — próximas mensagens vão direto para a pasta de lixo eletrônico.', { a: [...addrs][0] })
                : t('{n} remetentes bloqueados — próximas mensagens vão para a pasta de lixo eletrônico.', { n: addrs.size }),
            );
            setTimeout(() => setNotice(''), 8000);
          }
        } else if (action === 'archive') {
          await moveIndexesTo(
            ctx.sourceId,
            ctx.folderId,
            indexes,
            arcFolderId(folderTrees[ctx.sourceId] || [], 'archive', lang),
            undefined,
            stIds,
          );
        } else if (action === 'delete') {
          // Mesma lógica do botão Excluir: fora da lixeira move direto para
          // "Itens Excluídos" (recuperável — sem pergunta); dentro da lixeira a
          // exclusão é permanente — pede confirmação.
          const ctxNode = findNode(folderTrees[ctx.sourceId] || [], ctx.folderId);
          const inTrash = isTrashFolder(ctx.folderId, ctxNode);
          const n = indexes.length;
          const run = async () => {
            setLoading(n > 1 ? t('Excluindo {n} itens…', { n }) : t('Excluindo…'));
            try {
              // Em lote: fora da lixeira TUDO vai para "Itens Excluídos"
              // (antes itens locais eram apagados permanentemente aqui);
              // na lixeira a exclusão é permanente.
              if (inTrash) {
                await window.outlook.deleteItems(ctx.sourceId, ctx.folderId, indexes, stIds);
              } else {
                await window.outlook.moveItems(
                  ctx.sourceId,
                  ctx.folderId,
                  indexes,
                  ctx.sourceId,
                  arcFolderId(folderTrees[ctx.sourceId] || [], 'deleted', lang),
                  stIds,
                );
              }
              setError('');
              setDetail(null);
              await refreshSources();
              if (selected) await loadFolder(selected.sourceId, selected.folder);
            } catch (e) {
              setError((e as Error).message);
            } finally {
              setLoading('');
            }
          };
          if (!inTrash) {
            await run();
          } else {
            setConfirm({
              title: 'Excluir mensagem',
              message:
                n > 1
                  ? `Excluir permanentemente ${n} itens? Eles não poderão ser recuperados.`
                  : 'Excluir permanentemente este item? Ele não poderá ser recuperado.',
              onOk: () => {
                setConfirm(null);
                void run();
              },
            });
          }
        } else if (action.startsWith('move:')) {
          const sep = action.indexOf('\u0000', 5);
          const destSourceId = sep >= 0 ? action.slice(5, sep) : ctx.sourceId;
          const destFolderId = sep >= 0 ? action.slice(sep + 1) : action.slice(5);
          await moveIndexesTo(ctx.sourceId, ctx.folderId, indexes, destFolderId, destSourceId);
        } else if (action === 'rules') {
          openSettings('rules');
        } else if (action === 'catsManage') {
          setShowCatsDlg(true);
        } else if (action === 'catClear') {
          await applyCategories(ctx.sourceId, ctx.folderId, indexes, null);
        } else if (action.startsWith('cat:')) {
          await applyCategories(ctx.sourceId, ctx.folderId, indexes, action.slice(4));
        }
      } catch (e) {
        setError((e as Error).message);
        setLoading('');
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [selected, loadFolder, refreshSources, moveIndexesTo, openSettings, defaultAccountId, applyCategories, folderTrees, lang],
  );

  const entries = applySourceOrder(sources, folderOrders).map((s) => ({
    source: s,
    tree: applyFolderOrder(folderTrees[s.id] || [], s.id, '', folderOrders),
  }));
  const hasAnySource = sources.length > 0;

  /** Todas as pastas de calendário das bases abertas (para "Meus Calendários"). */
  const allCalendars = useMemo(() => {
    const list: Array<{ key: string; sourceId: string; folderId: string; name: string; sourceLabel: string }> = [];
    for (const e of entries) {
      const walk = (nodes: FolderNode[]) => {
        for (const n of nodes) {
          if (
            n.containerClass === 'IPF.Appointment' ||
            folderKindOf(n.name) === 'calendar' ||
            folderKindOf(n.archiveName || '') === 'calendar'
          ) {
            list.push({
              key: `${e.source.id}|${n.id}`,
              sourceId: e.source.id,
              folderId: n.id,
              name: n.name,
              sourceLabel: e.source.label,
            });
          }
          walk(n.children);
        }
      };
      walk(e.tree);
    }
    return list;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [folderTrees, sources]);

  /** Pasta selecionada + destinos possíveis para "Mover Pasta" (aba Pasta). */
  const selFolderProp = selected
    ? {
        sourceId: selected.sourceId,
        folderId: selected.folder.id,
        name: selected.folder.name,
        isLocal: selected.folder.id.startsWith('arc:'),
      }
    : null;
  /** Pastas de enviados não têm filtro "Não lidos" — tudo que você enviou já é lido. */
  const isSentFolder =
    !!selected &&
    (folderKindOf(selected.folder.name) === 'sent' ||
      folderKindOf(selected.folder.name) === 'outbox' ||
      folderKindOf(selected.folder.name) === 'drafts' ||
      /itens enviados|caixa de sa[ií]da|sent items|outbox|rascunhos|drafts/i.test(
        selected.folder.name,
      ));
  const folderTargets: Array<{ id: string | null; label: string }> = [];
  if (selected) {
    folderTargets.push({ id: null, label: t('（raiz da base）') });
    const flat = (nodes: FolderNode[], depth: number) => {
      for (const n of nodes) {
        if (n.id === selected.folder.id || n.id.startsWith(`${selected.folder.id}/`)) continue;
        folderTargets.push({
          id: n.id,
          label: `${'　'.repeat(depth)}${localizeFolderName(n.name, lang)}`,
        });
        flat(n.children, depth + 1);
      }
    };
    flat(folderTrees[selected.sourceId] || [], 0);
  }

  return (
    <div className={`app${winChrome ? ' with-titlebar' : ''}`}>
      {winChrome && (
        <div className="win-titlebar">
          <Icon name="envelope" size={16} />
          <span className="win-title">Unit Mail</span>
          <span className="beta-pill" title="Versão Beta — em fase de testes. Reporte problemas no GitHub.">
            Beta
          </span>
        </div>
      )}

      <Ribbon
        module={module}
        pstOpen={!!selected}
        hasAccounts={accounts.length > 0}
        canReply={module === 'mail' && !!detail && accounts.length > 0}
        accounts={accounts}
        sourcesCount={sources.length}
        rpLayout={rpLayout}
        onOpenSettings={() => openSettings('email')}
        onRefresh={refresh}
        onNew={openCompose}
        onNewAppt={openNewAppt}
        onNewContact={() => openNewItem('contact')}
        onNewTask={() => openNewItem('task')}
        onReply={() => openReply(false)}
        onReplyAll={() => openReply(true)}
        onForward={openForward}
        onSendReceive={sendReceive}
        onSyncAccount={syncOne}
        onShowProgress={() => setShowSyncDialog(true)}
        onLayoutChange={setLayout}
        onDeleteItem={deleteItem}
        onMoveItemTo={moveItemTo}
        onOpenSettingsTab={openSettings}
        selFolder={selFolderProp}
        folderTargets={folderTargets}
        onFolderNew={(name) =>
          selected && createFolder(selected.sourceId, selected.folder.id, name)
        }
        onFolderRename={(name) =>
          selected && renameFolder(selected.sourceId, selected.folder.id, name)
        }
        onFolderDelete={() => selected && deleteFolder(selected.sourceId, selected.folder.id)}
        onFolderCopy={() => selected && copyFolder(selected.sourceId, selected.folder.id)}
        onFolderMoveTo={(dest) =>
          selected && moveFolder(selected.sourceId, selected.folder.id, dest)
        }
        onMarkAllReadSel={() =>
          selected && markAllRead(selected.sourceId, selected.folder.id)
        }
        onRunRules={runRules}
        onExportBackup={exportBackup}
        onImportBackup={importBackup}
        onSaveAs={saveItemAs}
        onPrint={printItem}
        categories={categories}
        onCategorize={(catId) =>
          detail &&
          detailRef &&
          applyCategories(
            detailRef.sourceId,
            detailRef.folderId,
            selIndexes.size > 1 ? [...selIndexes] : [detail.index],
            catId,
          )
        }
        onCatsManage={() => setShowCatsDlg(true)}
      />
      {!hasAnySource ? null : (
        <>
          <div className="main-area">
            {module !== 'calendar' && (
            <>
            <div className="pane-slot-v" style={{ width: folderW }}>
              <FolderPane
                entries={entries}
                selected={selected ? { sourceId: selected.sourceId, folderId: selected.folder.id } : null}
                module={module}
                onSelect={loadFolder}
                onMoveItem={moveItem}
                onMoveFolder={moveFolder}
                onReorderFolder={reorderFolder}
                onReorderSource={reorderSource}
                onCreateFolder={createFolder}
                onRenameFolder={renameFolder}
                onDeleteFolder={deleteFolder}
                onCopyFolder={copyFolder}
                onEmptyFolder={emptyFolder}
                onMarkAllRead={markAllRead}
                onSortChildren={sortChildren}
                onMoveFolderStep={moveFolderStep}
                onToggleFavorite={toggleFavorite}
                onSetFolderHidden={setFolderHidden}
                favorites={favorites}
              />
            </div>
            <Splitter
              onDragStart={() => { dragBase.current = folderW; }}
              onDelta={(d) => paneResize('pane.folderW', setFolderW, dragBase.current, d, 160, 620)}
            />
            </>
            )}
            {module === 'mail' && rpLayout === 'right' && (
              <>
                <div className="pane-slot-v" style={{ width: msgW }}>
                  <MessageList
                    items={items}
                    total={itemCount}
                    selectedIndex={detail?.index ?? null}
                    searchResults={searchResults}
                    searchBusy={searchBusy}
                    onSearch={onSearch}
                    dragRef={selected ? { sourceId: selected.sourceId, folderId: selected.folder.id } : null}
                    onSelect={selectItem}
                    onOpenItem={openMessageWindow}
                    selIndexes={selIndexes}
                    onSelChange={onSelChange}
                    onDeleteKey={deleteItem}
                    moveEntries={entries}
                    onMoveMenuOpen={() => void refreshSources()}
                    onCtxAction={msgCtxAction}
                    onLoadMore={loadMore}
                    hasMore={items.length < itemCount}
                    loading={listLoading}
                    noUnreadFilter={isSentFolder}
                    listKey={selected ? `${selected.sourceId}|${selected.folder.id}` : ''}
                    displayedRef={displayedItemsRef}
                    categories={categories}
                  />
                </div>
                <Splitter
                  onDragStart={() => { dragBase.current = msgW; }}
                  onDelta={(d) => paneResize('pane.msgW', setMsgW, dragBase.current, d, 240, 800)}
                />
                <ReadingPane
                  item={detail}
                  loading={detailLoading}
                  onSaveAttachment={saveAttachment}
                  onOpenAttachment={openAttachment}
                  onSaveAllAttachments={saveAllAttachments}
                  loadRemoteImages={loadImages}
                  onRespondInvite={respondInvite}
                  onReply={() => openReply(false)}
                  onReplyAll={() => openReply(true)}
                  onForward={openForward}
                  onDelete={doDeleteItem}
                  onPrint={printItem}
                  onSaveAs={saveItemAs}
                  onEditItem={editDetailCb}
                  categories={categories}
                />
              </>
            )}
            {module === 'mail' && rpLayout === 'bottom' && (
              <div className="mail-vertical">
                <div className="pane-slot-h" style={{ height: msgH }}>
                  <MessageList
                    items={items}
                    total={itemCount}
                    selectedIndex={detail?.index ?? null}
                    searchResults={searchResults}
                    searchBusy={searchBusy}
                    onSearch={onSearch}
                    dragRef={selected ? { sourceId: selected.sourceId, folderId: selected.folder.id } : null}
                    onSelect={selectItem}
                    onOpenItem={openMessageWindow}
                    selIndexes={selIndexes}
                    onSelChange={onSelChange}
                    onDeleteKey={deleteItem}
                    moveEntries={entries}
                    onMoveMenuOpen={() => void refreshSources()}
                    onCtxAction={msgCtxAction}
                    onLoadMore={loadMore}
                    hasMore={items.length < itemCount}
                    loading={listLoading}
                    noUnreadFilter={isSentFolder}
                    listKey={selected ? `${selected.sourceId}|${selected.folder.id}` : ''}
                    displayedRef={displayedItemsRef}
                    categories={categories}
                  />
                </div>
                <Splitter
                  vertical={false}
                  onDragStart={() => { dragBase.current = msgH; }}
                  onDelta={(d) => paneResize('pane.msgH', setMsgH, dragBase.current, d, 120, 700)}
                />
                <ReadingPane
                  item={detail}
                  loading={detailLoading}
                  onSaveAttachment={saveAttachment}
                  onOpenAttachment={openAttachment}
                  onSaveAllAttachments={saveAllAttachments}
                  loadRemoteImages={loadImages}
                  onRespondInvite={respondInvite}
                  onReply={() => openReply(false)}
                  onReplyAll={() => openReply(true)}
                  onForward={openForward}
                  onDelete={doDeleteItem}
                  onPrint={printItem}
                  onSaveAs={saveItemAs}
                  onEditItem={editDetailCb}
                  categories={categories}
                />
              </div>
            )}
            {module === 'mail' && rpLayout === 'off' && (
              <div className="mail-full">
                <MessageList
                  items={items}
                  total={itemCount}
                  selectedIndex={detail?.index ?? null}
                  searchResults={searchResults}
                  searchBusy={searchBusy}
                  onSearch={onSearch}
                  dragRef={selected ? { sourceId: selected.sourceId, folderId: selected.folder.id } : null}
                  onSelect={selectItem}
                    onOpenItem={openMessageWindow}
                  selIndexes={selIndexes}
                  onSelChange={onSelChange}
                  onDeleteKey={deleteItem}
                  moveEntries={entries}
                    onMoveMenuOpen={() => void refreshSources()}
                  onCtxAction={msgCtxAction}
                  onLoadMore={loadMore}
                  hasMore={items.length < itemCount}
                  loading={listLoading}
                  noUnreadFilter={isSentFolder}
                  listKey={selected ? `${selected.sourceId}|${selected.folder.id}` : ''}
                  displayedRef={displayedItemsRef}
                  categories={categories}
                />
              </div>
            )}
            {module === 'calendar' && (
              <>
                <CalendarModule
                  sourceId={selected?.sourceId ?? null}
                  folder={selected?.folder ?? null}
                  items={items}
                  selectedIndex={detail?.index ?? null}
                  onSelect={selectItem}
                  calendars={allCalendars}
                  onEditItem={openItemEditor}
                />
              </>
            )}
            {module === 'people' && (
              <>
                <ContactsModule
                  folder={selected?.folder ?? null}
                  items={items}
                  selectedIndex={detail?.index ?? null}
                  onSelect={selectItem}
                />
                <ReadingPane
                  item={detail}
                  loading={detailLoading}
                  onSaveAttachment={saveAttachment}
                  onOpenAttachment={openAttachment}
                  onSaveAllAttachments={saveAllAttachments}
                  loadRemoteImages={loadImages}
                  onRespondInvite={respondInvite}
                  onEditItem={editDetailCb}
                />
              </>
            )}
            {module === 'tasks' && (
              <>
                <TasksModule
                  folder={selected?.folder ?? null}
                  sourceId={selected?.sourceId ?? null}
                  items={items}
                  selectedIndex={detail?.index ?? null}
                  onSelect={selectItem}
                  onEditItem={openItemEditor}
                  onChanged={() => selected && void loadFolder(selected.sourceId, selected.folder, true)}
                />
                <ReadingPane
                  item={detail}
                  loading={detailLoading}
                  onSaveAttachment={saveAttachment}
                  onOpenAttachment={openAttachment}
                  onSaveAllAttachments={saveAllAttachments}
                  loadRemoteImages={loadImages}
                  onRespondInvite={respondInvite}
                  onEditItem={editDetailCb}
                />
              </>
            )}
          </div>
          <ModuleBar
            module={module}
            onChange={switchModule}
            syncStatus={
              syncTasks.length
                ? {
                    active: syncTasks.some(
                      (t) => t.status === 'processing' || t.status === 'pending',
                    ),
                    doneTasks: syncTasks.filter(
                      (t) => t.status !== 'processing' && t.status !== 'pending',
                    ).length,
                    totalTasks: syncTasks.length,
                    onClick: () => setShowSyncDialog(true),
                  }
                : undefined
            }
          />
        </>
      )}
      {loading && (
        <div className="loading-overlay">
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
            <div className="spinner" />
            <div>{loading}</div>
          </div>
        </div>
      )}
      {error && hasAnySource && (
        <div
          style={{
            position: 'fixed',
            bottom: 56,
            right: 12,
            background: '#fde7e9',
            color: '#a4262c',
            border: '1px solid #f1aeb5',
            padding: '8px 14px',
            borderRadius: 6,
            zIndex: 60,
            cursor: 'pointer',
          }}
          onClick={() => setError('')}
        >
          {error}
        </div>
      )}
      {notice && !error && (
        <div
          style={{
            position: 'fixed',
            bottom: 56,
            right: 12,
            background: '#dff6dd',
            color: '#107c10',
            border: '1px solid #9fd89f',
            padding: '8px 14px',
            borderRadius: 6,
            zIndex: 60,
            cursor: 'pointer',
            boxShadow: '0 2px 10px rgba(0,0,0,0.12)',
          }}
          onClick={() => setNotice('')}
        >
          {notice}
        </div>
      )}
      {showSettings && (
        <AccountsSettings
          sources={sources}
          initialTab={settingsInitTab}
          editAccountId={settingsEditId}
          initialAction={settingsAction}
          onClose={() => {
            setShowSettings(false);
            setSettingsEditId(undefined);
            setSettingsAction(undefined);
            refreshImagePref();
            void refreshSources();
          }}
          onCloseSource={closeSource}
          onAccountsChanged={onAccountsChanged}
        />
      )}
      {confirm && (
        <ConfirmDialog
          title={confirm.title}
          message={confirm.message}
          confirmLabel={confirm.confirmLabel}
          onConfirm={confirm.onOk}
          onCancel={() => setConfirm(null)}
        />
      )}
      {showSyncDialog && (
        <SyncProgressDialog
          tasks={syncTasks}
          onClose={() => setShowSyncDialog(false)}
          onCancelAll={() => window.outlook.syncCancel()}
          onCancelTask={(taskId) => window.outlook.syncCancelTask?.(taskId)}
        />
      )}
      {openMsg && (
        <div className="modal-overlay" onMouseDown={(e) => { if (e.target === e.currentTarget) { setOpenMsg(null); setMsgMax(false); } }}>
          <div
            className={`appt-window msg-window${msgMax ? ' maximized' : ''}`}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="compose-title msg-title">
              <h2 title={openMsg.detail.subject}>
                <Icon name="envelope" size={16} /> {openMsg.detail.subject || t('(sem assunto)')}
              </h2>
              <div className="msg-actions">
                <button
                  className="btn btn-ghost"
                  onClick={() =>
                    openDraft(
                      replyDraftOf(openMsg.detail, false, {
                        sourceId: openMsg.sourceId,
                        folderId: openMsg.folderId,
                        index: openMsg.detail.index,
                      }),
                    )
                  }
                >
                  <Icon name="reply" size={14} /> {t('Responder')}
                </button>
                <button
                  className="btn btn-ghost"
                  onClick={() =>
                    openDraft(
                      replyDraftOf(openMsg.detail, true, {
                        sourceId: openMsg.sourceId,
                        folderId: openMsg.folderId,
                        index: openMsg.detail.index,
                      }),
                    )
                  }
                >
                  <Icon name="replyAll" size={14} /> {t('Responder a Todos')}
                </button>
                <button
                  className="btn btn-ghost"
                  onClick={() =>
                    openDraft(
                      forwardDraftOf(openMsg.detail, {
                        sourceId: openMsg.sourceId,
                        folderId: openMsg.folderId,
                        index: openMsg.detail.index,
                      }),
                    )
                  }
                >
                  <Icon name="forward" size={14} /> {t('Encaminhar')}
                </button>
                <button className="btn btn-ghost" onClick={printOpenMsg}>
                  <Icon name="printer" size={14} /> {t('Imprimir')}
                </button>
              </div>
              <div className="msg-winctl">
                <span
                  className="winctl"
                  title={t('Minimizar')}
                  onClick={() => { setOpenMsg(null); setMsgMax(false); }}
                >
                  –
                </span>
                <span
                  className="winctl"
                  title={msgMax ? t('Restaurar') : t('Maximizar')}
                  onClick={() => setMsgMax(!msgMax)}
                >
                  {msgMax ? '❐' : '▢'}
                </span>
                <span
                  className="winctl close"
                  title={t('Fechar')}
                  onClick={() => { setOpenMsg(null); setMsgMax(false); }}
                >
                  <Icon name="close" size={14} />
                </span>
              </div>
            </div>
            <ReadingPane
              item={openMsg.detail}
              loading={false}
              onSaveAttachment={saveOpenMsgAttachment}
              onOpenAttachment={openOpenMsgAttachment}
              onSaveAllAttachments={saveAllOpenMsgAttachments}
              loadRemoteImages={loadImages}
              onRespondInvite={respondInvite}
              categories={categories}
            />
          </div>
        </div>
      )}
      {newItemKind && (
        <NewItemDialog
          kind={newItemKind}
          sources={sources.filter((s) => s.kind === 'local' || s.kind === 'pst')}
          primarySourceId={apptSourceId}
          onClose={() => setNewItemKind(null)}
          onSave={saveSpecialItem}
        />
      )}
      {editingItem && (
        <NewItemDialog
          kind={editingItem.item.kind}
          sources={sources.filter((s) => s.kind === 'local' || s.kind === 'pst')}
          // Edição grava na própria base do item — não depende do arquivo principal.
          primarySourceId={editingItem.sourceId}
          initial={editingItem.item}
          onClose={() => setEditingItem(null)}
          onSave={(_sid, item) =>
            updateSpecialItem(
              {
                sourceId: editingItem.sourceId,
                folderId: editingItem.folderId,
                index: editingItem.index,
                stableId: editingItem.stableId,
              },
              item,
            )
          }
        />
      )}
      {reminders.length > 0 && (
        <ReminderDialog
          reminders={reminders}
          onDismiss={dismissReminderKey}
          onDismissAll={dismissAllReminders}
          onSnooze={snoozeReminders}
          onOpen={(r) => void openReminderItem(r)}
        />
      )}
      {showCatsDlg && (
        <CategoriesDialog
          cats={categories}
          onClose={() => setShowCatsDlg(false)}
          onSave={async (cats) => {
            await window.outlook.saveCategories(cats);
            setCategories(cats);
            setShowCatsDlg(false);
          }}
        />
      )}
      {profileMissing && (
        <div className="modal-overlay" style={{ zIndex: 250 }}>
          <div className="confirm-dialog" style={{ maxWidth: 460 }}>
            <div className="confirm-title">
              <Icon name="junk" size={18} />
              <span>{t('Pasta de dados não encontrada')}</span>
            </div>
            <div className="confirm-msg">
              {t('A pasta configurada não está acessível agora — se ela está num drive externo/removível, conecte-o ou aguarde a montagem e tente de novo. Nada foi alterado ou sincronizado.')}
            </div>
            <div
              className="confirm-msg"
              style={{
                fontFamily: 'monospace',
                fontSize: 11,
                wordBreak: 'break-all',
                background: 'var(--surface-mid)',
                border: '1px solid var(--border)',
                borderRadius: 6,
                padding: '6px 8px',
                marginTop: 4,
              }}
            >
              {profileMissing}
            </div>
            <div className="confirm-actions">
              <button
                className="ghost-btn"
                onClick={async () => {
                  const res = await window.outlook.pickDataDirPath?.();
                  if (!res) return;
                  await window.outlook.applyDataDir?.(res.dir);
                  await window.outlook.restartApp?.();
                }}
              >
                {t('Escolher outra pasta…')}
              </button>
              <button
                className="btn btn-primary"
                onClick={async () => {
                  const ps = await window.outlook.getProfileStatus?.();
                  if (ps?.available) await window.outlook.restartApp?.();
                }}
                autoFocus
              >
                {t('Tentar novamente')}
              </button>
            </div>
          </div>
        </div>
      )}
      {showOnboarding && (
        <OnboardingWizard
          onChanged={() => {
            void refreshSources();
            refreshAccounts();
          }}
          onDone={() => {
            setShowOnboarding(false);
            void refreshSources();
            refreshAccounts();
          }}
        />
      )}
    </div>
  );
}
