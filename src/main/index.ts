import { app, BrowserWindow, clipboard, dialog, ipcMain, Menu, nativeTheme, net, Notification, screen, shell } from 'electron';
import { basename, join, resolve } from 'path';
import { tmpdir } from 'os';
import { randomUUID, createHash } from 'crypto';
import { readFileSync, writeFileSync, existsSync, mkdirSync, copyFileSync, cpSync, readdirSync, statSync, rmSync } from 'fs';
import { writeFile } from 'fs/promises';
import { sourceManager } from './sourceManager';
import { PstService } from './pstService';
import { ImapMailSource } from './imapSource';
import { accountStore, secureStorageAvailable, migrateSecrets } from './accountStore';
import { sendMail, testAccount, resolveSignatureHtml, composeRawMessage } from './smtpService';
import { avStatus } from './antivirus';
import { blockedSenders } from './blockedSenders';
import { recordRecipients, searchContacts } from './contactHistory';
import { ruleEngine, ensureDeliveryStore, folderDirFor, kindDirFor, resolveStoreFolderName } from './ruleEngine';
import { signatureStore } from './signatureStore';
import { createLocalStore, deleteLocalStore, isLocalStorePath, listLocalStores, registerLocalStore, remapLocalStorePath, renameLocalStore, unregisterLocalStore } from './localStore';
import { dataDir, mailRoot, mailRootPath, setDataDir, isUnitMailProfile, scanMailDir, profilePointerExists, profilePointerTarget } from './dataDir';
import { appLanguage } from './language';
import { log } from './logger';
import { migratePstToLocalStore } from './pstImport';
import { isWindowsImportAvailable, importOutlookAccountsFromRegistry } from './outlookImportWin';
import { scanThunderbirdAccounts, scanThunderbirdMail } from './thunderbirdImport';
import { importThunderbirdMail } from './thunderbirdMailImport';
import type { AccountDraft, ComposeInput, FolderNode, ImportedAccountGuess, MailRule, Signature } from '../shared/types';
import { DEFAULT_CATEGORIES } from '../shared/types';
import { safeFsName } from '../shared/fsName';
import { folderKindOf, type FolderKind } from '../shared/folderNames';

const isDev = !!process.env.ELECTRON_RENDERER_URL;

interface Settings {
  recentFiles: string[];
  openPstFiles: string[];
  openLocalStores: string[];
  lastOpened?: string;
  /** Intervalo do Enviar/Receber automático em minutos (0 = desligado). Padrão 30, como o Outlook. */
  syncIntervalMin?: number;
  /** Baixar imagens externas no corpo HTML dos emails. Padrão false (como o Outlook, que bloqueia). */
  loadRemoteImages?: boolean;
  /** Verificar mensagens recebidas no antivírus local do sistema (AMSI, Defender, ClamAV…). Padrão true. */
  virusScan?: boolean;
  /** Jornada de primeiro acesso concluída. */
  setupComplete?: boolean;
  /** Ordem de exibição das pastas: chave "sourceId|parentId" → ids ordenados. */
  folderOrders?: Record<string, string[]>;
  /** Arquivo de dados principal — destino padrão de mensagens sem regra. */
  defaultStorePath?: string;
  /** Tema da interface: 'system' segue o SO. */
  theme?: 'system' | 'light' | 'dark';
  /** Idioma da interface — pt-BR ou en (pastas novas seguem o idioma). */
  language?: 'pt' | 'en';
  /** Re-sincronização total única já executada (correção do estado "visto"
   *  envenenado por versões que entregavam em diretório fantasma). */
  fullResyncDone?: boolean;
  /** Geometria da janela principal ao fechar (bounds normais + maximizada). */
  windowState?: { width: number; height: number; x?: number; y?: number; maximized?: boolean };
}

function settingsPath(): string {
  return join(dataDir(), 'settings.json');
}

function loadSettings(): Settings {
  try {
    if (existsSync(settingsPath())) {
      const parsed = JSON.parse(readFileSync(settingsPath(), 'utf-8'));
      return { recentFiles: [], openPstFiles: [], openLocalStores: [], ...parsed };
    }
  } catch {
    // arquivo corrompido — começa do zero
  }
  return { recentFiles: [], openPstFiles: [], openLocalStores: [] };
}

function saveSettings(s: Settings): void {
  try {
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(settingsPath(), JSON.stringify(s, null, 2));
  } catch {
    // falha ao salvar — não crítico
  }
}

// ---------- Lista-mestre de categorias (categories.json na raiz de dados) ----------

function categoriesPath(): string {
  return join(dataDir(), 'categories.json');
}

function loadCategories(): import('../shared/types').MailCategory[] {
  try {
    if (existsSync(categoriesPath())) {
      const parsed = JSON.parse(readFileSync(categoriesPath(), 'utf-8'));
      if (Array.isArray(parsed)) {
        return parsed
          .filter((c) => c && typeof c.id === 'string' && typeof c.name === 'string')
          .map((c) => ({ id: c.id, name: c.name, color: c.color || '#0078d4' }));
      }
    }
  } catch {
    // corrompido — volta aos padrões
  }
  return [...DEFAULT_CATEGORIES];
}

function saveCategories(cats: import('../shared/types').MailCategory[]): void {
  try {
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(categoriesPath(), JSON.stringify(cats, null, 2));
  } catch {
    // falha ao salvar — não crítico
  }
}

function addRecent(filePath: string): void {
  const s = loadSettings();
  s.recentFiles = [filePath, ...s.recentFiles.filter((f) => f !== filePath)].slice(0, 8);
  s.lastOpened = filePath;
  if (!s.openPstFiles.includes(filePath)) s.openPstFiles.push(filePath);
  saveSettings(s);
}

function removeOpenPst(filePath: string): void {
  const s = loadSettings();
  s.openPstFiles = s.openPstFiles.filter((f) => f !== filePath);
  saveSettings(s);
}

/** Mesmo id que SourceManager.stableId('pst', …) — sha1 do caminho resolvido. */
function pstSourceIdFor(filePath: string): string {
  return (
    'pst:' +
    createHash('sha1').update(resolve(filePath).toLowerCase()).digest('hex').slice(0, 16)
  );
}

/**
 * O PST foi reconhecido como o mesmo arquivo que vivia em `oldPath` — remapeia
 * TODAS as referências persistidas para o novo caminho/id: ordem de pastas,
 * listas de bases abertas, base principal, entrega das contas e alvo de regras.
 */
function remapPstReferences(oldPath: string, newPath: string): void {
  const s = loadSettings();
  const oldId = pstSourceIdFor(oldPath);
  const newId = pstSourceIdFor(newPath);
  const eq = (a?: string, b?: string) =>
    !!a && !!b && resolve(a).toLowerCase() === resolve(b).toLowerCase();
  const sw = (p: string) => (eq(p, oldPath) ? newPath : p);

  s.openPstFiles = s.openPstFiles.map(sw);
  s.openLocalStores = s.openLocalStores.map(sw);
  s.recentFiles = s.recentFiles.map(sw);
  if (eq(s.lastOpened, oldPath)) s.lastOpened = newPath;
  if (eq(s.defaultStorePath, oldPath)) s.defaultStorePath = newPath;

  const orders = s.folderOrders ?? {};
  // Remapeia chaves de ordem de pasta tanto de PST quanto de base local —
  // ambas derivam o sourceId do hash do caminho.
  const idPairs = [
    [oldId, newId],
    [
      'local:' + createHash('sha1').update(resolve(oldPath).toLowerCase()).digest('hex').slice(0, 16),
      'local:' + createHash('sha1').update(resolve(newPath).toLowerCase()).digest('hex').slice(0, 16),
    ],
  ];
  for (const [o, n] of idPairs) {
    for (const k of Object.keys(orders)) {
      if (k.startsWith(`${o}|`)) {
        orders[`${n}${k.slice(o.length)}`] = orders[k];
        delete orders[k];
      }
    }
  }
  saveSettings(s);

  // Registro de bases locais — sem isso o localstores.json continuava com o
  // caminho antigo (ex.: "D:\…") e a UI mostrava a base "quebrada".
  remapLocalStorePath(oldPath, newPath);

  // Entrega das contas apontando para o caminho antigo
  for (const acc of accountStore.list()) {
    if (eq(acc.deliveryPath, oldPath)) accountStore.patch(acc.id, { deliveryPath: newPath });
  }
  // Alvo das regras apontando para o caminho antigo
  for (const r of ruleEngine.list()) {
    if (eq(r.targetPstPath, oldPath)) ruleEngine.upsert({ ...r, targetPstPath: newPath });
  }
}

// Quando um PST movido é reconhecido por fingerprint, todas as referências
// ao caminho antigo são remapeadas para o novo — sem perder estado.
sourceManager.onPstMoved = remapPstReferences;

function addOpenLocalStore(path: string): void {
  const s = loadSettings();
  if (!s.openLocalStores.includes(path)) s.openLocalStores.push(path);
  saveSettings(s);
}

/** O primeiro arquivo de dados criado vira o principal (calendário e entrega padrão).
 *  Criações seguintes não mudam a principal — isso só se altera em Configurações. */
function adoptFirstAsDefault(path: string): void {
  const s = loadSettings();
  if (!s.defaultStorePath) {
    s.defaultStorePath = path;
    saveSettings(s);
  }
}

function removeOpenLocalStore(path: string): void {
  const s = loadSettings();
  s.openLocalStores = s.openLocalStores.filter((f) => f !== path);
  saveSettings(s);
}

// ---------- Toast de notificação (substitui a nativa do SO — estilizável) ----------
const toastWins: BrowserWindow[] = [];
const TOAST_W = 364;
const TOAST_H = 96;
const TOAST_GAP = 10;
const TOAST_MS = 8000;

function repositionToasts(): void {
  // Âncora: canto inferior direito da JANELA DO APP (como o aviso do Outlook
  // que gruda na janela). Se a janela estiver fechada/minimizada, cai pro
  // canto da tela onde está o cursor.
  let x: number;
  let baseY: number;
  const mw = mainWindow && !mainWindow.isDestroyed() && mainWindow.isVisible() && !mainWindow.isMinimized()
    ? mainWindow.getBounds()
    : null;
  if (mw) {
    x = mw.x + mw.width - TOAST_W - 14;
    baseY = mw.y + mw.height - TOAST_H - 14;
  } else {
    const area = screen.getDisplayNearestPoint(screen.getCursorScreenPoint()).workArea;
    x = area.x + area.width - TOAST_W - 14;
    baseY = area.y + area.height - TOAST_H - 14;
  }
  toastWins.forEach((w, i) => {
    if (w.isDestroyed()) return;
    w.setPosition(x, baseY - i * (TOAST_H + TOAST_GAP));
  });
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Toast no canto inferior direito, estilo Outlook — clicável (abre a janela). */
function showToast(kind: 'mail' | 'reminder' | 'warn' | 'error', title: string, body: string): void {
  try {
    const dark = nativeTheme.shouldUseDarkColors;
    const accent = { mail: '#0f6cbd', reminder: '#8764b8', warn: '#b34700', error: '#c50f1f' }[kind];
    const icon = {
      mail: '<path d="M4 5h16a1 1 0 011 1v12a1 1 0 01-1 1H4a1 1 0 01-1-1V6a1 1 0 011-1zm8 7.2L4.6 7v11h14.8V7L12 12.2z"/><path d="M4.4 6l7.6 5.7L19.6 6H4.4z" fill="none" stroke="currentColor" stroke-width="1.4"/>',
      reminder:
        '<circle cx="12" cy="13" r="7" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M12 9.5V13l2.4 1.6" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/><path d="M9 4h6" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>',
      warn: '<path d="M12 3.5l9.5 16.5H2.5L12 3.5z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"/><path d="M12 10v4.5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/><circle cx="12" cy="17.3" r="1.1"/>',
      error:
        '<circle cx="12" cy="12" r="8.5" fill="none" stroke="currentColor" stroke-width="1.7"/><path d="M9 9l6 6M15 9l-6 6" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>',
    }[kind];
    const html = `<!doctype html><html><head><meta charset="utf-8"><style>
      *{margin:0;box-sizing:border-box}html,body{height:100%;overflow:hidden;background:transparent}
      body{font-family:'Segoe UI',system-ui,sans-serif}
      .card{display:flex;align-items:stretch;height:100%;border-radius:10px;overflow:hidden;
        background:${dark ? '#2b2b2b' : '#ffffff'};color:${dark ? '#eee' : '#1b1b1b'};
        border:1px solid ${dark ? '#4a4a4a' : '#e0e0e0'};cursor:pointer;
        box-shadow:0 6px 24px rgba(0,0,0,${dark ? '.55' : '.22'});
        animation:slide .28s cubic-bezier(.2,.9,.3,1.2)}
      @keyframes slide{from{transform:translateX(60px);opacity:0}to{transform:none;opacity:1}}
      .bar{width:5px;flex:none;background:${accent}}
      .ico{width:40px;display:flex;align-items:center;justify-content:center;flex:none;color:${accent}}
      .ico svg{width:24px;height:24px}
      .tx{flex:1;min-width:0;padding:10px 4px 10px 0;display:flex;flex-direction:column;justify-content:center}
      .cap{font-size:10px;letter-spacing:.8px;text-transform:uppercase;color:${dark ? '#9a9a9a' : '#8a8a8a'}}
      .ti{font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:1px}
      .bo{font-size:12px;color:${dark ? '#c8c8c8' : '#5c5c5c'};margin-top:2px;
        display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
      .x{width:30px;flex:none;display:flex;align-items:flex-start;justify-content:center;
        color:${dark ? '#9a9a9a' : '#909090'};font-size:15px;padding-top:7px}
      .x:hover{color:${dark ? '#fff' : '#000'}}
    </style></head><body>
      <div class="card" onclick="location.href='um://open'"><div class="bar"></div>
        <div class="ico"><svg viewBox="0 0 24 24" fill="currentColor">${icon}</svg></div>
        <div class="tx"><div class="cap">Unit Mail</div><div class="ti">${esc(title)}</div>
        <div class="bo">${esc(body)}</div></div>
        <div class="x" onclick="event.stopPropagation();location.href='um://close'">&#10005;</div>
      </div></body></html>`;
    const win = new BrowserWindow({
      width: TOAST_W,
      height: TOAST_H,
      frame: false,
      transparent: true,
      resizable: false,
      movable: false,
      minimizable: false,
      maximizable: false,
      fullscreenable: false,
      alwaysOnTop: true,
      skipTaskbar: true,
      focusable: false,
      show: false,
      webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true },
    });
    win.setAlwaysOnTop(true, 'screen-saver');
    win.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
    win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    win.webContents.on('will-navigate', (e, url) => {
      e.preventDefault();
      if (url.startsWith('um://open')) {
        if (mainWindow) {
          if (mainWindow.isMinimized()) mainWindow.restore();
          mainWindow.show();
          mainWindow.focus();
        }
      }
      if (!win.isDestroyed()) win.close();
    });
    toastWins.push(win);
    win.on('closed', () => {
      const i = toastWins.indexOf(win);
      if (i >= 0) toastWins.splice(i, 1);
      repositionToasts();
    });
    void win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    win.once('ready-to-show', () => {
      repositionToasts();
      win.showInactive();
    });
    setTimeout(() => {
      if (!win.isDestroyed()) win.close();
    }, TOAST_MS);
  } catch {
    if (Notification.isSupported()) new Notification({ title, body }).show();
  }
}

// ---------- Enviar/Receber automático ----------
let syncTimer: ReturnType<typeof setInterval> | null = null;

/** Roda o sync e avisa o renderer + notifica o SO quando chegaram mensagens. */
async function runSyncAndNotify(): Promise<void> {
  try {
    ruleEngine.onProgress = (ev) => {
      mainWindow?.webContents.send('sync:progress', ev);
    };
    const results = await ruleEngine.syncAll();
    mainWindow?.webContents.send('sync:done', results);
    const archived = results.reduce((a, r) => a + r.archived, 0);
    const infected = results.reduce((a, r) => a + (r.infected ?? 0), 0);
    if (infected > 0) {
      showToast(
        'warn',
        'Ameaça detectada em mensagem recebida',
        `${infected} mensagem(ns) movida(s) para a pasta "Itens infectados".`,
      );
    }
    if (archived > 0 && mainWindow && !mainWindow.isFocused()) {
      showToast('mail', 'Novas mensagens', `${archived} mensagem(ns) recebida(s) e arquivada(s).`);
    }
    // IMAP: mensagem nova no servidor (nada foi arquivado localmente).
    const newOnServer = results.reduce((a, r) => a + (r.newOnServer ?? 0), 0);
    if (newOnServer > 0 && mainWindow && !mainWindow.isFocused()) {
      showToast('mail', 'Novas mensagens', `${newOnServer} mensagem(ns) nova(s) no servidor.`);
    }
  } catch {
    // erros de rede não devem derrubar o timer
  }
}

function scheduleAutoSync(): void {
  if (syncTimer) clearInterval(syncTimer);
  syncTimer = null;
  const min = loadSettings().syncIntervalMin ?? 30;
  if (min > 0) {
    syncTimer = setInterval(() => {
      runSyncAndNotify();
    }, min * 60000);
  }
}

// ---------- Lembretes de compromissos ----------
const remindedKeys = new Set<string>();
/** "Adiar" do modal de lembrete — chave → timestamp em que dispara de novo. */
const snoozedUntil = new Map<string, number>();

/** Lembretes "dispensados" — persistidos em disco para não voltarem a
 *  disparar quando o app reabre dentro da janela do compromisso. */
const dismissedReminders = new Set<string>();
let dismissedLoaded = false;

function dismissedRemindersPath(): string {
  return join(dataDir(), 'dismissed-reminders.json');
}

function loadDismissedReminders(): void {
  if (dismissedLoaded) return;
  dismissedLoaded = true;
  try {
    const raw = JSON.parse(readFileSync(dismissedRemindersPath(), 'utf-8')) as unknown;
    if (Array.isArray(raw)) for (const k of raw) dismissedReminders.add(String(k));
  } catch {
    // sem arquivo ainda
  }
}

function saveDismissedReminders(): void {
  try {
    // Poda: a key termina com o startTime ISO — descarta compromissos
    // passados há mais de 2 dias para o arquivo não crescer sem fim.
    const cutoff = Date.now() - 2 * 24 * 60 * 60 * 1000;
    const kept = [...dismissedReminders].filter((k) => {
      const ts = Date.parse(k.slice(k.lastIndexOf('|') + 1));
      return isNaN(ts) || ts >= cutoff;
    });
    dismissedReminders.clear();
    for (const k of kept) dismissedReminders.add(k);
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(dismissedRemindersPath(), JSON.stringify(kept));
  } catch {
    // best-effort
  }
}

/** Respostas já enviadas a convites — UID do VEVENT → 'accept'|'decline'|
 *  'counter'. Persistido: a barra de convite mostra o estado ao reabrir a
 *  mensagem em vez dos botões ativos. */
const respondedInvites = new Map<string, string>();
let respondedLoaded = false;

function respondedInvitesPath(): string {
  return join(dataDir(), 'responded-invites.json');
}

function loadRespondedInvites(): void {
  if (respondedLoaded) return;
  respondedLoaded = true;
  try {
    const raw = JSON.parse(readFileSync(respondedInvitesPath(), 'utf-8')) as unknown;
    if (raw && typeof raw === 'object') {
      for (const [k, v] of Object.entries(raw as Record<string, string>)) {
        respondedInvites.set(k, String(v));
      }
    }
  } catch {
    // sem arquivo ainda
  }
}

function markInviteResponded(uid: string, response: string): void {
  if (!uid) return;
  loadRespondedInvites();
  respondedInvites.set(uid, response);
  try {
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(respondedInvitesPath(), JSON.stringify(Object.fromEntries(respondedInvites), null, 1));
  } catch {
    // best-effort
  }
}

function findFolderNode(nodes: FolderNode[], id: string): FolderNode | null {
  for (const n of nodes) {
    if (n.id === id) return n;
    const f = findFolderNode(n.children, id);
    if (f) return f;
  }
  return null;
}

/** Tipo semântico da pasta de um item — arc:* pelo nome; pastas do PST pelo
 *  containerClass/nome do nó na árvore. */
function mailFolderKind(sourceId: string, folderId: string): FolderKind | null {
  if (folderId.startsWith('arc:')) {
    return folderKindOf(folderId.slice(4).split('/').pop() || '');
  }
  try {
    const svc = sourceManager.backend(sourceId) as import('./pstService').PstService;
    if (typeof svc.getFolderTree !== 'function') return null;
    const n = findFolderNode(svc.getFolderTree(), folderId);
    if (!n) return null;
    if (n.containerClass === 'IPF.Appointment') return 'calendar';
    if (n.containerClass === 'IPF.Contact') return 'contacts';
    if (n.containerClass === 'IPF.Task') return 'tasks';
    return folderKindOf(n.archiveName || '') || folderKindOf(n.name);
  } catch {
    return null;
  }
}

/** UIDs dos VEVENTs presentes nas pastas de Calendário — um convite cujo
 *  compromisso já existe no calendário foi aceito (cobre respostas feitas
 *  antes da persistência de responded-invites). Cache curto: a leitura das
 *  pastas usa o índice de resumos, mas varremos todas as bases. */
let calUidCache: { at: number; uids: Set<string> } | null = null;

function invalidateCalendarUids(): void {
  calUidCache = null;
}

async function calendarUids(): Promise<Set<string>> {
  if (calUidCache && Date.now() - calUidCache.at < 60_000) return calUidCache.uids;
  const uids = new Set<string>();
  for (const src of sourceManager.listAll()) {
    const svc = sourceManager.backend(src.id);
    let tree: FolderNode[];
    try {
      // IMAP returns the tree asynchronously; bases PST/local return it
      // synchronously. Await funciona nos dois casos e evita passar uma
      // Promise para calendarFolders quando uma conta IMAP está aberta.
      tree = await svc.getFolderTree();
    } catch {
      continue;
    }
    for (const cal of calendarFolders(tree)) {
      try {
        const items = await svc.getItems(cal.id, 0, 5000);
        for (const it of items) if (it.apptUid) uids.add(it.apptUid);
      } catch {
        // pasta inacessível — segue nas demais
      }
    }
  }
  calUidCache = { at: Date.now(), uids };
  return uids;
}

/** Procura pastas de calendário na árvore de uma base. */
function calendarFolders(nodes: FolderNode[]): FolderNode[] {
  const out: FolderNode[] = [];
  const walk = (ns: FolderNode[]) => {
    for (const n of ns) {
      if (
        n.containerClass === 'IPF.Appointment' ||
        folderKindOf(n.name) === 'calendar' ||
        folderKindOf(n.archiveName || '') === 'calendar'
      ) {
        out.push(n);
      }
      walk(n.children);
    }
  };
  walk(nodes);
  return out;
}

async function checkReminders(): Promise<void> {
  loadDismissedReminders();
  const now = Date.now();
  for (const src of sourceManager.listAll()) {
    const svc = sourceManager.backend(src.id) as import('./pstService').PstService;
    if (typeof svc.getFolderTree !== 'function' || typeof svc.getItems !== 'function') continue;
    let tree: FolderNode[];
    try {
      tree = svc.getFolderTree();
    } catch {
      continue;
    }
    for (const cal of calendarFolders(tree)) {
      let items;
      try {
        items = await svc.getItems(cal.id, 0, 400);
      } catch {
        continue;
      }
      for (const it of items) {
        if (!it.startTime) continue;
        const start = new Date(it.startTime).getTime();
        if (isNaN(start)) continue;
        // Sem lembrete marcado no PST → não avisa (comportamento do Outlook)
        if (it.reminderSet === false) continue;
        const minBefore = it.reminderMinutes ?? 15;
        const remindAt = start - minBefore * 60000;
        // stableId na chave — o índice posicional muda quando a pasta
        // recebe itens, o que re-disparava ou perdia lembretes.
        const key = `${src.id}|${cal.id}|${it.stableId ?? it.index}|${it.startTime}`;
        if (remindedKeys.has(key) || dismissedReminders.has(key)) continue;
        // "Adiar": dentro da janela de soneca, pula; ao expirar, toca de novo
        const snz = snoozedUntil.get(key);
        if (snz) {
          if (now < snz) continue;
          snoozedUntil.delete(key);
        }
        // Janela: de (início - X min) até 1h depois do início
        if (now >= remindAt && now <= start + 60 * 60000) {
          remindedKeys.add(key);
          // Idioma da UI também na notificação nativa do SO
          const lang = appLanguage();
          const body = `${new Date(it.startTime).toLocaleString(
            lang === 'en' ? 'en-US' : 'pt-BR',
          )}${it.location ? ` — ${it.location}` : ''}`;
          try {
            showToast(
              'reminder',
              lang === 'en' ? `Reminder: ${it.subject}` : `Lembrete: ${it.subject}`,
              body,
            );
          } catch {
            // notificação pode falhar em ambientes sem suporte
          }
          mainWindow?.webContents.send('calendar:reminder', {
            key,
            subject: it.subject,
            startTime: it.startTime,
            location: it.location || '',
            sourceId: src.id,
            folderId: cal.id,
            index: it.index,
            stableId: it.stableId,
          });
        }
      }
    }
  }
}

function startReminderTimer(): void {
  // Primeira verificação após 10s (bases restauradas), depois a cada minuto
  setTimeout(() => {
    checkReminders().catch(() => {});
    setInterval(() => checkReminders().catch(() => {}), 60 * 1000);
  }, 10 * 1000);
}

/** Extrai endereços de uma lista de participantes ("a@b; Nome <c@d>, e@f"). */
function parseAttendees(raw?: string): string[] {
  const out: string[] = [];
  for (const tok of (raw || '').split(/[;,\n]+/)) {
    const t = tok.trim();
    if (!t) continue;
    const m = t.match(/<([^>]+)>/);
    const email = (m ? m[1] : t).trim();
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) && !out.includes(email)) out.push(email);
  }
  return out;
}

/**
 * Envia convites de reunião iMIP (METHOD:REQUEST) aos participantes do
 * compromisso recém-criado. O .ics vai como alternativa text/calendar, então
 * Outlook/Thunderbird/Gmail exibem Aceitar/Recusar nativos. O UID é o mesmo
 * gravado no compromisso — respostas e atualizações correlacionam.
 * Retorna null quando não há participantes; senão o resumo do envio.
 */
/** Conta associada a uma base de dados (deliveryPath); senão a primeira
 *  com email — mesma regra do "Responder" no renderer. */
function accountForStore(sourceId: string) {
  const norm = (p: string) => p.replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();
  const src = sourceManager.listAll().find((s) => s.id === sourceId);
  const accts = accountStore.list().filter((a) => a.email);
  return (
    (src?.filePath &&
      accts.find((a) => a.deliveryPath && norm(a.deliveryPath) === norm(src.filePath!))) ||
    accts[0] ||
    undefined
  );
}

async function sendApptInvites(
  sourceId: string,
  item: import('../shared/types').NewSpecialItem,
): Promise<import('../shared/types').CreateSpecialResult['invite']> {
  const attendees = parseAttendees(item.attendees);
  if (!attendees.length) return null;

  // Conta remetente: a vinculada à base de dados do compromisso; senão a
  // primeira configurada — mesma regra do "Responder" no renderer.
  const account = accountForStore(sourceId);
  if (!account) return { sent: 0, failed: attendees.length, noAccount: true };

  const en = appLanguage() === 'en';
  const loc = en ? 'en-US' : 'pt-BR';
  const L = en
    ? { invited: 'You are invited to a meeting.', when: 'When', where: 'Location', by: 'Organizer' }
    : { invited: 'Você foi convidado para um compromisso.', when: 'Quando', where: 'Local', by: 'Organizador' };
  const esc = (s: string) =>
    s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
  const fold = (s: string) => s.replace(/([\\;,])/g, '\\$1').replace(/\r?\n/g, '\\n');
  const icsUtc = (iso?: string | null): string =>
    iso && !isNaN(new Date(iso).getTime())
      ? new Date(iso).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '')
      : '';
  const icsDate = (iso?: string | null): string => {
    const d = iso ? new Date(iso) : null;
    if (!d || isNaN(d.getTime())) return '';
    const p = (n: number) => String(n).padStart(2, '0');
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`;
  };
  const when = item.start
    ? `${new Date(item.start).toLocaleString(loc)}${item.end ? ` — ${new Date(item.end).toLocaleString(loc)}` : ''}`
    : '';

  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Unit Mail//Compromisso//PT-BR',
    'CALSCALE:GREGORIAN',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    `UID:${item.uid || `${randomUUID()}@unitmail`}`,
    `DTSTAMP:${icsUtc(new Date().toISOString())}`,
    item.allDay && item.start
      ? `DTSTART;VALUE=DATE:${icsDate(item.start)}`
      : item.start
        ? `DTSTART:${icsUtc(item.start)}`
        : '',
    item.allDay && item.end
      ? `DTEND;VALUE=DATE:${icsDate(item.end)}`
      : item.end
        ? `DTEND:${icsUtc(item.end)}`
        : '',
    `SUMMARY:${fold(item.subject || '')}`,
    item.location ? `LOCATION:${fold(item.location)}` : '',
    item.body ? `DESCRIPTION:${fold(item.body)}` : '',
    `ORGANIZER;CN=${account.label || account.email}:mailto:${account.email}`,
    ...attendees.map(
      (a) =>
        `ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=${a}:mailto:${a}`,
    ),
    `SEQUENCE:${item.seq ?? 0}`,
    'STATUS:CONFIRMED',
    'TRANSP:OPAQUE',
    item.reminderMin != null
      ? `BEGIN:VALARM\nTRIGGER:-PT${item.reminderMin}M\nACTION:DISPLAY\nEND:VALARM`
      : '',
    'END:VEVENT',
    'END:VCALENDAR',
  ]
    .filter(Boolean)
    .join('\r\n');

  const html = `
    <p>${L.invited}</p>
    <table style="font-family:sans-serif;font-size:14px">
      <tr><td style="padding-right:12px;color:#666"><b>${L.when}</b></td><td>${esc(when)}</td></tr>
      ${item.location ? `<tr><td style="padding-right:12px;color:#666"><b>${L.where}</b></td><td>${esc(item.location)}</td></tr>` : ''}
      <tr><td style="padding-right:12px;color:#666"><b>${L.by}</b></td><td>${esc(account.label || account.email)} &lt;${esc(account.email)}&gt;</td></tr>
    </table>
    ${item.body ? `<p>${esc(item.body).replace(/\n/g, '<br>')}</p>` : ''}`;

  const input: import('../shared/types').ComposeInput = {
    accountId: account.id,
    to: attendees.join(', '),
    subject: item.subject || (en ? 'Meeting invitation' : 'Convite de compromisso'),
    html,
    noSignature: true,
    icalEvent: { method: 'REQUEST', contentBase64: Buffer.from(ics, 'utf8').toString('base64') },
    rawAttachments: [
      {
        filename: 'invite.ics',
        contentBase64: Buffer.from(ics, 'utf8').toString('base64'),
        contentType: 'text/calendar; method=REQUEST',
      },
    ],
  };

  try {
    await sendMail(
      account,
      accountStore.getPassword(account.id),
      input,
      accountStore.getSmtpPassword(account.id),
    );
    // Cópia em Itens Enviados — mesmo fluxo do envio normal de mensagens
    try {
      const raw = await composeRawMessage(account, input);
      let copied = false;
      if (account.protocol === 'imap') {
        try {
          const be = sourceManager.backend(`acct:${account.id}`);
          if (be instanceof ImapMailSource) {
            const sentPath = await be.specialFolderPath('sent');
            if (sentPath) {
              await be.appendItem(sentPath, raw, true);
              copied = true;
            }
          }
        } catch {
          // fallback local abaixo
        }
      }
      const store = copied
        ? ''
        : account.protocol === 'imap'
          ? loadSettings().defaultStorePath || ''
          : ensureDeliveryStore(account);
      if (store) {
        let svc = sourceManager.serviceForPath(store);
        if (!svc) {
          const info = store.toLowerCase().endsWith('.pst')
            ? sourceManager.openPst(store)
            : sourceManager.openLocalStore(store);
          const b = sourceManager.backend(info.id);
          if (b instanceof PstService) svc = b;
        }
        if (svc) {
          svc.importNamedEml(resolveStoreFolderName(store, 'sent'), raw, true);
        } else {
          const dir = kindDirFor(store, 'sent');
          mkdirSync(dir, { recursive: true });
          writeFileSync(join(dir, `${Date.now()}-${randomUUID().slice(0, 8)}.eml`), raw);
        }
      }
    } catch {
      // cópia local best-effort — o convite já foi enviado
    }
    return { sent: attendees.length, failed: 0 };
  } catch (e) {
    return { sent: 0, failed: attendees.length, error: (e as Error).message };
  }
}

let mainWindow: BrowserWindow | null = null;
let splashWin: BrowserWindow | null = null;
/** Resolvida quando o conteúdo da janela principal está pronto para exibir. */
let mainWindowReady: Promise<void> = Promise.resolve();

/** Ícone da janela/app (Linux exibe na barra de tarefas/alt-tab). */
function appIcon(): string | undefined {
  const p = existsSync(join(process.resourcesPath || '', 'icon.png'))
    ? join(process.resourcesPath, 'icon.png')
    : join(app.getAppPath(), 'build', 'icon.png');
  return existsSync(p) ? p : undefined;
}

/**
 * Splash screen de abertura — janela pequena sem moldura, carregada de um
 * data: URL (sem depender de arquivo em disco). Aparece em <1s e informa o
 * progresso da restauração das bases, que numa instalação com vários GB
 * de email pode levar alguns segundos.
 */
function splashHtml(): string {
  const v = app.getVersion();
  return `<!doctype html><html><head><meta charset="utf-8"><style>
    * { margin: 0; box-sizing: border-box; }
    html, body { height: 100%; }
    body {
      font-family: 'Segoe UI', 'Noto Sans', system-ui, sans-serif;
      background: linear-gradient(160deg, #0a5cb8 0%, #0173c7 55%, #01579b 100%);
      color: #fff; display: flex; flex-direction: column; align-items: center;
      justify-content: center; gap: 14px; user-select: none; cursor: default;
      -webkit-app-region: drag;
    }
    .logo { filter: drop-shadow(0 6px 18px rgba(0,0,0,.35)); }
    h1 { font-size: 26px; font-weight: 600; letter-spacing: .3px; }
    .sub { font-size: 12px; opacity: .82; }
    .spin {
      width: 26px; height: 26px; border-radius: 50%;
      border: 3px solid rgba(255,255,255,.25); border-top-color: #fff;
      animation: rot .8s linear infinite; margin-top: 8px;
    }
    @keyframes rot { to { transform: rotate(360deg) } }
    #st { font-size: 11px; opacity: .75; min-height: 14px; max-width: 380px;
          overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .ver { position: absolute; bottom: 12px; font-size: 10px; opacity: .55; }
  </style></head><body>
    <svg class="logo" width="72" height="72" viewBox="0 0 64 64" fill="none">
      <rect x="4" y="12" width="56" height="40" rx="5" fill="#fff"/>
      <path d="M6 16 L32 36 L58 16" stroke="#0173c7" stroke-width="4"
            fill="none" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
    <h1>Unit Mail</h1>
    <div class="sub">Seu email, do seu jeito</div>
    <div class="spin"></div>
    <div id="st">Iniciando…</div>
    <div class="ver">v${v}</div>
  </body></html>`;
}

function createSplash(): void {
  splashWin = new BrowserWindow({
    width: 440,
    height: 340,
    frame: false,
    resizable: false,
    center: true,
    show: false,
    skipTaskbar: true,
    backgroundColor: '#0173c7',
    icon: appIcon(),
  });
  void splashWin.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(splashHtml())}`);
  splashWin.once('ready-to-show', () => splashWin?.show());
}

/** Atualiza a linha de status da splash ("Abrindo base X…"). */
function splashStatus(msg: string): void {
  const w = splashWin;
  if (!w || w.isDestroyed()) return;
  w.webContents
    .executeJavaScript(
      `var e=document.getElementById('st');if(e)e.textContent=${JSON.stringify(msg)};`,
    )
    .catch(() => {});
}

/** Estado salvo da janela — posição descartada se cair fora de todo monitor. */
function savedWindowState(): {
  width: number;
  height: number;
  x?: number;
  y?: number;
  maximized?: boolean;
} | null {
  const ws = loadSettings().windowState;
  if (!ws || typeof ws.width !== 'number' || typeof ws.height !== 'number') return null;
  const st = { ...ws, maximized: !!ws.maximized };
  if (st.x !== undefined && st.y !== undefined) {
    const onScreen = screen.getAllDisplays().some((d) => {
      const a = d.workArea;
      return (
        st.x! >= a.x - 40 &&
        st.y! >= a.y - 40 &&
        st.x! < a.x + a.width &&
        st.y! < a.y + a.height
      );
    });
    if (!onScreen) {
      st.x = undefined;
      st.y = undefined;
    }
  }
  return st;
}

/** Chrome da barra de título integrada (faixa azul) — principal e composição. */
function windowChrome():
  | {
      titleBarStyle: 'hidden';
      titleBarOverlay: { color: string; symbolColor: string; height: number };
    }
  | Record<string, never> {
  // Barra de título integrada: os botões min/max/close do sistema ficam
  // sobre uma faixa azul renderizada pelo app (estilo Office moderno).
  // titleBarOverlay é suportado no Windows e no Linux; sem ele o Linux
  // mostrava a moldura nativa (barra preta do GNOME/KDE) em vez do azul.
  return process.platform === 'win32' || process.platform === 'linux'
    ? {
        titleBarStyle: 'hidden' as const,
        titleBarOverlay: { color: '#0173c7', symbolColor: '#ffffff', height: 36 },
      }
    : {};
}

/** Cor de fundo da janela ANTES do renderer pintar — segue o tema salvo
 *  para não piscar claro num perfil escuro (e vice-versa). */
function windowBgForTheme(): string {
  const theme = loadSettings().theme ?? 'system';
  const dark = theme === 'dark' || (theme === 'system' && nativeTheme.shouldUseDarkColors);
  return dark ? '#171717' : '#f3f3f3';
}

function createWindow(): BrowserWindow {
  const iconPath = appIcon();
  const winChrome = windowChrome();

  const ws = savedWindowState();
  const win = new BrowserWindow({
    width: Math.max(1024, ws?.width ?? 1440),
    height: Math.max(600, ws?.height ?? 900),
    x: ws?.x,
    y: ws?.y,
    minWidth: 1024,
    minHeight: 600,
    show: false,
    title: 'Unit Mail',
    icon: iconPath,
    backgroundColor: windowBgForTheme(),
    ...winChrome,
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  win.setMenuBarVisibility(false);

  // Links do corpo dos emails (iframe sandboxed geram popups que caem aqui):
  // http(s) abre no navegador do SO; mailto: abre rascunho no próprio app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('http://') || url.startsWith('https://')) {
      void shell.openExternal(url);
    } else if (url.startsWith('mailto:')) {
      win.webContents.send('mail:openMailto', url);
    }
    return { action: 'deny' };
  });

  // Menu de contexto (botão direito) — cobre o corpo do email dentro do iframe
  // sandboxed (o evento do renderer não chega lá) e os campos editáveis.
  // Links externos só oferecem "abrir" para http(s) — esquemas tipo chatgpt://
  // não devem disparar apps de terceiros.
  win.webContents.on('context-menu', (_e, p) => {
    const en = appLanguage() === 'en';
    const menu: Electron.MenuItemConstructorOptions[] = [];
    if (p.linkURL) {
      if (/^https?:/i.test(p.linkURL)) {
        menu.push({
          label: en ? 'Open link in browser' : 'Abrir link no navegador',
          click: () => void shell.openExternal(p.linkURL),
        });
      }
      menu.push(
        {
          label: en ? 'Copy link address' : 'Copiar endereço do link',
          click: () => clipboard.writeText(p.linkURL),
        },
        { type: 'separator' },
      );
    }
    if (p.isEditable) {
      menu.push(
        { label: en ? 'Undo' : 'Desfazer', role: 'undo', enabled: p.editFlags.canUndo },
        { label: en ? 'Redo' : 'Refazer', role: 'redo', enabled: p.editFlags.canRedo },
        { type: 'separator' },
        { label: en ? 'Cut' : 'Recortar', role: 'cut', enabled: p.editFlags.canCut },
        { label: en ? 'Copy' : 'Copiar', role: 'copy', enabled: p.editFlags.canCopy },
        { label: en ? 'Paste' : 'Colar', role: 'paste', enabled: p.editFlags.canPaste },
        { type: 'separator' },
        { label: en ? 'Select all' : 'Selecionar tudo', role: 'selectAll' },
      );
    } else {
      if (p.selectionText.trim()) {
        menu.push({ label: en ? 'Copy' : 'Copiar', role: 'copy' });
      }
      menu.push({ label: en ? 'Select all' : 'Selecionar tudo', role: 'selectAll' });
    }
    if (menu.length) Menu.buildFromTemplate(menu).popup({ window: win });
  });

  // Estava maximizada ao fechar → reabre maximizada. Chamar maximize() antes
  // do show é seguro: a janela aparece já maximizada, sem piscar.
  if (ws?.maximized) win.maximize();

  // Persistência da geometria: getNormalBounds() devolve o tamanho de quando
  // a janela NÃO está maximizada — é o valor certo para restaurar depois.
  const saveWinState = () => {
    if (win.isDestroyed()) return;
    const s = loadSettings();
    s.windowState = { ...win.getNormalBounds(), maximized: win.isMaximized() };
    saveSettings(s);
  };
  let wsTimer: ReturnType<typeof setTimeout> | null = null;
  const queueWinState = () => {
    if (wsTimer) clearTimeout(wsTimer);
    wsTimer = setTimeout(saveWinState, 600);
  };
  win.on('resize', queueWinState);
  win.on('move', queueWinState);
  win.on('maximize', queueWinState);
  win.on('unmaximize', queueWinState);
  win.on('close', saveWinState);
  // Toasts seguem a janela quando ela anda/muda de tamanho
  win.on('move', repositionToasts);
  win.on('resize', repositionToasts);
  win.on('maximize', repositionToasts);
  win.on('unmaximize', repositionToasts);

  mainWindow = win;
  win.on('closed', () => {
    if (wsTimer) clearTimeout(wsTimer);
    if (mainWindow === win) mainWindow = null;
  });

  mainWindowReady = new Promise((resolve) => {
    // ready-to-show dispara quando a página PODE pintar — mas o React ainda
    // não montou e a janela aparecia em branco por um instante. Esperamos o
    // renderer sinalizar o primeiro paint real (app:firstPaint, enviado
    // depois do mount + double-RAF). Fallback: ready-to-show + margem, caso
    // um renderer antigo não envie o sinal.
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      resolve();
    };
    ipcMain.once('app:firstPaint', finish);
    win.once('ready-to-show', () => setTimeout(finish, 1500));
  });

  if (isDev) {
    win.loadURL(process.env.ELECTRON_RENDERER_URL!);
    win.webContents.openDevTools({ mode: 'detach' });
  } else {
    win.loadFile(join(__dirname, '../renderer/index.html'));
  }
  return win;
}

// ---------- Janela de composição ----------
// Rascunhos (novo email, resposta, encaminhamento) abrem numa janela REAL —
// igual ao Outlook: o usuário pode minimizá-la e voltar ao app para copiar
// dados, e reabrir a janela pela barra de tarefas. O rascunho viaja por um
// id; a janela o busca via 'compose:getDraft' ao carregar.
const composeDrafts = new Map<string, Record<string, unknown>>();

function openComposeWindow(draft: Record<string, unknown>): void {
  const id = randomUUID();
  composeDrafts.set(id, draft);
  const w = new BrowserWindow({
    width: 980,
    height: 720,
    minWidth: 640,
    minHeight: 460,
    show: false,
    title: 'Unit Mail — Mensagem',
    icon: appIcon(),
    backgroundColor: windowBgForTheme(),
    ...windowChrome(),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  w.setMenuBarVisibility(false);
  w.once('ready-to-show', () => w.show());
  w.on('closed', () => composeDrafts.delete(id));
  if (isDev) {
    w.loadURL(`${process.env.ELECTRON_RENDERER_URL}#/compose/${id}`);
  } else {
    w.loadFile(join(__dirname, '../renderer/index.html'), { hash: `/compose/${id}` });
  }
}

function registerIpc(): void {
  // ---------- Janela de composição ----------
  ipcMain.handle('compose:openWindow', (_e, draft: Record<string, unknown>) =>
    openComposeWindow(draft),
  );
  ipcMain.handle('compose:getDraft', (_e, id: string) => composeDrafts.get(id) ?? null);
  ipcMain.handle('compose:notifySent', () => {
    // A janela principal atualiza a lista (Itens Enviados) quando um
    // rascunho aberto em janela separada é enviado.
    mainWindow?.webContents.send('compose:sent');
  });

  // ---------- Diálogos de arquivo ----------
  // ---------- Ajuda / Sobre ----------
  /** "1.2.10" > "1.2.3"? — comparação semver simples (só números). */
  const newerThan = (a: string, b: string): boolean => {
    const pa = a.split(/[.-]/).map((n) => parseInt(n, 10) || 0);
    const pb = b.split(/[.-]/).map((n) => parseInt(n, 10) || 0);
    for (let i = 0; i < 4; i++) {
      if ((pa[i] || 0) > (pb[i] || 0)) return true;
      if ((pa[i] || 0) < (pb[i] || 0)) return false;
    }
    return false;
  };

  /** Verifica a última release no GitHub — usado na abertura e no menu Ajuda. */
  ipcMain.handle('app:checkUpdate', async () => {
    try {
      const res = await net.fetch(
        'https://api.github.com/repos/unitwebinfo-oss/unit-mail/releases/latest',
        {
          headers: {
            'User-Agent': `unit-mail/${app.getVersion()}`,
            Accept: 'application/vnd.github+json',
          },
        },
      );
      if (!res.ok) return null;
      const data = (await res.json()) as { tag_name?: string; html_url?: string };
      const latest = (data.tag_name || '').replace(/^v/i, '');
      const current = app.getVersion();
      return {
        current,
        latest,
        url: data.html_url || `https://github.com/unitwebinfo-oss/unit-mail/releases`,
        newer: !!latest && newerThan(latest, current),
      };
    } catch {
      return null; // sem internet / rate-limit — silencioso
    }
  });

  ipcMain.handle('app:info', () => ({
    name: 'Unit Mail',
    version: app.getVersion(),
    license: 'GPL-3.0-only',
    userData: dataDir(),
    // mailRootPath: só devolve o caminho — mailRoot() CRIARIA a pasta e
    // o renderer chama app:info no boot (antes do wizard escolher a raiz).
    documentsStoreDir: mailRootPath(),
    electron: process.versions.electron,
    chrome: process.versions.chrome,
    node: process.versions.node,
    platform: process.platform,
    secureStorage: secureStorageAvailable(),
  }));
  ipcMain.handle('app:openDataDir', async () => {
    const dir = dataDir();
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
    return shell.openPath(dir);
  });

  // Abre URL externa no navegador do SO — somente http/https, nada mais.
  ipcMain.handle('app:openExternal', (_e, url: string) => {
    if (typeof url === 'string' && /^https?:\/\//i.test(url)) {
      return shell.openExternal(url);
    }
  });

  // ---------- Pasta de dados (perfil portátil — modelo Thunderbird) ----------
  ipcMain.handle('app:getDataDir', () => dataDir());

  // Estado do perfil configurado: o ponteiro pode apontar para uma pasta
  // temporariamente ausente — ex.: drive removível (/run/media/…) que só
  // monta sob demanda. O renderer usa isso para mostrar o aviso "pasta de
  // dados não encontrada" em vez de abrir vazio num perfil fantasma.
  ipcMain.handle('app:profileStatus', () => {
    const target = profilePointerTarget();
    return { hasPointer: !!target, target, available: !!target && existsSync(target) };
  });

  // Passo 1: só abre o seletor e informa se a pasta já é um perfil —
  // NADA é gravado. A escolha só é aplicada em app:applyDataDir.
  ipcMain.handle('app:pickDataDirPath', async () => {
    const res = await dialog.showOpenDialog({
      title: 'Escolha a pasta de dados do Unit Mail',
      buttonLabel: 'Selecionar',
      properties: ['openDirectory', 'createDirectory'],
      defaultPath: app.getPath('documents'),
    });
    if (res.canceled || !res.filePaths[0]) return null;
    const dir = res.filePaths[0];
    return { dir, isProfile: isUnitMailProfile(dir) };
  });

  // Passo 2: aplica a pasta escolhida (grava ponteiro, migra dados se nova).
  ipcMain.handle('app:applyDataDir', (_e, dir: string) => {
    const { adopted } = setDataDir(dir);
    if (!adopted) {
      // Pasta nova: traz as bases locais existentes para <raiz>/Mail/ e
      // atualiza TODAS as referências (contas, regras, ordem de pastas…).
      for (const store of listLocalStores()) {
        const dest = join(mailRoot(), basename(store.path));
        if (resolve(store.path).toLowerCase() === resolve(dest).toLowerCase()) continue;
        if (!existsSync(store.path)) continue;
        try {
          // Mail/ já foi copiada inteira por setDataDir — só copia de novo
          // bases que ficavam FORA da pasta de dados antiga.
          if (!existsSync(dest)) cpSync(store.path, dest, { recursive: true });
          remapPstReferences(store.path, dest);
          unregisterLocalStore(store.path);
          registerLocalStore(dest);
        } catch {
          // cópia falhou — a base continua no caminho antigo
        }
      }
    }
    return { dir, adopted };
  });

  // Reinicia o app — usado quando uma pasta de dados com perfil é adotada:
  // marca o setup como concluído e relança para o boot restaurar tudo.
  ipcMain.handle('app:restart', () => {
    const s = loadSettings();
    s.setupComplete = true;
    saveSettings(s);
    app.relaunch();
    app.exit(0);
  });

  // Chamado quando o wizard termina: se nenhum arquivo de dados existe
  // (usuário pulou a configuração de conta), cria a base padrão "Unit Mail
  // Data" — como o Outlook, que cria um "Arquivo de Dados do Outlook" no
  // primeiro uso. Sem base, Calendário/Pessoas/Tarefas ficam sem destino.
  ipcMain.handle('app:ensureDefaultStore', () => {
    const s = loadSettings();
    const hasAny =
      s.openLocalStores.some((p) => existsSync(p)) ||
      s.openPstFiles.some((p) => existsSync(p));
    if (hasAny) return null;
    const def = createLocalStore('Unit Mail Data', mailRoot());
    addOpenLocalStore(def.path);
    adoptFirstAsDefault(def.path);
    return sourceManager.openLocalStore(def.path, def.name);
  });

  ipcMain.handle('dialog:openPst', async () => {
    const res = await dialog.showOpenDialog({
      title: 'Abrir arquivo PST',
      filters: [
        { name: 'Arquivo de dados do Outlook', extensions: ['pst'] },
        { name: 'Todos os arquivos', extensions: ['*'] },
      ],
      properties: ['openFile'],
    });
    return res.canceled ? null : res.filePaths[0];
  });

  // ---------- Bases (PST + contas) ----------
  ipcMain.handle('sources:list', () => sourceManager.list());

  ipcMain.handle('sources:openPst', (_e, filePath: string) => {
    const info = sourceManager.openPst(filePath);
    addRecent(filePath);
    return info;
  });

  ipcMain.handle(
    'items:move',
    (
      _e,
      srcId: string,
      srcFolderId: string,
      index: number,
      dstId: string,
      dstFolderId: string,
      stableId?: string,
    ) => sourceManager.moveItem(srcId, srcFolderId, index, dstId, dstFolderId, stableId),
  );

  // Em lote: excluir/mover centenas de itens numa chamada — estados em disco
  // gravados uma vez no fim (readstate/hidden.json), não por item.
  ipcMain.handle(
    'items:moveMany',
    (
      _e,
      srcId: string,
      srcFolderId: string,
      indexes: number[],
      dstId: string,
      dstFolderId: string,
      stableIds?: (string | undefined)[],
    ) => sourceManager.moveItems(srcId, srcFolderId, indexes, dstId, dstFolderId, stableIds),
  );

  ipcMain.handle(
    'items:isLocal',
    (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.isLocalItem !== 'function') return false;
      return pst.isLocalItem(folderId, index, stableId);
    },
  );

  ipcMain.handle(
    'items:delete',
    async (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.deleteItem !== 'function') {
        throw new Error('Excluir só é suportado em arquivos de dados (PST/base local)');
      }
      await pst.deleteItem(folderId, index, stableId);
      return true;
    },
  );

  ipcMain.handle(
    'items:deleteMany',
    (_e, sourceId: string, folderId: string, indexes: number[], stableIds?: (string | undefined)[]) =>
      sourceManager.deleteItems(sourceId, folderId, indexes, stableIds),
  );

  // ---------- Categorias (Categorizar) ----------
  ipcMain.handle('cats:list', () => loadCategories());
  ipcMain.handle(
    'cats:save',
    (_e, cats: import('../shared/types').MailCategory[]) => {
      saveCategories(cats);
      return true;
    },
  );
  ipcMain.handle(
    'items:getCategories',
    (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.getItemCategories !== 'function') return [];
      return pst.getItemCategories(folderId, index, stableId);
    },
  );
  ipcMain.handle(
    'items:setCategories',
    (
      _e,
      sourceId: string,
      folderId: string,
      indexes: number[],
      cats: string[],
      stableIds?: (string | undefined)[],
    ) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.setItemCategories !== 'function') return false;
      pst.batch(() => {
        indexes.forEach((ix, i) => {
          try {
            pst.setItemCategories(folderId, ix, cats, stableIds?.[i]);
          } catch {
            // item sumiu — segue com os demais
          }
        });
      });
      return true;
    },
  );

  ipcMain.handle(
    'folders:move',
    (
      _e,
      sourceId: string,
      folderId: string,
      newParentFolderId: string | null,
      dstSourceId?: string,
    ) => sourceManager.moveFolder(sourceId, folderId, newParentFolderId, dstSourceId),
  );

  ipcMain.handle(
    'folders:create',
    (_e, sourceId: string, name: string, parentFolderId: string | null) =>
      sourceManager.createFolder(sourceId, name, parentFolderId),
  );

  // Novo compromisso / contato / tarefa (Novos Itens da Página Inicial).
  // Compromisso com participantes → envia convite iMIP (METHOD:REQUEST) pela
  // conta vinculada à base de dados (ou a primeira configurada) — como o
  // Outlook faz ao salvar uma reunião com convidados.
  ipcMain.handle(
    'items:createSpecial',
    async (
      _e,
      sourceId: string,
      item: import('../shared/types').NewSpecialItem,
    ): Promise<import('../shared/types').CreateSpecialResult> => {
      const isAppt = item.kind === 'appt';
      const uid = isAppt ? `${randomUUID()}@unitmail` : undefined;
      // Organizador = a conta da base — vira o From:/ORGANIZER do .eml salvo
      // (antes saía sem From: e o "De" aparecia vazio na lista/inspetor).
      const orgAcct = isAppt && !item.organizer ? accountForStore(sourceId) : undefined;
      const withOrg = orgAcct
        ? { ...item, organizer: `${orgAcct.label || orgAcct.email} <${orgAcct.email}>` }
        : item;
      const tree = await sourceManager.createSpecialItem(sourceId, uid ? { ...withOrg, uid } : withOrg);
      if (isAppt) invalidateCalendarUids();
      const invite = await sendApptInvites(sourceId, uid ? { ...withOrg, uid } : withOrg);
      return { tree, invite };
    },
  );

  // Edita compromisso/contato/tarefa existente — regrava o .eml e, se o
  // compromisso tem participantes, reenvia o convite (mesmo UID, SEQUENCE+1
  // → os clientes atualizam a reunião em vez de duplicá-la).
  ipcMain.handle(
    'items:updateSpecial',
    async (
      _e,
      sourceId: string,
      folderId: string,
      index: number,
      item: import('../shared/types').NewSpecialItem,
      stableId?: string,
    ): Promise<import('../shared/types').CreateSpecialResult> => {
      const res = await sourceManager.updateSpecialItem(sourceId, folderId, index, item, stableId);
      if (res.item.kind === 'appt') invalidateCalendarUids();
      const invite = await sendApptInvites(sourceId, res.item);
      return { tree: res.tree, invite };
    },
  );

  // "Adiar" do lembrete — remove da lista de disparados e agenda novo disparo.
  // minutes <= 0 significa "dispensar de vez" — grava em disco para não
  // voltar a disparar após reiniciar o app dentro da janela do compromisso.
  ipcMain.handle('calendar:snooze', (_e, key: string, minutes: number) => {
    if (typeof key === 'string' && key) {
      if (minutes > 0) {
        remindedKeys.delete(key);
        snoozedUntil.set(key, Date.now() + Math.min(minutes, 24 * 60) * 60000);
      } else {
        loadDismissedReminders();
        dismissedReminders.add(key);
        saveDismissedReminders();
      }
    }
    return true;
  });

  ipcMain.handle('folders:rename', (_e, sourceId: string, folderId: string, newName: string) =>
    sourceManager.renameFolder(sourceId, folderId, newName),
  );
  ipcMain.handle('folders:delete', (_e, sourceId: string, folderId: string) =>
    sourceManager.deleteFolder(sourceId, folderId),
  );
  ipcMain.handle(
    'folders:setHidden',
    (_e, sourceId: string, folderId: string, hidden: boolean) =>
      sourceManager.setFolderHidden(sourceId, folderId, hidden),
  );
  ipcMain.handle('folders:copy', (_e, sourceId: string, folderId: string) =>
    sourceManager.copyFolder(sourceId, folderId),
  );
  ipcMain.handle('folders:empty', (_e, sourceId: string, folderId: string) =>
    sourceManager.emptyFolder(sourceId, folderId),
  );
  ipcMain.handle('folders:markAllRead', async (_e, sourceId: string, folderId: string) => {
    await sourceManager.markAllRead(sourceId, folderId);
    return true;
  });

  // Lista de remetentes bloqueados — mensagens deles caem direto em
  // "Lixo eletrônico" na próxima entrega (como o Blocked Senders do Outlook).
  ipcMain.handle('blocked:list', () => blockedSenders.list());
  ipcMain.handle('blocked:add', (_e, entry: string) => blockedSenders.add(entry));
  ipcMain.handle('blocked:remove', (_e, entry: string) => blockedSenders.remove(entry));

  // "Executar Regras" — varre a pasta em páginas, do fim para o começo (remover
  // um item só desloca índices maiores, já processados). Reporta progresso via
  // 'rules:progress', cede o event loop a cada lote (a janela não fica "não
  // respondendo") e pode ser cancelada por 'rules:cancelApply'.
  // Com ruleIds, roda exatamente a regra escolhida — mesmo desativada.
  // Sem ruleIds, usa as habilitadas (botão da barra de ações).
  let rulesApplyCancelled = false;
  ipcMain.on('rules:cancelApply', () => {
    rulesApplyCancelled = true;
  });
  ipcMain.handle(
    'rules:applyToFolder',
    async (e, sourceId: string, folderId: string, ruleIds?: string[]) => {
    const { matchRule } = await import('./ruleEngine');
    const svc = sourceManager.backend(sourceId);
    const pst = svc as import('./pstService').PstService;
    if (typeof pst.isLocalItem !== 'function') return { applied: 0, done: 0 };
    const all = ruleEngine.list();
    const rules = ruleIds?.length
      ? all.filter((r) => ruleIds.includes(r.id))
      : all.filter((r) => r.enabled);
    const sources = sourceManager.listAll();
    // E-mail por conta — o vínculo conta↔regra é checado mensagem a mensagem.
    const acctEmail = new Map(
      accountStore.list().map((a) => [a.id, (a.email || '').toLowerCase()]),
    );
    const total = await svc.getItemCount(folderId);
    let applied = 0;
    let done = 0;
    let targetMissing = 0;
    rulesApplyCancelled = false;
    const progress = () => {
      try {
        e.sender.send('rules:progress', { done, total, applied });
      } catch {
        // janela fechada no meio — segue
      }
    };
    const breathe = () => new Promise<void>((r) => setImmediate(r));
    const PAGE = 250;
    for (let end = total; end > 0 && !rulesApplyCancelled; ) {
      const start = Math.max(0, end - PAGE);
      const items = await svc.getItems(folderId, start, end - start);
      for (const it of [...items].sort((a, b) => b.index - a.index)) {
        if (rulesApplyCancelled) break;
        // Regras de email só valem para emails — compromissos/contatos/tarefas
        // (.eml com X-UM-Class) não entram na verificação, mas contam no
        // progresso (foram "verificados").
        const mc = it.messageClass || 'IPM.Note';
        if (mc.startsWith('IPM.Note')) {
        for (const rule of rules) {
          // O vínculo conta↔regra é uma condição em si: no recebimento a regra
          // só é avaliada para mensagens da conta dela. Na execução manual a
          // pasta pode ter mensagens de qualquer conta — então a regra da conta
          // X só casa itens em que o e-mail de X aparece (Para/Cc = recebida
          // por ela, De = enviada por ela). Sem isso, uma regra de conta sem
          // condições esvaziava a pasta inteira.
          if (rule.accountId && rule.accountId !== 'all') {
            const em = acctEmail.get(rule.accountId);
            if (!em) break; // conta removida — regra não verificável
            // O carimbo X-UM-Account (gravado na entrega) prova a conta mesmo
            // quando o cabeçalho To não mostra (BCC/listas). Itens sem carimbo
            // — importados de PST/Thunderbird — caem na checagem por
            // De/Para/Cc do cabeçalho.
            const stamped = (it.accountEmail || '').toLowerCase();
            if (stamped) {
              if (stamped !== em) continue;
            } else {
              const hay =
                `${it.displayTo} ${it.displayCC} ${it.senderEmailAddress}`.toLowerCase();
              if (!hay.includes(em)) continue;
            }
          }
          if (
            !matchRule(rule, {
              from: `${it.senderEmailAddress} ${it.senderName}`,
              to: `${it.displayTo} ${it.displayCC}`,
              // O resumo guarda '(sem assunto)' como placeholder — não é texto
              // real da mensagem. Sem isso, "Assunto contém sem/assunto" casava
              // com todas as mensagens sem assunto.
              subject: /^[(]sem (assunto|t[ií]tulo)[)]$/i.test(it.subject || '') ? '' : it.subject,
            })
          ) {
            continue;
          }
          const target = sources.find((s) => s.filePath === rule.targetPstPath);
          if (!target) {
            targetMissing++;
            break;
          }
          // Destino = a própria pasta varrida → já está onde a regra manda.
          // Mover só criaria uma cópia reimportada a cada execução.
          if (target.id === sourceId && `arc:${rule.targetFolder}` === folderId) break;
          try {
            await sourceManager.moveItem(
              sourceId,
              folderId,
              it.index,
              target.id,
              `arc:${rule.targetFolder}`,
              it.stableId,
            );
            applied++;
          } catch {
            // destino não gravável ou item já movido — segue
          }
          break; // moveu (ou tentou) — próxima mensagem
        }
        }
        done++;
        if (done % 10 === 0) {
          progress();
          await breathe();
        }
      }
      end = start;
      await breathe();
    }
    progress();
    return { applied, done, cancelled: rulesApplyCancelled, targetMissing };
  });

  // Ordem de exibição das pastas (persistida — permite reordenar arrastando,
  // inclusive em PSTs onde a ordem real não pode ser alterada)
  ipcMain.handle('settings:getFolderOrders', () => {
    const s = loadSettings();
    const orders = s.folderOrders ?? {};
    // Limpa chaves mortas de versões antigas (sourceId aleatório estilo uuid
    // nunca mais casa — o id atual é sha1 estável do caminho).
    let changed = false;
    for (const k of Object.keys(orders)) {
      const src = k.split('|')[0];
      if (/^(pst|local|acct):[0-9a-f]{8}-[0-9a-f]{4}-/.test(src)) {
        delete orders[k];
        changed = true;
      }
    }
    if (changed) {
      s.folderOrders = orders;
      saveSettings(s);
    }
    return orders;
  });
  ipcMain.handle(
    'settings:setFolderOrder',
    (_e, key: string, orderedIds: string[]) => {
      const s = loadSettings();
      s.folderOrders = { ...(s.folderOrders ?? {}), [key]: orderedIds };
      saveSettings(s);
      return true;
    },
  );

  ipcMain.handle('sources:closeSource', async (_e, sourceId: string) => {
    const entry = sourceManager.list().find((s) => s.id === sourceId);
    await sourceManager.close(sourceId);
    if (entry?.filePath) {
      if (entry.kind === 'pst') removeOpenPst(entry.filePath);
      else if (entry.kind === 'local') {
        removeOpenLocalStore(entry.filePath);
        // Fechar uma base local a remove do app de verdade — sai do registro
        // (localstores.json) e deixa de ser alvo de regras/entrega. Os arquivos
        // ficam no disco; o usuário pode reabrir em "Abrir base local existente".
        unregisterLocalStore(entry.filePath);
      }
    }
    return true;
  });

  // ---------- Bases de dados locais ("novo arquivo de dados") ----------
  ipcMain.handle('stores:createLocal', async (_e, name: string) => {
    const res = await dialog.showOpenDialog({
      title: 'Escolha a pasta onde a base de dados será criada',
      properties: ['openDirectory', 'createDirectory'],
    });
    if (res.canceled || !res.filePaths[0]) return null;
    const def = createLocalStore(name, res.filePaths[0]);
    addOpenLocalStore(def.path);
    adoptFirstAsDefault(def.path);
    return sourceManager.openLocalStore(def.path, def.name);
  });

  ipcMain.handle('stores:renameLocal', (_e, path: string, newName: string) => {
    const def = renameLocalStore(path, newName);
    if (!def) throw new Error('Base local não encontrada');
    // Atualiza o label da fonte aberta, se estiver aberta
    const open = sourceManager.list().find((s) => s.kind === 'local' && s.filePath === path);
    if (open) open.label = newName;
    return def;
  });

  // Cria base local na pasta de dados do app (<raiz>/Mail/<nome>) — modelo
  // Thunderbird: cada conta tem sua própria pasta dentro de Mail/.
  ipcMain.handle('stores:createLocalAuto', async (_e, name: string) => {
    const def = createLocalStore(name, mailRoot());
    addOpenLocalStore(def.path);
    adoptFirstAsDefault(def.path);
    return sourceManager.openLocalStore(def.path, def.name);
  });

  // Cria base local num diretório escolhido pelo usuário (sem diálogo — a
  // pasta já foi escolhida antes via stores:pickStoreDir)
  ipcMain.handle('stores:createLocalIn', async (_e, name: string, parentDir: string) => {
    mkdirSync(parentDir, { recursive: true });
    const def = createLocalStore(name, parentDir);
    addOpenLocalStore(def.path);
    adoptFirstAsDefault(def.path);
    return sourceManager.openLocalStore(def.path, def.name);
  });

  // Migra TODO o conteúdo de um PST (ou base aberta) para uma base nativa —
  // depois disso, excluir apaga de verdade e o PST original fica intocado.
  ipcMain.handle(
    'stores:migratePst',
    async (_e, sourceId: string, destParentDir: string, name: string, existingStorePath?: string, folderIds?: string[]) => {
      const info = await migratePstToLocalStore(
        sourceId,
        destParentDir || mailRoot(),
        name,
        (p) => {
          mainWindow?.webContents.send('migrate:progress', p);
        },
        existingStorePath,
        folderIds ? new Set(folderIds) : undefined,
      );
      addOpenLocalStore(info.filePath!);
      return info;
    },
  );

  // Importa as mensagens (.mbox) de um perfil do Thunderbird para uma base
  // nativa — mesmo fluxo da migração de PST (progresso em 'migrate:progress').
  ipcMain.handle('thunderbirdImport:scanMail', () => scanThunderbirdMail());
  ipcMain.handle(
    'stores:importThunderbird',
    async (_e, profilePath: string, destParentDir: string, name: string, existingStorePath?: string) => {
      const info = await importThunderbirdMail(
        profilePath,
        destParentDir || mailRoot(),
        name,
        (p) => {
          mainWindow?.webContents.send('migrate:progress', p);
        },
        existingStorePath,
      );
      if (!existingStorePath) addOpenLocalStore(info.filePath!);
      return info;
    },
  );

  // Abre o seletor de pasta para o usuário escolher onde criar a base local
  ipcMain.handle('stores:pickStoreDir', async () => {
    const res = await dialog.showOpenDialog({
      title: 'Escolha a pasta onde o arquivo de dados será criado',
      properties: ['openDirectory', 'createDirectory'],
      defaultPath: app.getPath('documents'),
    });
    if (res.canceled || !res.filePaths[0]) return null;
    return res.filePaths[0];
  });

  ipcMain.handle('stores:openLocal', async (_e) => {
    const res = await dialog.showOpenDialog({
      title: 'Abrir base de dados local (pasta criada pelo app)',
      properties: ['openDirectory'],
    });
    if (res.canceled || !res.filePaths[0]) return null;
    const dir = res.filePaths[0];
    const def = registerLocalStore(dir);
    addOpenLocalStore(def!.path);
    return sourceManager.openLocalStore(def!.path, def!.name);
  });

  ipcMain.handle('app:getRecentFiles', () => loadSettings().recentFiles);
  ipcMain.handle('app:getLastOpened', () => loadSettings().lastOpened || null);

  // ---------- Contas de email (IMAP/SMTP) ----------
  ipcMain.handle('accounts:list', () => accountStore.list());

  ipcMain.handle('accounts:upsert', async (_e, draft: AccountDraft) => {
    // O dropdown "entregar mensagens em" grava direto em account.deliveryPath/
    // deliveryFolder — o motor de regras já usa esses campos como destino
    // padrão (ensureDeliveryStore/imapInboxDir), então NÃO é preciso criar
    // uma regra visível para isso (evita poluir a aba Regras).
    const saved = accountStore.upsert(draft);
    // A sessão aberta guarda a senha e o servidor da época em que conectou —
    // fecha para que a próxima conexão use o que acabou de ser salvo.
    await sourceManager.close(`acct:${saved.id}`);
    return saved;
  });

  ipcMain.handle('accounts:remove', async (_e, accountId: string, deleteStore?: boolean) => {
    const acc = accountStore.getFull(accountId);
    await sourceManager.close(`acct:${accountId}`);
    accountStore.remove(accountId);
    // Base local vinculada só à conta removida — apaga de verdade (nunca um
    // .pst: o arquivo original do usuário nunca é excluído pelo app).
    if (deleteStore && acc?.deliveryPath && isLocalStorePath(acc.deliveryPath)) {
      await sourceManager.close(sourceManager.listAll().find((s) => s.filePath === acc.deliveryPath)?.id || '');
      deleteLocalStore(acc.deliveryPath);
    }
    return true;
  });

  // Apaga de verdade uma base de dados local (nunca um .pst — esse só "fecha").
  ipcMain.handle('stores:deleteLocal', async (_e, storePath: string) => {
    const src = sourceManager.listAll().find((s) => s.filePath === storePath);
    if (src) await sourceManager.close(src.id);
    deleteLocalStore(storePath);
    return true;
  });

  ipcMain.handle('accounts:test', async (_e, draft: AccountDraft) => {
    try {
      const { message, imapSecureFix, allowSelfSignedFix } = await testAccount(
        {
          id: draft.id || 'tmp',
          label: draft.label,
          email: draft.email,
          protocol: draft.protocol || 'imap',
          signature: draft.signature || '',
          signatureNewId: draft.signatureNewId,
          signatureReplyId: draft.signatureReplyId,
          deliveryPath: draft.deliveryPath,
          leaveOnServer: draft.leaveOnServer ?? true,
          leaveDays: draft.leaveDays ?? 0,
          allowSelfSigned: draft.allowSelfSigned,
          imapHost: draft.imapHost,
          imapPort: draft.imapPort,
          imapSecure: draft.imapSecure,
          imapUser: draft.imapUser,
          smtpHost: draft.smtpHost,
          smtpPort: draft.smtpPort,
          smtpSecure: draft.smtpSecure,
          smtpUser: draft.smtpUser,
        },
        // Se o campo de senha ficou em branco numa conta já salva, usa a senha armazenada
        draft.password || (draft.id ? accountStore.getPassword(draft.id) : ''),
        draft.smtpPassword ||
          (draft.id ? accountStore.getSmtpPassword(draft.id) : '') ||
          undefined,
      );
      return { ok: true, message, imapSecureFix, allowSelfSignedFix };
    } catch (e) {
      return { ok: false, message: (e as Error).message };
    }
  });

  ipcMain.handle('accounts:connect', (_e, accountId: string) => sourceManager.openAccount(accountId));

  ipcMain.handle('accounts:send', async (_e, input: ComposeInput) => {
    const accountId = input.accountId;
    const account = accountStore.getFull(accountId);
    if (!account) throw new Error('Conta não encontrada');
    const password = accountStore.getPassword(accountId);

    // Gera o .eml da mensagem para a cópia local (Caixa de Saída / Itens Enviados)
    let raw: Buffer | null = null;
    try {
      raw = await composeRawMessage(account, input);
    } catch {
      // cópia local é best-effort — não impede o envio
    }

    // Fluxo do Outlook: clicar em Enviar fecha o modal na hora — a mensagem
    // vai para a Caixa de Saída (já marcada como LIDA, itens enviados nunca
    // aparecem como não lidos) e o envio acontece em segundo plano.
    let svc: PstService | null = null;
    let outboxFile: string | null = null;
    // Fora do try — o envio assíncrono abaixo também usa (Caixa de Saída →
    // Itens Enviados).
    let store = '';
    try {
      store = ensureDeliveryStore(account);
      // Sem base de entrega configurada ('') pula a cópia local — o envio
      // continua independente disso.
      if (store) {
      svc = sourceManager.serviceForPath(store);
      if (!svc) {
        // Base ainda não aberta — abre para gravar com estado de leitura
        const info = store.toLowerCase().endsWith('.pst')
          ? sourceManager.openPst(store)
          : sourceManager.openLocalStore(store);
        const b = sourceManager.backend(info.id);
        if (b instanceof PstService) svc = b;
      }
      if (raw && svc) {
        outboxFile = svc.importNamedEml(resolveStoreFolderName(store, 'outbox'), raw, true);
      } else if (raw) {
        // Fallback sem serviço: grava direto (o seed do readstate marca
        // arquivos antigos como lidos na primeira abertura da base)
        const dir = kindDirFor(store, 'outbox');
        mkdirSync(dir, { recursive: true });
        writeFileSync(join(dir, `${Date.now()}-${randomUUID().slice(0, 8)}.eml`), raw);
      }
      }
    } catch {
      // cópia local best-effort — não impede o envio
    }

    // Envio SMTP em segundo plano — o renderer já fechou o modal
    void (async () => {
      try {
        await sendMail(account, password, input, accountStore.getSmtpPassword(accountId));
        // Marca a mensagem ORIGINAL como respondida/encaminhada no
        // servidor IMAP (flag \Answered/$Forwarded — o ícone ↩/→ da lista).
        if (input.replyRef && (input.kind === 'reply' || input.kind === 'forward')) {
          try {
            const src = sourceManager.backend(input.replyRef.sourceId);
            if (src instanceof ImapMailSource) {
              await src.addItemFlags(input.replyRef.folderId, input.replyRef.index, [
                input.kind === 'forward' ? '$Forwarded' : '\\Answered',
              ]);
            }
          } catch {
            // marcação best-effort — não impede o envio
          }
        }
        // Sai da Caixa de Saída → Itens Enviados, como no Outlook
        if (svc && outboxFile) {
          svc.moveNamedFile(
            resolveStoreFolderName(store, 'outbox'),
            resolveStoreFolderName(store, 'sent'),
            outboxFile,
          );
        } else if (raw) {
          try {
            let copied = false;
            // IMAP: a cópia de enviados fica no servidor (pasta \Sent) —
            // nenhum arquivo local é criado para isso.
            if (account.protocol === 'imap') {
              try {
                const be = sourceManager.backend(`acct:${accountId}`);
                if (be instanceof ImapMailSource) {
                  const sentPath = await be.specialFolderPath('sent');
                  if (sentPath) {
                    await be.appendItem(sentPath, raw, true);
                    copied = true;
                  }
                }
              } catch {
                // servidor inacessível/sem pasta de enviados — cai no local
              }
            }
            if (!copied) {
              const sentStore =
                account.protocol === 'imap'
                  ? loadSettings().defaultStorePath || ''
                  : ensureDeliveryStore(account);
              if (sentStore) {
                const dir = kindDirFor(sentStore, 'sent');
                mkdirSync(dir, { recursive: true });
                writeFileSync(join(dir, `${Date.now()}-${randomUUID().slice(0, 8)}.eml`), raw);
              }
            }
          } catch {
            // ignora falha de cópia local
          }
        }
        // Cache de contatos: destinatários entram no autocompletar
        recordRecipients([input.to, input.cc, input.bcc].filter(Boolean).join(';'));
        mainWindow?.webContents.send('mail:sendResult', { ok: true, subject: input.subject });
      } catch (e) {
        // Falhou → fica na Caixa de Saída, como no Outlook (pode reenviar depois)
        const msg = (e as Error).message || String(e);
        mainWindow?.webContents.send('mail:sendResult', {
          ok: false,
          subject: input.subject,
          error: msg,
        });
        showToast(
          'error',
          'Falha ao enviar mensagem',
          `"${input.subject || '(sem assunto)'}" ficou na Caixa de Saída: ${msg}`,
        );
      }
    })();
    return true;
  });

  // Autocompletar de contatos — histórico de endereços enviados
  ipcMain.handle('contacts:suggest', (_e, q: string) => searchContacts(q));

  ipcMain.handle('dialog:pickAttachments', async () => {
    const res = await dialog.showOpenDialog({
      title: 'Anexar arquivos',
      properties: ['openFile', 'multiSelections'],
    });
    return res.canceled ? [] : res.filePaths;
  });

  // Imagem para inserir no corpo (assinatura ou mensagem) — retorna data URI
  ipcMain.handle('dialog:pickImage', async () => {
    const res = await dialog.showOpenDialog({
      title: 'Inserir imagem',
      filters: [{ name: 'Imagens', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'] }],
      properties: ['openFile'],
    });
    if (res.canceled || !res.filePaths[0]) return null;
    const p = res.filePaths[0];
    const ext = p.split('.').pop()?.toLowerCase() || 'png';
    const mime = ext === 'svg' ? 'image/svg+xml' : `image/${ext === 'jpg' ? 'jpeg' : ext}`;
    const buf = readFileSync(p);
    return `data:${mime};base64,${buf.toString('base64')}`;
  });

  // ---------- Regras de recebimento ----------
  ipcMain.handle('rules:list', () => ruleEngine.list());
  ipcMain.handle('rules:upsert', (_e, rule: Omit<MailRule, 'id'> & { id?: string }) =>
    ruleEngine.upsert(rule),
  );
  ipcMain.handle('rules:remove', (_e, id: string) => {
    ruleEngine.remove(id);
    return true;
  });
  // Resolve a assinatura da conta por tipo de mensagem (para o editor mostrar)
  ipcMain.handle('accounts:signatureFor', (_e, accountId: string, kind?: 'new' | 'reply' | 'forward') => {
    const account = accountStore.getFull(accountId);
    if (!account) return '';
    return resolveSignatureHtml(account, kind);
  });

  // ---------- Convites de reunião (Aceitar/Recusar/Propor novo horário) ----------
  ipcMain.handle(
    'mail:respondInvite',
    async (
      _e,
      input: {
        sourceId: string;
        folderId: string;
        index: number;
        stableId?: string;
        accountId: string;
        response: 'accept' | 'decline' | 'counter';
        newStart?: string;
        newEnd?: string;
      },
    ) => {
      const svc = sourceManager.backend(input.sourceId);
      const detail = await svc.getItem(input.folderId, input.index, input.stableId);
      const inv = detail.invite;
      if (!inv) throw new Error('A mensagem não é um convite de reunião');
      const account = accountStore.getFull(input.accountId);
      if (!account) throw new Error('Conta não encontrada');
      const password = accountStore.getPassword(input.accountId);

      const icsDate = (iso?: string | null): string =>
        iso ? new Date(iso).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '') : '';
      const fold = (s: string) =>
        s.replace(/([\\;,])/g, '\\$1').replace(/\r?\n/g, '\\n');
      const orgName = inv.organizer ? `CN=${inv.organizer}:` : '';
      const toOrg = inv.organizerEmail || inv.organizer;
      if (!toOrg) throw new Error('Convite sem organizador para responder');

      const isCounter = input.response === 'counter';
      const partstat = input.response === 'accept' ? 'ACCEPTED' : 'DECLINED';
      const method = isCounter ? 'COUNTER' : 'REPLY';
      const dtStart = icsDate(isCounter ? input.newStart : inv.start);
      const dtEnd = icsDate(isCounter ? input.newEnd : inv.end);
      const ics = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//OutlookLinux//Calendario//PT-BR',
        `METHOD:${method}`,
        'BEGIN:VEVENT',
        `UID:${inv.uid || `ol-${Date.now()}`}`,
        'SEQUENCE:0',
        `SUMMARY:${fold(inv.summary || detail.subject)}`,
        dtStart && `DTSTART:${dtStart}`,
        dtEnd && `DTEND:${dtEnd}`,
        inv.location && `LOCATION:${fold(inv.location)}`,
        `ORGANIZER;${orgName}mailto:${toOrg}`,
        `ATTENDEE;PARTSTAT=${isCounter ? 'TENTATIVE' : partstat};CN="${account.label || account.email}";RSVP=FALSE:mailto:${account.email}`,
        'END:VEVENT',
        'END:VCALENDAR',
      ]
        .filter(Boolean)
        .join('\r\n');

      const verb =
        input.response === 'accept'
          ? 'Aceito'
          : input.response === 'decline'
            ? 'Recusado'
            : 'Nova Proposta de Horário';
      await sendMail(account, password, {
        accountId: input.accountId,
        kind: 'reply',
        to: toOrg,
        subject: `${verb}: ${inv.summary || detail.subject}`,
        html: `<p>${verb === 'Nova Proposta de Horário' ? `Proponho ${dtStart ? new Date(input.newStart!).toLocaleString('pt-BR') : ''}${dtEnd ? ` até ${new Date(input.newEnd!).toLocaleString('pt-BR')}` : ''}.` : `${verb} o convite "${inv.summary || detail.subject}".`}</p>`,
        rawAttachments: [
          {
            filename: 'invite.ics',
            contentBase64: Buffer.from(ics, 'utf8').toString('base64'),
            contentType: `text/calendar; method=${method}`,
          },
        ],
      }, accountStore.getSmtpPassword(input.accountId));

      // Resposta enviada com sucesso → marca o convite (por UID, estável
      // entre pastas/sessões) — a barra deixa de oferecer os botões.
      markInviteResponded(inv.uid || `${input.sourceId}|${input.folderId}|${input.index}`, input.response);

      // Comportamento padrão do Outlook ("excluir convites de reunião após
      // responder"): a mensagem do convite sai da listagem e vai para Itens
      // Excluídos — o compromisso vive no Calendário e nos lembretes.
      try {
        const src = sourceManager.listAll().find((s) => s.id === input.sourceId);
        const kind = mailFolderKind(input.sourceId, input.folderId);
        if (src?.filePath && kind !== 'deleted' && kind !== 'calendar') {
          const delName = resolveStoreFolderName(src.filePath, 'deleted');
          await sourceManager.moveItem(
            input.sourceId,
            input.folderId,
            input.index,
            input.sourceId,
            `arc:${delName}`,
            input.stableId,
          );
        }
      } catch {
        // remoção é best-effort — a resposta já foi enviada
      }

      // Aceito → grava o compromisso na pasta Calendário da base (local/PST-sombra)
      if (input.response === 'accept') {
        const eventIcs =
          inv.ics ||
          [
            'BEGIN:VCALENDAR',
            'VERSION:2.0',
            'METHOD:REQUEST',
            'BEGIN:VEVENT',
            `UID:${inv.uid || `ol-${Date.now()}`}`,
            `SUMMARY:${fold(inv.summary || detail.subject)}`,
            dtStart && `DTSTART:${dtStart}`,
            dtEnd && `DTEND:${dtEnd}`,
            inv.location && `LOCATION:${fold(inv.location)}`,
            `ORGANIZER;${orgName}mailto:${toOrg}`,
            'BEGIN:VALARM',
            'TRIGGER:-PT15M',
            'ACTION:DISPLAY',
            'END:VALARM',
            'END:VEVENT',
            'END:VCALENDAR',
          ]
            .filter(Boolean)
            .join('\r\n');
        const eml = [
          `From: ${toOrg}`,
          `To: ${account.email}`,
          `Subject: ${inv.summary || detail.subject}`,
          `Date: ${new Date().toUTCString()}`,
          'MIME-Version: 1.0',
          'Content-Type: multipart/mixed; boundary="ol-cal"',
          '',
          '--ol-cal',
          'Content-Type: text/plain; charset=utf-8',
          '',
          `Compromisso: ${inv.summary || detail.subject}`,
          inv.location ? `Local: ${inv.location}` : '',
          '',
          '--ol-cal',
          'Content-Type: text/calendar; method=REQUEST; name="invite.ics"',
          'Content-Disposition: attachment; filename="invite.ics"',
          '',
          eventIcs,
          '--ol-cal--',
        ].join('\r\n');
        // Compromissos sempre vão para o Calendário do arquivo de dados
        // principal — mesmo que o convite tenha chegado em outra base.
        const defPath = loadSettings().defaultStorePath;
        const defSrc = defPath
          ? sourceManager.listAll().find((s) => s.filePath === defPath)
          : undefined;
        const pst = (defSrc ? sourceManager.backend(defSrc.id) : svc) as import('./pstService').PstService;
        try {
          if (typeof pst.importItemEml === 'function') {
            // Pasta de calendário já existente (qualquer idioma) ou nome localizado
            const calName = resolveStoreFolderName(defSrc?.filePath || '', 'calendar');
            await pst.importItemEml(`arc:${calName}`, Buffer.from(eml, 'utf8'));
            invalidateCalendarUids();
          }
        } catch {
          // base não gravável — a resposta já foi enviada
        }
      }
      return { sent: true };
    },
  );

  ipcMain.handle('accounts:sync', async (_e, accountId: string) => {
    ruleEngine.onProgress = (ev) => mainWindow?.webContents.send('sync:progress', ev);
    const r = await ruleEngine.syncAccount(accountId);
    mainWindow?.webContents.send('sync:done', [r]);
    return r;
  });
  ipcMain.handle('accounts:syncAll', async () => {
    ruleEngine.onProgress = (ev) => mainWindow?.webContents.send('sync:progress', ev);
    const r = await ruleEngine.syncAll();
    mainWindow?.webContents.send('sync:done', r);
    return r;
  });
  ipcMain.handle('accounts:syncCancel', () => {
    ruleEngine.cancelSync();
    return true;
  });
  ipcMain.handle('accounts:syncCancelTask', (_e, taskId: string) => {
    ruleEngine.cancelTask(taskId);
    return true;
  });
  // "Re-sincronizar do servidor" — esquece o que já foi processado para a
  // conta e re-baixa tudo que ainda existe no servidor (a dedupe por pasta
  // impede duplicar). Recuperação de mensagens marcadas como entregues por
  // engano (ex.: diretório "D:\…" fantasma num backup restaurado).
  ipcMain.handle('accounts:resetSync', (_e, accountId: string) => {
    ruleEngine.resetSyncState(accountId);
    return true;
  });
  ipcMain.handle('settings:getSyncInterval', () => loadSettings().syncIntervalMin ?? 30);
  ipcMain.handle('settings:getLoadImages', () => loadSettings().loadRemoteImages ?? false);
  ipcMain.handle('settings:getSetupComplete', () => {
    const s = loadSettings();
    if (s.setupComplete) return true;
    // O wizard de primeira execução não faz sentido quando o perfil já está
    // de pé: ponteiro gravado (pasta de dados já escolhida uma vez — cobre
    // quem fechou o app no meio do wizard, para quem a tela reaparecia a
    // cada abertura), conta configurada ou base registrada (perfil
    // restaurado de outro PC ou configurado direto em Configurações).
    const done =
      profilePointerExists() ||
      accountStore.list().length > 0 ||
      s.openPstFiles.length > 0 ||
      s.openLocalStores.length > 0 ||
      listLocalStores().length > 0;
    if (done) {
      // Persiste só quando a pasta de dados existe de fato — ponteiro com a
      // pasta temporariamente ausente (drive desmontado) não deve criar um
      // diretório fantasma ao gravar o settings.json.
      if (existsSync(dataDir())) {
        s.setupComplete = true;
        saveSettings(s);
      }
    }
    return done;
  });
  ipcMain.handle('settings:setSetupComplete', (_e, v: boolean) => {
    const s = loadSettings();
    s.setupComplete = !!v;
    saveSettings(s);
    return s.setupComplete;
  });
  ipcMain.handle('settings:setLoadImages', (_e, v: boolean) => {
    const s = loadSettings();
    s.loadRemoteImages = !!v;
    saveSettings(s);
    return s.loadRemoteImages;
  });
  ipcMain.handle('settings:getVirusScan', () => loadSettings().virusScan ?? true);
  ipcMain.handle('settings:setVirusScan', (_e, v: boolean) => {
    const s = loadSettings();
    s.virusScan = !!v;
    saveSettings(s);
    return s.virusScan;
  });
  // Qual motor de antivírus local está disponível (AMSI/Windows, ClamAV/Linux)
  ipcMain.handle('antivirus:status', () => avStatus());
  ipcMain.handle('settings:setSyncInterval', (_e, min: number) => {
    const s = loadSettings();
    s.syncIntervalMin = Math.max(0, Math.min(240, Math.floor(min || 0)));
    saveSettings(s);
    scheduleAutoSync();
    return s.syncIntervalMin;
  });
  // Arquivo de dados principal — destino padrão de mensagens sem regra
  ipcMain.handle(
    'items:setRead',
    async (_e, sourceId: string, folderId: string, index: number, read: boolean, stableId?: string) => {
      await sourceManager.setItemRead(sourceId, folderId, index, read, stableId);
      return true;
    },
  );

  ipcMain.handle('settings:getDefaultStore', () => loadSettings().defaultStorePath ?? '');
  ipcMain.handle('settings:setDefaultStore', (_e, p: string) => {
    const s = loadSettings();
    s.defaultStorePath = p || undefined;
    saveSettings(s);
    return s.defaultStorePath ?? '';
  });
  // Tema claro/escuro — 'system' segue o SO (nativeTheme governa prefers-color-scheme)
  ipcMain.handle('settings:getTheme', () => loadSettings().theme ?? 'system');
  ipcMain.handle('settings:setTheme', (_e, t: 'system' | 'light' | 'dark') => {
    const s = loadSettings();
    s.theme = t === 'light' || t === 'dark' ? t : 'system';
    saveSettings(s);
    nativeTheme.themeSource = s.theme;
    return s.theme;
  });
  // Idioma da interface — decide nomes de pastas NOVAS e a exibição; pastas
  // já criadas no disco nunca são renomeadas.
  // Sem escolha salva → appLanguage() detecta o idioma do SO (pt-* → pt, resto → en)
  ipcMain.handle('settings:getLanguage', () => appLanguage());
  ipcMain.handle('settings:setLanguage', (_e, l: 'pt' | 'en') => {
    const s = loadSettings();
    s.language = l === 'en' ? 'en' : 'pt';
    saveSettings(s);
    return s.language;
  });

  // ---------- Assinaturas ----------
  ipcMain.handle('signatures:list', () => signatureStore.list());
  ipcMain.handle(
    'signatures:upsert',
    (_e, sig: Omit<Signature, 'id'> & { id?: string }) => signatureStore.upsert(sig),
  );
  ipcMain.handle('signatures:remove', (_e, id: string) => {
    signatureStore.remove(id);
    return true;
  });

  // ---------- Importação de contas do Outlook (Windows) ----------
  ipcMain.handle('outlookImport:available', () => isWindowsImportAvailable());
  ipcMain.handle('outlookImport:scan', () => importOutlookAccountsFromRegistry());
  ipcMain.handle('thunderbirdImport:scan', () => scanThunderbirdAccounts());
  ipcMain.handle('outlookImport:saveExport', async (_e, guesses) => {
    const res = await dialog.showSaveDialog({
      title: 'Salvar exportação de contas do Outlook',
      defaultPath: 'outlook-accounts-export.json',
      filters: [{ name: 'JSON', extensions: ['json'] }],
    });
    if (res.canceled || !res.filePath) return null;
    // Nunca grava senha em texto puro num arquivo portátil (pendrive, nuvem…) —
    // a senha recuperada só preenche o formulário nesta sessão/máquina.
    const safe = (guesses as Array<Record<string, unknown>>).map(
      ({ password: _p, smtpPassword: _sp, ...rest }) => rest,
    );
    await writeFile(res.filePath, JSON.stringify(safe, null, 2));
    return res.filePath;
  });
  ipcMain.handle('outlookImport:openExportFile', async () => {
    const res = await dialog.showOpenDialog({
      title: 'Importar exportação de contas do Outlook',
      filters: [{ name: 'JSON', extensions: ['json'] }],
      properties: ['openFile'],
    });
    if (res.canceled || !res.filePaths[0]) return null;
    let list: ImportedAccountGuess[];
    try {
      list = JSON.parse(readFileSync(res.filePaths[0], 'utf-8'));
    } catch {
      throw new Error('Arquivo de exportação inválido');
    }
    // O arquivo nunca traz senhas (por segurança). No Windows, as senhas são
    // recuperadas do Registro para as contas de mesmo e-mail — só em memória.
    if (process.platform === 'win32') {
      try {
        const live = await importOutlookAccountsFromRegistry();
        for (const g of list) {
          const m = live.find((x) => x.email.toLowerCase() === String(g.email).toLowerCase());
          if (m) {
            g.password = m.password;
            g.smtpPassword = m.smtpPassword;
          }
        }
      } catch {
        // sem leitura do Registro — contas entram sem senha
      }
    }
    return list;
  });

  // ---------- Backup de configurações (contas + regras + assinaturas) ----------
  /** Exporta as configurações do app para um arquivo JSON — migração entre máquinas. */
  ipcMain.handle('backup:export', async () => {
    const res = await dialog.showSaveDialog({
      title: 'Exportar configurações do Unit Mail',
      defaultPath: 'unit-mail-backup.json',
      filters: [{ name: 'Backup do Unit Mail', extensions: ['json'] }],
    });
    if (res.canceled || !res.filePath) return null;
    // Senhas exportadas em texto puro — o arquivo de backup é sensível,
    // igual a um arquivo .pst com as mensagens.
    const accounts = accountStore.list().map((a) => ({
      ...a,
      password: accountStore.getPassword(a.id),
      smtpPassword: a.hasSmtpPassword ? accountStore.getSmtpPassword(a.id) : undefined,
    }));
    const payload = {
      app: 'unit-mail',
      version: 1,
      exportedAt: new Date().toISOString(),
      accounts,
      rules: ruleEngine.list(),
      signatures: signatureStore.list(),
      settings: {
        syncIntervalMin: loadSettings().syncIntervalMin ?? 30,
        loadRemoteImages: loadSettings().loadRemoteImages ?? false,
        theme: loadSettings().theme ?? 'system',
        language: appLanguage(),
      },
    };
    await writeFile(res.filePath, JSON.stringify(payload, null, 2));
    return res.filePath;
  });

  /** Importa um backup: mescla contas (por email), regras e assinaturas (por id). */
  ipcMain.handle('backup:import', async () => {
    const res = await dialog.showOpenDialog({
      title: 'Importar backup do Unit Mail',
      filters: [{ name: 'Backup do Unit Mail', extensions: ['json'] }],
      properties: ['openFile'],
    });
    if (res.canceled || !res.filePaths[0]) return null;
    let data: {
      accounts?: Array<AccountDraft & { id?: string; password?: string }>;
      rules?: MailRule[];
      signatures?: Signature[];
      settings?: { syncIntervalMin?: number; loadRemoteImages?: boolean; theme?: 'system' | 'light' | 'dark'; language?: 'pt' | 'en' };
    };
    try {
      data = JSON.parse(readFileSync(res.filePaths[0], 'utf-8'));
    } catch {
      throw new Error('Arquivo de backup inválido');
    }
    const summary = { accounts: 0, rules: 0, signatures: 0, warnings: [] as string[] };

    for (const acc of data.accounts ?? []) {
      try {
        // Senha importada só vale se veio em texto puro no backup
        accountStore.upsert({
          ...acc,
          id: acc.id,
          password: acc.password || '',
          smtpPassword: acc.smtpPassword || undefined,
        });
        summary.accounts++;
      } catch {
        summary.warnings.push(`Conta ${acc.email}: não foi possível importar`);
      }
    }
    for (const rule of data.rules ?? []) {
      try {
        ruleEngine.upsert(rule);
        summary.rules++;
      } catch {
        summary.warnings.push(`Regra "${rule.name}": não foi possível importar`);
      }
    }
    for (const sig of data.signatures ?? []) {
      try {
        signatureStore.upsert(sig);
        summary.signatures++;
      } catch {
        summary.warnings.push(`Assinatura "${sig.name}": não foi possível importar`);
      }
    }
    if (data.settings) {
      const s = loadSettings();
      if (data.settings.syncIntervalMin !== undefined) s.syncIntervalMin = data.settings.syncIntervalMin;
      if (data.settings.loadRemoteImages !== undefined) s.loadRemoteImages = data.settings.loadRemoteImages;
      if (data.settings.theme) s.theme = data.settings.theme;
      if (data.settings.language) s.language = data.settings.language;
      saveSettings(s);
      if (data.settings.theme) nativeTheme.themeSource = data.settings.theme;
    }
    // deliveryPath do backup pode apontar para base que ainda não existe nesta
    // máquina — avisa o usuário para re-vincular em Arquivos de Dados.
    const missingBase = accountStore
      .list()
      .filter((a) => a.deliveryPath && !existsSync(a.deliveryPath));
    if (missingBase.length) {
      summary.warnings.push(
        `${missingBase.length} conta(s) apontam para arquivos de dados que não existem nesta máquina — adicione as bases e re-vincule em Configurações → Email.`,
      );
    }
    return summary;
  });

  // ---------- Operações de email (unificadas por sourceId) ----------
  ipcMain.handle('mail:folderTree', (_e, sourceId: string) => sourceManager.backend(sourceId).getFolderTree());

  ipcMain.handle('mail:itemCount', (_e, sourceId: string, folderId: string) =>
    sourceManager.backend(sourceId).getItemCount(folderId),
  );

  ipcMain.handle(
    'mail:items',
    (_e, sourceId: string, folderId: string, offset: number, limit: number) =>
      sourceManager.backend(sourceId).getItems(folderId, offset, limit),
  );

  ipcMain.handle(
    'mail:item',
    async (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
    const d = await sourceManager.backend(sourceId).getItem(folderId, index, stableId);
    // Convite já respondido → barra mostra o estado, não os botões.
    if (d?.invite) {
      loadRespondedInvites();
      let resp =
        (d.invite.uid && respondedInvites.get(d.invite.uid)) ||
        respondedInvites.get(`${sourceId}|${folderId}|${index}`);
      // Supressões abaixo valem só para o que exige resposta (REQUEST e
      // COUNTER — proposta de novo horário). REPLY/CANCEL mantêm o texto
      // informativo próprio do método.
      const needsResponse = d.invite.method === 'REQUEST' || d.invite.method === 'COUNTER';
      if (!resp && needsResponse) {
        const kind = mailFolderKind(sourceId, folderId);
        if (kind === 'sent' || kind === 'outbox' || kind === 'drafts') {
          // Cópia do convite que NÓS enviamos — não é um convite a responder.
          resp = 'sent';
        } else if (kind === 'calendar') {
          // Item dentro do Calendário = compromisso, não convite pendente —
          // como o Outlook, que não mostra Aceitar/Recusar na reunião.
          resp = 'accept';
        } else if (d.invite.uid && (await calendarUids()).has(d.invite.uid)) {
          // Convite aceito antes da persistência existir — o compromisso
          // com o mesmo UID já está no Calendário.
          resp = 'accept';
        }
      }
      if (resp) d.invite.responded = resp;
    }
    return d;
  });

  // ---------- Rascunhos ----------
  // "Salvar" do editor grava o .eml na pasta Rascunhos do arquivo de dados
  // principal (como o auto-save do Outlook). Duplo-clique no rascunho reabre
  // a janela de composição; ao enviar, o rascunho sai da pasta.
  ipcMain.handle('mail:saveDraft', async (_e, input: ComposeInput) => {
    const account =
      accountStore.getFull(input.accountId) || accountStore.list().find((a) => a.email);
    if (!account) throw new Error('Configure uma conta para salvar rascunhos');
    const raw = await composeRawMessage(account, { ...input, draft: true });
    // IMAP: o rascunho fica na pasta Rascunhos do servidor — aparece em
    // qualquer cliente e não cria arquivo de dados local.
    if (account.protocol === 'imap') {
      try {
        const be = sourceManager.backend(`acct:${input.accountId}`);
        if (be instanceof ImapMailSource) {
          const draftsPath = await be.specialFolderPath('drafts');
          if (draftsPath) {
            await be.appendItem(draftsPath, raw, true);
            return true;
          }
        }
      } catch {
        // servidor inacessível — cai no fallback local abaixo
      }
    }
    const store = loadSettings().defaultStorePath || ensureDeliveryStore(account);
    if (!store) throw new Error('Nenhum arquivo de dados disponível para rascunhos');
    let svc = sourceManager.serviceForPath(store);
    if (!svc) {
      const info = store.toLowerCase().endsWith('.pst')
        ? sourceManager.openPst(store)
        : sourceManager.openLocalStore(store);
      const b = sourceManager.backend(info.id);
      if (b instanceof PstService) svc = b;
    }
    if (svc) {
      svc.importNamedEml(resolveStoreFolderName(store, 'drafts'), raw, true);
    } else {
      const dir = kindDirFor(store, 'drafts');
      mkdirSync(dir, { recursive: true });
      writeFileSync(join(dir, `${Date.now()}-${randomUUID().slice(0, 8)}.eml`), raw);
    }
    return true;
  });

  // Reabre um rascunho para edição — devolve o payload do compose com os
  // anexos extraídos para arquivos temporários (reanexados no envio).
  ipcMain.handle(
    'mail:openDraft',
    async (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
      const svc = sourceManager.backend(sourceId);
      const d = await svc.getItem(folderId, index, stableId);
      if (!d) return null;
      // Limpa cópias temp de rascunhos abertos há mais de um dia
      const tmpRoot = join(tmpdir(), 'unitmail-drafts');
      try {
        for (const e of readdirSync(tmpRoot, { withFileTypes: true })) {
          const p = join(tmpRoot, e.name);
          if (Date.now() - statSync(p).mtimeMs > 24 * 3600 * 1000) rmSync(p, { recursive: true, force: true });
        }
      } catch {
        // raiz ainda não existe
      }
      const atts: string[] = [];
      const dir = join(tmpRoot, `${Date.now()}-${randomUUID().slice(0, 8)}`);
      for (const a of d.attachments || []) {
        try {
          const data = await svc.getAttachmentData(folderId, index, a.index, stableId);
          mkdirSync(dir, { recursive: true });
          const p = join(dir, safeFsName(a.filename || `anexo-${a.index}`));
          writeFileSync(p, Buffer.from(data.dataBase64, 'base64'));
          atts.push(p);
        } catch {
          // anexo ilegível — não impede reabrir o rascunho
        }
      }
      const hdrVal = (k: string) =>
        new RegExp(`^${k}:\\s*(.+)$`, 'im').exec(d.transportMessageHeaders || '')?.[1]?.trim();
      const bcc = hdrVal('bcc') || '';
      const acct = accountStore
        .list()
        .find(
          (a) => a.email && a.email.toLowerCase() === (d.senderEmailAddress || '').toLowerCase(),
        );
      const payload: import('../shared/types').ComposeDraftPayload = {
        accountId: acct?.id || '',
        to: d.displayTo || '',
        cc: d.displayCC || undefined,
        bcc: bcc || undefined,
        subject: d.subject || '',
        bodyHtml: d.bodyHTML || undefined,
        bodyText: d.body || '',
        attachmentPaths: atts.length ? atts : undefined,
        // Cabeçalhos de threading — rascunho de resposta volta a responder
        // na mesma thread ao enviar.
        inReplyTo: hdrVal('in-reply-to') || undefined,
        references: hdrVal('references') || undefined,
        draftRef: { sourceId, folderId, index, stableId: d.stableId },
      };
      return payload;
    },
  );

  ipcMain.handle(
    'mail:search',
    (e, sourceId: string, query: string, containerClass?: string, folderId?: string) =>
      // Live search — lotes por pasta chegam por evento enquanto a busca roda
      sourceManager.backend(sourceId).search(query, containerClass, undefined, (hits) => {
        if (!e.sender.isDestroyed()) e.sender.send('mail:searchBatch', sourceId, hits);
      }, folderId),
  );

  ipcMain.handle(
    'mail:attachmentData',
    (_e, sourceId: string, folderId: string, index: number, attachIndex: number, stableId?: string) =>
      sourceManager.backend(sourceId).getAttachmentData(folderId, index, attachIndex, stableId),
  );

  // ---------- Salvar como (.eml) e Imprimir ----------

  /** Exporta a mensagem aberta como arquivo .eml — funciona para itens do
   *  PST (recomponhe o MIME) e locais (copia o arquivo). */
  ipcMain.handle(
    'items:saveAs',
    async (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.exportItemEml !== 'function') {
        throw new Error('Salvar como só é suportado em arquivos de dados');
      }
      const detail = await svc.getItem(folderId, index, stableId);
      const safe = safeFsName(detail.subject || 'mensagem', 'mensagem').slice(0, 80);
      const res = await dialog.showSaveDialog({
        title: 'Salvar mensagem como',
        defaultPath: `${safe}.eml`,
        filters: [{ name: 'Mensagem de email', extensions: ['eml'] }],
      });
      if (res.canceled || !res.filePath) return null;
      const raw = await pst.exportItemEml(folderId, index, stableId);
      await writeFile(res.filePath, raw);
      return res.filePath;
    },
  );

  /** Imprime a mensagem aberta numa janela oculta do Electron. */
  ipcMain.handle(
    'items:print',
    async (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
      const svc = sourceManager.backend(sourceId);
      const detail = await svc.getItem(folderId, index, stableId);
      const esc = (s: string) =>
        s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const when = detail.deliveryTime || detail.sentTime;
      const html = `<!doctype html><html><head><meta charset="utf-8"><style>
        body { font-family: Calibri, 'Segoe UI', sans-serif; margin: 24px; color: #000; }
        .hdr { border-bottom: 1px solid #ccc; padding-bottom: 10px; margin-bottom: 14px; }
        .hdr h1 { font-size: 17px; margin: 0 0 8px; }
        .hdr div { font-size: 12px; margin: 2px 0; }
        .hdr b { display: inline-block; width: 90px; color: #555; }
        .body { font-size: 13px; }
        .body img { max-width: 100%; }
      </style></head><body>
        <div class="hdr">
          <h1>${esc(detail.subject || '(sem assunto)')}</h1>
          <div><b>De:</b> ${esc(detail.senderName || detail.senderEmailAddress)}</div>
          <div><b>Para:</b> ${esc(detail.displayTo || '')}</div>
          ${detail.displayCC ? `<div><b>Cc:</b> ${esc(detail.displayCC)}</div>` : ''}
          ${when ? `<div><b>Enviado em:</b> ${esc(new Date(when).toLocaleString('pt-BR'))}</div>` : ''}
        </div>
        <div class="body">${detail.bodyHTML || `<pre>${esc(detail.body || '')}</pre>`}</div>
      </body></html>`;
      const win = new BrowserWindow({ show: false, webPreferences: { sandbox: true } });
      try {
        await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
        await new Promise<void>((resolve, reject) => {
          win.webContents.print({ silent: false }, (ok, reason) =>
            ok ? resolve() : reject(new Error(reason || 'Impressão cancelada')),
          );
        });
        return true;
      } finally {
        win.destroy();
      }
    },
  );

  ipcMain.handle(
    'mail:saveAttachment',
    async (_e, sourceId: string, folderId: string, index: number, attachIndex: number, stableId?: string) => {
      const data = await sourceManager.backend(sourceId).getAttachmentData(folderId, index, attachIndex, stableId);
      if (!data.dataBase64) throw new Error('Anexo vazio ou não suportado');
      const res = await dialog.showSaveDialog({ title: 'Salvar anexo', defaultPath: data.filename });
      if (res.canceled || !res.filePath) return null;
      await writeFile(res.filePath, Buffer.from(data.dataBase64, 'base64'));
      return res.filePath;
    },
  );

  // Abrir anexo sem salvar: grava num diretório temporário e abre com o
  // aplicativo padrão do SO (como o "Abrir" do Outlook).
  ipcMain.handle(
    'mail:openAttachment',
    async (_e, sourceId: string, folderId: string, index: number, attachIndex: number, stableId?: string) => {
      const data = await sourceManager.backend(sourceId).getAttachmentData(folderId, index, attachIndex, stableId);
      if (!data.dataBase64) throw new Error('Anexo vazio ou não suportado');
      const dir = join(app.getPath('temp'), 'UnitMail-Attachments');
      mkdirSync(dir, { recursive: true });
      const safeName = safeFsName(data.filename || 'anexo', 'anexo');
      const path = join(dir, safeName);
      await writeFile(path, Buffer.from(data.dataBase64, 'base64'));
      const err = await shell.openPath(path); // '' = abriu; string = mensagem de erro
      if (err) throw new Error(err);
      return true;
    },
  );

  // "Salvar todos": escolhe uma pasta e grava todos os anexos da mensagem.
  ipcMain.handle(
    'mail:saveAllAttachments',
    async (_e, sourceId: string, folderId: string, index: number, stableId?: string) => {
      const svc = sourceManager.backend(sourceId);
      const detail = await svc.getItem(folderId, index, stableId);
      if (!detail.attachments.length) return null;
      const res = await dialog.showOpenDialog({
        title: 'Salvar todos os anexos em…',
        properties: ['openDirectory', 'createDirectory'],
      });
      if (res.canceled || !res.filePaths[0]) return null;
      const dir = res.filePaths[0];
      let saved = 0;
      for (const a of detail.attachments) {
        try {
          const data = await svc.getAttachmentData(folderId, index, a.index, stableId);
          if (!data.dataBase64) continue;
          const safeName = safeFsName(data.filename || `anexo-${a.index}`, `anexo-${a.index}`);
          await writeFile(join(dir, safeName), Buffer.from(data.dataBase64, 'base64'));
          saved++;
        } catch (e) {
          log.error(`anexo ${a.filename}`, e);
        }
      }
      return { dir, saved, total: detail.attachments.length };
    },
  );
}

/** Avisa o renderer que uma nova fonte entrou — a árvore recarrega e a
 *  base aparece assim que termina de abrir (restauração em 2º plano). */
function notifySourcesChanged(): void {
  try {
    mainWindow?.webContents.send('sources:changed');
  } catch {
    // janela ainda não criada — ignora
  }
}

async function restoreSession(): Promise<void> {
  // Regrava senhas antigas (safeStorage/DPAPI) no formato portátil antes de
  // qualquer uso — depois disso a pasta de dados restaura tudo em outro PC.
  migrateSecrets();
  // Fallback de restauração: uma base/PST cujo caminho original não existe
  // nesta máquina é procurado em <raiz>/Mail/<nome> — cobre backup restaurado
  // em outro PC onde a letra de drive/pasta é diferente. remapPstReferences
  // atualiza TODAS as referências (entrega de conta, regras, base principal…).
  {
    const s0 = loadSettings();
    // basename cross-platform: no Linux, '\' é caractere comum e
    // path.basename('D:\x\y') devolve a string INTEIRA — caminhos vindos de
    // um backup Windows nunca casavam e o remap não acontecia (as referências
    // continuavam "D:\..." e a entrega falhava silenciosamente).
    const anyBase = (p: string) => p.split(/[\\/]/).filter(Boolean).pop() || p;
    const seen = new Set<string>();
    // TODAS as referências persistidas a arquivos/bases — não só a lista de
    // abertos: a entrega de conta e a base principal também guardam caminhos
    // e eram as que quebravam o Enviar/Receber num backup restaurado.
    const foreignRe =
      process.platform === 'win32' ? /^\// : /^[a-zA-Z]:[\\/]/;
    // Contas cuja entrega apontava para caminho de outro SO tiveram o estado
    // de "já vistas" possivelmente envenenado: versões anteriores "entregavam"
    // num diretório fantasma ("D:\…" é nome válido no Linux) e marcavam a
    // mensagem como processada — ela nunca chegou na base e nunca seria
    // baixada de novo. Reseta o estado → o próximo Enviar/Receber re-baixa
    // TUDO que ainda estiver no servidor e a base converge com o outro PC
    // (a dedupe por pasta impede duplicar o que já está nela).
    const foreignTargets = new Set(
      ruleEngine
        .list()
        .filter((r) => r.targetPstPath && foreignRe.test(r.targetPstPath))
        .map((r) => r.accountId),
    );
    const poisonedAccounts = accountStore
      .list()
      .filter(
        (a) =>
          (a.deliveryPath && foreignRe.test(a.deliveryPath)) ||
          foreignTargets.has(a.id) ||
          foreignTargets.has('all'),
      )
      .map((a) => a.id);
    const candidates = new Set<string>([
      ...s0.openPstFiles,
      ...s0.openLocalStores,
      ...(s0.defaultStorePath ? [s0.defaultStorePath] : []),
      ...accountStore.list().map((a) => a.deliveryPath || '').filter(Boolean),
      ...ruleEngine.list().map((r) => r.targetPstPath || '').filter(Boolean),
    ]);
    for (const oldPath of candidates) {
      const key = resolve(oldPath).toLowerCase();
      if (seen.has(key) || existsSync(oldPath)) continue;
      seen.add(key);
      const alt = join(mailRoot(), anyBase(oldPath));
      if (existsSync(alt)) remapPstReferences(oldPath, alt);
    }
    // Entradas de registro com sintaxe de OUTRO SO que o remap não resolveu
    // ("D:\…" no Linux, "/home/…" no Windows): nunca podem ser um destino
    // válido — manter no localstores.json fazia isLocalStorePath devolver
    // true e a entrega por regra gravava em diretório fantasma.
    for (const s of listLocalStores()) {
      if (foreignRe.test(s.path) && !existsSync(s.path)) {
        unregisterLocalStore(s.path);
        log.info('restore', `base com caminho de outro SO removida do registro: ${s.path}`);
      }
    }
    for (const id of poisonedAccounts) {
      ruleEngine.resetSyncState(id);
      log.info(
        'restore',
        `estado de sincronização da conta ${id} reiniciado — entrega apontava para caminho de outro SO`,
      );
    }
  }
  // Re-sincronização total UMA vez por instalação: quem restaurou base com a
  // versão que entregava em diretório fantasma tem o mapa de "já vistas"
  // envenenado — os caminhos já foram corrigidos, então a detecção por
  // caminho estrangeiro acima não pega mais. Zera o estado de TODAS as
  // contas uma única vez: o próximo Enviar/Receber re-baixa tudo do
  // servidor e a base converge com a do outro PC (a dedupe por pasta
  // impede duplicar o que já está gravado).
  {
    const sFlag = loadSettings();
    if (!sFlag.fullResyncDone && accountStore.list().length > 0) {
      for (const acc of accountStore.list()) ruleEngine.resetSyncState(acc.id);
      sFlag.fullResyncDone = true;
      saveSettings(sFlag);
      log.info(
        'restore',
        're-sincronização total agendada — próximo Enviar/Receber re-baixa todas as mensagens do servidor',
      );
    }
  }
  const settings = loadSettings();
  for (const file of settings.openPstFiles) {
    try {
      if (existsSync(file)) {
        splashStatus(`Abrindo ${basename(file)}…`);
        sourceManager.openPst(file);
        notifySourcesChanged();
      }
    } catch {
      // arquivo removido/corrompido — ignora na restauração
    }
  }
  for (const storePath of settings.openLocalStores) {
    try {
      if (existsSync(storePath)) {
        splashStatus(`Abrindo ${basename(storePath)}…`);
        const def = registerLocalStore(storePath);
        sourceManager.openLocalStore(storePath, def?.name);
        notifySourcesChanged();
      }
    } catch {
      // base removida — ignora
    }
  }
  // Auto-descoberta: toda base dentro de <raiz>/Mail é reconhecida e aberta —
  // é o que faz um backup copiado para outro PC "voltar a funcionar" sem
  // recadastrar nada (modelo perfil do Thunderbird).
  splashStatus('Procurando arquivos de dados…');
  for (const found of scanMailDir()) {
    try {
      const def = registerLocalStore(found.path);
      if (!sourceManager.list().some((s) => s.filePath === found.path)) {
        sourceManager.openLocalStore(found.path, def?.name || found.name);
        addOpenLocalStore(found.path);
        notifySourcesChanged();
      }
    } catch {
      // base corrompida — ignora
    }
  }
  // Migração única: assinatura embutida da conta vira uma assinatura nomeada
  // na tela "Assinaturas" e a conta passa a apontar para ela.
  for (const acc of accountStore.list()) {
    if (acc.signature && !acc.signatureNewId && !acc.signatureReplyId) {
      const sig = signatureStore.upsert({ name: `Assinatura — ${acc.email}`, html: acc.signature });
      accountStore.patch(acc.id, {
        signatureNewId: sig.id,
        signatureReplyId: sig.id,
        signature: '',
      });
    }
  }
  // Bases locais registradas mas não abertas na sessão anterior não são
  // reabertas — só as que estavam abertas (openLocalStores).
  splashStatus('Conectando contas…');
  for (const acc of accountStore.list()) {
    sourceManager
      .openAccount(acc.id)
      .then(() => notifySourcesChanged())
      .catch(() => {
        // falha de conexão é reportada via info.error, não derruba o app
        notifySourcesChanged();
      });
  }
}

/**
 * Migração de userData: quando o app foi renomeado para "Unit Mail", o
 * diretório de dados do Electron passou de "%APPDATA%/outlook-linux" para
 * "%APPDATA%/Unit Mail". Copia arquivos/dados da pasta legada que ainda não
 * existem no novo local — assim o usuário não perde contas, regras, bases
 * locais nem mensagens arquivadas.
 */
function migrateLegacyUserData(): void {
  const current = app.getPath('userData');
  const candidates = [
    join(app.getPath('appData'), 'outlook-linux'),
    join(app.getPath('appData'), 'Outlook Linux'),
  ];
  const legacy = candidates.find((d) => d !== current && existsSync(d));
  if (!legacy) return;

  const { readdirSync, copyFileSync, cpSync, statSync } =
    require('fs') as typeof import('fs');
  mkdirSync(current, { recursive: true });

  // Arquivos JSON de configuração + caches de sync (não sobrescreve o novo)
  for (const f of readdirSync(legacy)) {
    const src = join(legacy, f);
    const dst = join(current, f);
    if (!statSync(src).isFile()) continue;
    if (!f.endsWith('.json') || f === 'settings.json' || existsSync(dst)) continue;
    try {
      copyFileSync(src, dst);
    } catch {
      // segue
    }
  }

  // settings.json: mescla — o legado preenche chaves ausentes no novo
  try {
    const lp = join(legacy, 'settings.json');
    const cp = join(current, 'settings.json');
    if (existsSync(lp)) {
      const legacySet = JSON.parse(readFileSync(lp, 'utf-8'));
      const curSet = existsSync(cp) ? JSON.parse(readFileSync(cp, 'utf-8')) : {};
      writeFileSync(cp, JSON.stringify({ ...legacySet, ...curSet }, null, 2));
    }
  } catch {
    // corrompido — ignora
  }

  // Diretórios de dados: archives (sombra .eml) e Data Stores (bases locais)
  for (const dir of ['archives', 'Data Stores', 'Local Storage']) {
    const src = join(legacy, dir);
    const dst = join(current, dir);
    if (!existsSync(src)) continue;
    try {
      if (!existsSync(dst)) {
        cpSync(src, dst, { recursive: true });
      } else {
        // copia só o que ainda não existe no novo local
        for (const sub of readdirSync(src)) {
          const s = join(src, sub);
          const d = join(dst, sub);
          if (!existsSync(d)) cpSync(s, d, { recursive: true });
        }
      }
    } catch {
      // diretório pode estar em uso — segue sem ele
    }
  }
}

// Instância única: se o app já estiver aberto, uma nova execução apenas
// traz a janela existente para frente em vez de abrir outra.
const gotSingleInstanceLock = app.requestSingleInstanceLock();
if (!gotSingleInstanceLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}

/**
 * Auto-instalação no menu do Linux: cria o .desktop + ícone em ~/.local ao
 * abrir o app. Funciona para qualquer pacote (tar.gz, AppImage, dir) —
 * o usuário roda o binário uma vez e o atalho aparece no menu.
 */
function installLinuxDesktopEntry(): void {
  if (process.platform !== 'linux') return;
  try {
    const home = app.getPath('home');
    const appsDir = join(home, '.local', 'share', 'applications');
    const iconsDir = join(home, '.local', 'share', 'icons', 'hicolor', '512x512', 'apps');
    mkdirSync(appsDir, { recursive: true });
    mkdirSync(iconsDir, { recursive: true });
    // Ícone: no pacote vem como extraResource; em dev pega do repo
    const iconSrc = existsSync(join(process.resourcesPath, 'icon.png'))
      ? join(process.resourcesPath, 'icon.png')
      : join(app.getAppPath(), 'build', 'icon.png');
    if (existsSync(iconSrc)) copyFileSync(iconSrc, join(iconsDir, 'unit-mail.png'));
    const desktop = [
      '[Desktop Entry]',
      'Type=Application',
      'Version=1.0',
      'Name=Unit Mail',
      'Comment=Cliente de email com suporte a arquivos PST',
      // Espaço no caminho do executável quebra o SUID sandbox do Chromium
      // (execvp corta no espaço) — ex.: "/opt/Unit Mail/unit-mail". Sem o
      // --no-sandbox o atalho gerado aqui abria uma vez e morria depois.
      `Exec="${process.execPath}"${/\s/.test(process.execPath) ? ' --no-sandbox' : ''} %U`,
      'Icon=unit-mail',
      'Terminal=false',
      'Categories=Network;Email;',
      'StartupWMClass=unit-mail',
      'MimeType=application/vnd.ms-outlook;',
      '',
    ].join('\n');
    writeFileSync(join(appsDir, 'unit-mail.desktop'), desktop);
  } catch {
    // sem permissão ou home indisponível — ignora (não impede o app de abrir)
  }
}

app.whenReady().then(async () => {
  if (!gotSingleInstanceLock) return;
  installLinuxDesktopEntry();
  migrateLegacyUserData();
  nativeTheme.themeSource = loadSettings().theme ?? 'system';
  registerIpc();
  // Splash primeiro — aparece instantaneamente enquanto a interface sobe
  createSplash();
  const win = createWindow();
  // A janela abre logo que o renderer pinta — NÃO espera as bases. PST
  // grande levava minutos na splash porque restoreSession() era await;
  // agora roda em segundo plano e cada base avisa o renderer ao ficar
  // pronta (evento 'sources:changed'), aparecendo uma a uma na árvore.
  splashStatus('Preparando a interface…');
  await Promise.race([mainWindowReady, new Promise((r) => setTimeout(r, 15000))]);
  if (!win.isDestroyed()) win.show();
  splashWin?.close();
  splashWin = null;
  // Ponteiro para pasta de dados que não existe AGORA (drive removível ainda
  // não montado, pasta renomeada): NÃO restaura bases nem roda sync — cada
  // gravação iria para o fallback "Documents/Unit Mail", um perfil fantasma
  // que ainda por cima parecia "perder" os dados do usuário. O renderer
  // mostra o aviso com "Tentar novamente"/"Escolher outra pasta".
  {
    const missing = profilePointerTarget();
    if (missing && !existsSync(missing)) {
      void mainWindowReady.then(() => {
        if (!win.isDestroyed()) win.webContents.send('profile:missing', missing);
      });
    } else {
      void restoreSession().finally(() => {
        // Enviar/Receber ao abrir (como o Outlook faz) — só depois que as
        // bases terminaram de restaurar, para não concorrer com a carga.
        setTimeout(() => {
          runSyncAndNotify();
        }, 5000);
      });
      scheduleAutoSync();
      startReminderTimer();
    }
  }
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      const w = createWindow();
      void mainWindowReady.then(() => {
        if (!w.isDestroyed()) w.show();
      });
    }
  });
});

app.on('window-all-closed', () => {
  sourceManager.closeAll();
  if (process.platform !== 'darwin') app.quit();
});
