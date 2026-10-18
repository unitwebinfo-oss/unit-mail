import { app, BrowserWindow, dialog, ipcMain, nativeTheme, Notification, shell } from 'electron';
import { basename, join, resolve } from 'path';
import { randomUUID, createHash } from 'crypto';
import { readFileSync, writeFileSync, existsSync, mkdirSync, copyFileSync, cpSync } from 'fs';
import { writeFile } from 'fs/promises';
import { sourceManager } from './sourceManager';
import { PstService } from './pstService';
import { accountStore, secureStorageAvailable, migrateSecrets } from './accountStore';
import { sendMail, testAccount, resolveSignatureHtml, composeRawMessage } from './smtpService';
import { avStatus } from './antivirus';
import { recordRecipients, searchContacts } from './contactHistory';
import { ruleEngine, ensureDeliveryStore, folderDirFor } from './ruleEngine';
import { signatureStore } from './signatureStore';
import { createLocalStore, deleteLocalStore, isLocalStorePath, listLocalStores, registerLocalStore, renameLocalStore, unregisterLocalStore } from './localStore';
import { dataDir, mailRoot, setDataDir, isUnitMailProfile, scanMailDir } from './dataDir';
import { log } from './logger';
import { migratePstToLocalStore } from './pstImport';
import { isWindowsImportAvailable, importOutlookAccountsFromRegistry } from './outlookImportWin';
import { scanThunderbirdAccounts, scanThunderbirdMail } from './thunderbirdImport';
import { importThunderbirdMail } from './thunderbirdMailImport';
import type { AccountDraft, ComposeInput, FolderNode, ImportedAccountGuess, MailRule, Signature } from '../shared/types';
import { DEFAULT_CATEGORIES } from '../shared/types';
import { safeFsName } from '../shared/fsName';

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
    if (infected > 0 && Notification.isSupported()) {
      new Notification({
        title: 'Ameaça detectada em mensagem recebida',
        body: `${infected} mensagem(ns) movida(s) para a pasta "Itens infectados".`,
      }).show();
    }
    if (archived > 0 && mainWindow && !mainWindow.isFocused() && Notification.isSupported()) {
      new Notification({
        title: 'Novas mensagens',
        body: `${archived} mensagem(ns) recebida(s) e arquivada(s).`,
      }).show();
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

/** Procura pastas de calendário na árvore de uma base. */
function calendarFolders(nodes: FolderNode[]): FolderNode[] {
  const out: FolderNode[] = [];
  const walk = (ns: FolderNode[]) => {
    for (const n of ns) {
      if (n.containerClass === 'IPF.Appointment' || /^(calend[aá]rio|calendar)$/i.test(n.name)) {
        out.push(n);
      }
      walk(n.children);
    }
  };
  walk(nodes);
  return out;
}

async function checkReminders(): Promise<void> {
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
        const key = `${src.id}|${cal.id}|${it.index}|${it.startTime}`;
        if (remindedKeys.has(key)) continue;
        // Janela: de (início - X min) até 1h depois do início
        if (now >= remindAt && now <= start + 60 * 60000) {
          remindedKeys.add(key);
          const body = `${new Date(it.startTime).toLocaleString('pt-BR')}${
            it.location ? ` — ${it.location}` : ''
          }`;
          try {
            if (Notification.isSupported()) {
              new Notification({ title: `Lembrete: ${it.subject}`, body }).show();
            }
          } catch {
            // notificação pode falhar em ambientes sem suporte
          }
          mainWindow?.webContents.send('calendar:reminder', {
            subject: it.subject,
            startTime: it.startTime,
            location: it.location || '',
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

function createWindow(): BrowserWindow {
  const iconPath = appIcon();
  // Barra de título integrada: os botões min/max/close do sistema ficam
  // sobre uma faixa azul renderizada pelo app (estilo Office moderno).
  // titleBarOverlay é suportado no Windows e no Linux; sem ele o Linux
  // mostrava a moldura nativa (barra preta do GNOME/KDE) em vez do azul.
  const winChrome =
    process.platform === 'win32' || process.platform === 'linux'
      ? {
          titleBarStyle: 'hidden' as const,
          titleBarOverlay: { color: '#0173c7', symbolColor: '#ffffff', height: 36 },
        }
      : {};

  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 600,
    show: false,
    title: 'Unit Mail',
    icon: iconPath,
    backgroundColor: '#f3f3f3',
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

  mainWindow = win;
  win.on('closed', () => {
    if (mainWindow === win) mainWindow = null;
  });

  mainWindowReady = new Promise((resolve) => {
    win.once('ready-to-show', resolve);
  });

  if (isDev) {
    win.loadURL(process.env.ELECTRON_RENDERER_URL!);
    win.webContents.openDevTools({ mode: 'detach' });
  } else {
    win.loadFile(join(__dirname, '../renderer/index.html'));
  }
  return win;
}

function registerIpc(): void {
  // ---------- Diálogos de arquivo ----------
  // ---------- Ajuda / Sobre ----------
  ipcMain.handle('app:info', () => ({
    name: 'Unit Mail',
    version: app.getVersion(),
    license: 'GPL-3.0-only',
    userData: dataDir(),
    documentsStoreDir: mailRoot(),
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
    (_e, srcId: string, srcFolderId: string, index: number, dstId: string, dstFolderId: string) =>
      sourceManager.moveItem(srcId, srcFolderId, index, dstId, dstFolderId),
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
    ) => sourceManager.moveItems(srcId, srcFolderId, indexes, dstId, dstFolderId),
  );

  ipcMain.handle('items:isLocal', (_e, sourceId: string, folderId: string, index: number) => {
    const svc = sourceManager.backend(sourceId);
    const pst = svc as import('./pstService').PstService;
    if (typeof pst.isLocalItem !== 'function') return false;
    return pst.isLocalItem(folderId, index);
  });

  ipcMain.handle('items:delete', (_e, sourceId: string, folderId: string, index: number) => {
    const svc = sourceManager.backend(sourceId);
    const pst = svc as import('./pstService').PstService;
    if (typeof pst.deleteItem !== 'function') {
      throw new Error('Excluir só é suportado em arquivos de dados (PST/base local)');
    }
    pst.deleteItem(folderId, index);
    return true;
  });

  ipcMain.handle(
    'items:deleteMany',
    (_e, sourceId: string, folderId: string, indexes: number[]) =>
      sourceManager.deleteItems(sourceId, folderId, indexes),
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
    (_e, sourceId: string, folderId: string, index: number) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.getItemCategories !== 'function') return [];
      return pst.getItemCategories(folderId, index);
    },
  );
  ipcMain.handle(
    'items:setCategories',
    (_e, sourceId: string, folderId: string, indexes: number[], cats: string[]) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.setItemCategories !== 'function') return false;
      pst.batch(() => {
        for (const ix of indexes) {
          try {
            pst.setItemCategories(folderId, ix, cats);
          } catch {
            // item sumiu — segue com os demais
          }
        }
      });
      return true;
    },
  );

  ipcMain.handle(
    'folders:move',
    (_e, sourceId: string, folderId: string, newParentFolderId: string | null) =>
      sourceManager.moveFolder(sourceId, folderId, newParentFolderId),
  );

  ipcMain.handle(
    'folders:create',
    (_e, sourceId: string, name: string, parentFolderId: string | null) =>
      sourceManager.createFolder(sourceId, name, parentFolderId),
  );

  // Novo compromisso / contato / tarefa (Novos Itens da Página Inicial)
  ipcMain.handle(
    'items:createSpecial',
    (_e, sourceId: string, item: import('../shared/types').NewSpecialItem) =>
      sourceManager.createSpecialItem(sourceId, item),
  );

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
  ipcMain.handle('folders:markAllRead', (_e, sourceId: string, folderId: string) => {
    sourceManager.markAllRead(sourceId, folderId);
    return true;
  });

  // "Executar Regras Agora" — aplica regras aos itens da pasta escolhida
  // (só itens locais são movidos; itens gravados no PST são copiados).
  // Com ruleIds, roda exatamente as regras marcadas — mesmo desativadas,
  // como no diálogo "Executar regras agora" do Outlook. Sem ruleIds,
  // usa as habilitadas (botão da barra de ações).
  ipcMain.handle(
    'rules:applyToFolder',
    async (_e, sourceId: string, folderId: string, ruleIds?: string[]) => {
    const { matchRule } = await import('./ruleEngine');
    const svc = sourceManager.backend(sourceId);
    const pst = svc as import('./pstService').PstService;
    if (typeof pst.isLocalItem !== 'function') return { applied: 0 };
    const all = ruleEngine.list();
    const rules = ruleIds?.length
      ? all.filter((r) => ruleIds.includes(r.id))
      : all.filter((r) => r.enabled);
    const sources = sourceManager.listAll();
    const items = await svc.getItems(folderId, 0, 5000);
    let applied = 0;
    // Processa do índice maior para o menor — remover itens não desloca os
    // índices dos que ainda serão avaliados.
    for (const it of [...items].reverse()) {
      for (const rule of rules) {
        if (
          !matchRule(rule, {
            from: `${it.senderEmailAddress} ${it.senderName}`,
            to: it.displayTo,
            subject: it.subject,
          })
        ) {
          continue;
        }
        const target = sources.find((s) => s.filePath === rule.targetPstPath);
        if (!target) break;
        try {
          await sourceManager.moveItem(
            sourceId,
            folderId,
            it.index,
            target.id,
            `arc:${rule.targetFolder}`,
          );
          applied++;
        } catch {
          // destino não gravável ou item já movido — segue
        }
        break; // moveu (ou tentou) — próxima mensagem
      }
    }
    return { applied };
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
    async (_e, sourceId: string, destParentDir: string, name: string, existingStorePath?: string) => {
      const info = await migratePstToLocalStore(
        sourceId,
        destParentDir || mailRoot(),
        name,
        (p) => {
          mainWindow?.webContents.send('migrate:progress', p);
        },
        existingStorePath,
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
      const message = await testAccount(
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
      return { ok: true, message };
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
    try {
      const store = ensureDeliveryStore(account);
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
        outboxFile = svc.importNamedEml('Caixa de Saída', raw, true);
      } else if (raw) {
        // Fallback sem serviço: grava direto (o seed do readstate marca
        // arquivos antigos como lidos na primeira abertura da base)
        const dir = folderDirFor(store, 'Caixa de Saída');
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
        // Sai da Caixa de Saída → Itens Enviados, como no Outlook
        if (svc && outboxFile) {
          svc.moveNamedFile('Caixa de Saída', 'Itens Enviados', outboxFile);
        } else if (raw) {
          try {
            const sentStore = ensureDeliveryStore(account);
            if (sentStore) {
              const dir = folderDirFor(sentStore, 'Itens Enviados');
              mkdirSync(dir, { recursive: true });
              writeFileSync(join(dir, `${Date.now()}-${randomUUID().slice(0, 8)}.eml`), raw);
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
        if (Notification.isSupported()) {
          new Notification({
            title: 'Falha ao enviar mensagem',
            body: `"${input.subject || '(sem assunto)'}" ficou na Caixa de Saída: ${msg}`,
          }).show();
        }
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
        accountId: string;
        response: 'accept' | 'decline' | 'counter';
        newStart?: string;
        newEnd?: string;
      },
    ) => {
      const svc = sourceManager.backend(input.sourceId);
      const detail = await svc.getItem(input.folderId, input.index);
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
            await pst.importItemEml('arc:Calendário', Buffer.from(eml, 'utf8'));
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
  ipcMain.handle('settings:getSyncInterval', () => loadSettings().syncIntervalMin ?? 30);
  ipcMain.handle('settings:getLoadImages', () => loadSettings().loadRemoteImages ?? false);
  ipcMain.handle('settings:getSetupComplete', () => loadSettings().setupComplete ?? false);
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
    (_e, sourceId: string, folderId: string, index: number, read: boolean) => {
      sourceManager.setItemRead(sourceId, folderId, index, read);
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
      settings?: { syncIntervalMin?: number; loadRemoteImages?: boolean; theme?: 'system' | 'light' | 'dark' };
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

  ipcMain.handle('mail:item', (_e, sourceId: string, folderId: string, index: number) =>
    sourceManager.backend(sourceId).getItem(folderId, index),
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
    (_e, sourceId: string, folderId: string, index: number, attachIndex: number) =>
      sourceManager.backend(sourceId).getAttachmentData(folderId, index, attachIndex),
  );

  // ---------- Salvar como (.eml) e Imprimir ----------

  /** Exporta a mensagem aberta como arquivo .eml — funciona para itens do
   *  PST (recomponhe o MIME) e locais (copia o arquivo). */
  ipcMain.handle(
    'items:saveAs',
    async (_e, sourceId: string, folderId: string, index: number) => {
      const svc = sourceManager.backend(sourceId);
      const pst = svc as import('./pstService').PstService;
      if (typeof pst.exportItemEml !== 'function') {
        throw new Error('Salvar como só é suportado em arquivos de dados');
      }
      const detail = await svc.getItem(folderId, index);
      const safe = safeFsName(detail.subject || 'mensagem', 'mensagem').slice(0, 80);
      const res = await dialog.showSaveDialog({
        title: 'Salvar mensagem como',
        defaultPath: `${safe}.eml`,
        filters: [{ name: 'Mensagem de email', extensions: ['eml'] }],
      });
      if (res.canceled || !res.filePath) return null;
      const raw = await pst.exportItemEml(folderId, index);
      await writeFile(res.filePath, raw);
      return res.filePath;
    },
  );

  /** Imprime a mensagem aberta numa janela oculta do Electron. */
  ipcMain.handle(
    'items:print',
    async (_e, sourceId: string, folderId: string, index: number) => {
      const svc = sourceManager.backend(sourceId);
      const detail = await svc.getItem(folderId, index);
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
    async (_e, sourceId: string, folderId: string, index: number, attachIndex: number) => {
      const data = await sourceManager.backend(sourceId).getAttachmentData(folderId, index, attachIndex);
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
    async (_e, sourceId: string, folderId: string, index: number, attachIndex: number) => {
      const data = await sourceManager.backend(sourceId).getAttachmentData(folderId, index, attachIndex);
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
    async (_e, sourceId: string, folderId: string, index: number) => {
      const svc = sourceManager.backend(sourceId);
      const detail = await svc.getItem(folderId, index);
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
          const data = await svc.getAttachmentData(folderId, index, a.index);
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
    const seen = new Set<string>();
    for (const oldPath of [...s0.openPstFiles, ...s0.openLocalStores]) {
      const key = resolve(oldPath).toLowerCase();
      if (seen.has(key) || existsSync(oldPath)) continue;
      seen.add(key);
      const alt = join(mailRoot(), basename(oldPath));
      if (existsSync(alt)) remapPstReferences(oldPath, alt);
    }
  }
  const settings = loadSettings();
  for (const file of settings.openPstFiles) {
    try {
      if (existsSync(file)) {
        splashStatus(`Abrindo ${basename(file)}…`);
        sourceManager.openPst(file);
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
    sourceManager.openAccount(acc.id).catch(() => {
      // falha de conexão é reportada via info.error, não derruba o app
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
      `Exec="${process.execPath}" %U`,
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
  // Splash primeiro — aparece instantaneamente enquanto as bases abrem
  createSplash();
  const win = createWindow();
  splashStatus('Restaurando arquivos de dados…');
  await restoreSession();
  // Troca splash → janela principal só quando o renderer está pronto
  splashStatus('Preparando a interface…');
  await Promise.race([mainWindowReady, new Promise((r) => setTimeout(r, 15000))]);
  if (!win.isDestroyed()) win.show();
  splashWin?.close();
  splashWin = null;
  scheduleAutoSync();
  startReminderTimer();
  // Enviar/Receber ao abrir (como o Outlook faz) — depois da UI subir
  setTimeout(() => {
    runSyncAndNotify();
  }, 10000);
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
