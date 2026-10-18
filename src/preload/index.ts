import { contextBridge, ipcRenderer } from 'electron';
import type {
  AccountConfig,
  AccountDraft,
  AccountTestResult,
  AttachmentData,
  ComposeInput,
  FolderNode,
  ImportedAccountGuess,
  ItemDetail,
  ItemSummary,
  MailCategory,
  MailRule,
  MailSource,
  NewSpecialItem,
  SearchResult,
  Signature,
  SyncProgress,
  SyncResult,
  TbMailProfile,
} from '../shared/types';

export interface OutlookApi {
  openPstDialog(): Promise<string | null>;

  // Bases (PST + contas)
  listSources(): Promise<MailSource[]>;
  openPst(filePath: string): Promise<MailSource>;
  closeSource(sourceId: string): Promise<boolean>;
  createLocalStore(name: string): Promise<MailSource | null>;
  openLocalStore(): Promise<MailSource | null>;
  getRecentFiles(): Promise<string[]>;
  getLastOpened(): Promise<string | null>;

  // Assinaturas
  listSignatures(): Promise<Signature[]>;
  upsertSignature(sig: Omit<Signature, 'id'> & { id?: string }): Promise<Signature>;
  removeSignature(id: string): Promise<boolean>;

  // Contas de email
  listAccounts(): Promise<AccountConfig[]>;
  upsertAccount(draft: AccountDraft): Promise<AccountConfig>;
  /** deleteStore=true apaga também a base local vinculada (nunca um .pst). */
  removeAccount(accountId: string, deleteStore?: boolean): Promise<boolean>;
  /** Apaga de verdade uma base de dados local do disco. */
  deleteLocalStore(storePath: string): Promise<boolean>;
  testAccount(draft: AccountDraft): Promise<AccountTestResult>;
  connectAccount(accountId: string): Promise<MailSource>;
  sendMail(input: ComposeInput): Promise<boolean>;
  pickAttachments(): Promise<string[]>;
  pickImage(): Promise<string | null>;
  signatureFor(accountId: string, kind?: 'new' | 'reply' | 'forward'): Promise<string>;
  renameLocalStore(path: string, newName: string): Promise<unknown>;
  createLocalStoreAuto(name: string): Promise<MailSource | null>;
  /** Cria base local num diretório escolhido pelo usuário. */
  createLocalStoreIn(name: string, parentDir: string): Promise<MailSource | null>;
  /** Abre o seletor para o usuário escolher a pasta-pai da nova base local. */
  pickStoreDir(): Promise<string | null>;
  /** Migra todo o conteúdo de um PST/base para uma base nativa do app (.eml).
   *  existingStorePath: importa dentro de uma base existente em vez de criar nova. */
  migratePst(sourceId: string, destParentDir: string, name: string, existingStorePath?: string): Promise<MailSource>;
  onMigrateProgress(
    cb: (p: { phase: 'folders' | 'items' | 'done'; currentFolder: string; done: number; total: number }) => void,
  ): () => void;
  getLoadImages(): Promise<boolean>;
  setLoadImages(v: boolean): Promise<boolean>;
  /** Verificação de vírus nas mensagens recebidas (antivírus local). */
  getVirusScan(): Promise<boolean>;
  setVirusScan(v: boolean): Promise<boolean>;
  /** Motor de antivírus local detectado (AMSI/Windows, ClamAV/Linux). */
  getAvStatus(): Promise<{ available: boolean; engine: string }>;
  /** Informações do app para a aba Ajuda (versão, licença, pastas de dados). */
  getAppInfo(): Promise<{
    name: string;
    version: string;
    license: string;
    userData: string;
    documentsStoreDir: string;
    electron: string;
    chrome: string;
    node: string;
    platform: string;
    /** false = senhas sem criptografia do SO (keyring indisponível) */
    secureStorage: boolean;
  }>;
  /** Abre a pasta de dados do app no explorador de arquivos. */
  openDataDir(): Promise<string>;
  /** Abre uma URL http(s) no navegador padrão do sistema. */
  openExternal(url: string): Promise<void>;
  /** Pasta raiz de dados (perfil portátil — contas, regras, bases Mail/). */
  getDataDir(): Promise<string>;
  /** Abre o seletor de pasta de dados — não grava nada (isProfile = já é perfil). */
  pickDataDirPath(): Promise<{ dir: string; isProfile: boolean } | null>;
  /** Aplica a pasta de dados escolhida (grava ponteiro, migra dados se nova). */
  applyDataDir(dir: string): Promise<{ dir: string; adopted: boolean } | null>;
  /** Reinicia o app (marca setup como concluído — usado ao adotar um perfil). */
  restartApp(): Promise<void>;
  /** Caminho do arquivo de dados principal ('' = nenhum). */
  getDefaultStore(): Promise<string>;
  setDefaultStore(path: string): Promise<string>;
  /** Tema da interface — 'system' segue o SO. */
  getTheme(): Promise<'system' | 'light' | 'dark'>;
  setTheme(t: 'system' | 'light' | 'dark'): Promise<string>;
  getSetupComplete(): Promise<boolean>;
  setSetupComplete(v: boolean): Promise<boolean>;
  moveItem(
    srcId: string,
    srcFolderId: string,
    index: number,
    dstId: string,
    dstFolderId: string,
  ): Promise<'movido' | 'copiado'>;
  isLocalItem(sourceId: string, folderId: string, index: number): Promise<boolean>;
  deleteItem(sourceId: string, folderId: string, index: number): Promise<boolean>;
  /** Em lote: excluir/mover centenas de itens numa chamada só. Retorna quantos. */
  deleteItems(sourceId: string, folderId: string, indexes: number[]): Promise<number>;
  moveItems(
    srcId: string,
    srcFolderId: string,
    indexes: number[],
    dstId: string,
    dstFolderId: string,
  ): Promise<number>;
  /** Marca item como lido/não lido (só itens locais — PST é somente leitura). */
  setItemRead(sourceId: string, folderId: string, index: number, read: boolean): Promise<boolean>;
  /** Lista-mestre de categorias coloridas (Categorizar). */
  getCategories(): Promise<MailCategory[]>;
  saveCategories(cats: MailCategory[]): Promise<boolean>;
  /** Categorias aplicadas a um item. */
  getItemCategories(sourceId: string, folderId: string, index: number): Promise<string[]>;
  /** Substitui as categorias de vários itens numa chamada (lote). */
  setItemCategories(sourceId: string, folderId: string, indexes: number[], cats: string[]): Promise<boolean>;
  moveFolder(sourceId: string, folderId: string, newParentFolderId: string | null): Promise<FolderNode[]>;
  createFolder(sourceId: string, name: string, parentFolderId: string | null): Promise<FolderNode[]>;
  /** Cria compromisso/contato/tarefa na base — retorna a árvore atualizada. */
  createSpecialItem(sourceId: string, item: NewSpecialItem): Promise<FolderNode[]>;
  renameFolder(sourceId: string, folderId: string, newName: string): Promise<FolderNode[]>;
  deleteFolder(sourceId: string, folderId: string): Promise<FolderNode[]>;
  /** Oculta/reexibe uma pasta na árvore — retorna a árvore atualizada. */
  setFolderHidden(sourceId: string, folderId: string, hidden: boolean): Promise<FolderNode[]>;
  copyFolder(sourceId: string, folderId: string): Promise<FolderNode[]>;
  emptyFolder(sourceId: string, folderId: string): Promise<number>;
  markAllRead(sourceId: string, folderId: string): Promise<boolean>;
  getFolderOrders(): Promise<Record<string, string[]>>;
  setFolderOrder(key: string, orderedIds: string[]): Promise<boolean>;
  /** Responde convite de reunião (Aceitar/Recusar/Propor novo horário). */
  respondInvite(input: {
    sourceId: string;
    folderId: string;
    index: number;
    accountId: string;
    response: 'accept' | 'decline' | 'counter';
    newStart?: string;
    newEnd?: string;
  }): Promise<{ sent: boolean }>;
  /** Evento do Enviar/Receber automático/manual do main — retorna unsubscribe. */
  onSyncDone(cb: (results: SyncResult[]) => void): () => void;
  /** Progresso por tarefa do Enviar/Receber — retorna unsubscribe. */
  onSyncProgress(cb: (ev: SyncProgress) => void): () => void;
  /** "Cancelar tudo" no diálogo de progresso. */
  syncCancel(): Promise<boolean>;
  /** Cancela uma tarefa específica do Enviar/Receber pelo taskId. */
  syncCancelTask(taskId: string): Promise<boolean>;
  /** Resultado do envio em segundo plano — modal fecha ao clicar em Enviar. */
  onSendResult(
    cb: (r: { ok: boolean; subject: string; error?: string }) => void,
  ): () => void;
  /** Link mailto: clicado num email — abre rascunho no app. Retorna unsubscribe. */
  onMailto(cb: (url: string) => void): () => void;
  /** Lembrete de compromisso disparado pelo main — retorna unsubscribe. */
  onCalendarReminder(
    cb: (e: { subject: string; startTime: string; location: string }) => void,
  ): () => void;
  syncAccount(accountId: string): Promise<SyncResult>;
  syncAll(): Promise<SyncResult[]>;
  /** Autocompletar de contatos — histórico de destinatários enviados. */
  suggestContacts(q: string): Promise<Array<{ name: string; email: string; count: number }>>;
  getSyncInterval(): Promise<number>;
  setSyncInterval(min: number): Promise<number>;
  listRules(): Promise<MailRule[]>;
  upsertRule(rule: Omit<MailRule, 'id'> & { id?: string }): Promise<MailRule>;
  removeRule(id: string): Promise<boolean>;
  /** "Executar Regras Agora" — aplica regras habilitadas aos itens da pasta. */
  applyRulesToFolder(sourceId: string, folderId: string, ruleIds?: string[]): Promise<{ applied: number }>;

  // Importação do Outlook (Windows)
  isOutlookImportAvailable(): Promise<boolean>;
  scanOutlookAccounts(): Promise<ImportedAccountGuess[]>;
  // Importação do Thunderbird (Linux, Windows e macOS)
  scanThunderbirdAccounts(): Promise<ImportedAccountGuess[]>;
  /** Perfis do Thunderbird com armazenamentos de email (.mbox) encontrados. */
  scanThunderbirdMail(): Promise<TbMailProfile[]>;
  /** Importa as mensagens de um perfil do Thunderbird para uma base nativa. */
  importThunderbird(profilePath: string, destParentDir: string, name: string, existingStorePath?: string): Promise<MailSource>;
  saveOutlookExport(guesses: ImportedAccountGuess[]): Promise<string | null>;
  openOutlookExportFile(): Promise<ImportedAccountGuess[] | null>;

  // Backup das configurações (contas + regras + assinaturas)
  exportBackup(): Promise<string | null>;
  importBackup(): Promise<{
    accounts: number;
    rules: number;
    signatures: number;
    warnings: string[];
  } | null>;

  // Operações de email (por sourceId)
  getFolderTree(sourceId: string): Promise<FolderNode[]>;
  getItemCount(sourceId: string, folderId: string): Promise<number>;
  getItems(sourceId: string, folderId: string, offset: number, limit: number): Promise<ItemSummary[]>;
  getItem(sourceId: string, folderId: string, index: number): Promise<ItemDetail>;
  search(
    sourceId: string,
    query: string,
    containerClass?: string,
    folderId?: string,
  ): Promise<SearchResult[]>;
  /** Lotes de resultados durante a busca (live search) — retorna unsubscribe. */
  onSearchBatch(cb: (sourceId: string, hits: SearchResult[]) => void): () => void;
  saveAttachment(
    sourceId: string,
    folderId: string,
    index: number,
    attachIndex: number,
  ): Promise<string | null>;
  /** Abre o anexo com o aplicativo padrão sem salvá-lo (cópia temporária). */
  openAttachment(
    sourceId: string,
    folderId: string,
    index: number,
    attachIndex: number,
  ): Promise<boolean>;
  /** Escolhe uma pasta e grava todos os anexos da mensagem. */
  saveAllAttachments(
    sourceId: string,
    folderId: string,
    index: number,
  ): Promise<{ dir: string; saved: number; total: number } | null>;
  getAttachmentData(
    sourceId: string,
    folderId: string,
    index: number,
    attachIndex: number,
  ): Promise<AttachmentData>;
  /** Salva a mensagem como .eml (diálogo). Retorna o caminho salvo ou null. */
  saveItemAs(sourceId: string, folderId: string, index: number): Promise<string | null>;
  /** Imprime a mensagem (diálogo de impressão do SO). */
  printItem(sourceId: string, folderId: string, index: number): Promise<boolean>;
}

const api: OutlookApi = {
  openPstDialog: () => ipcRenderer.invoke('dialog:openPst'),

  listSources: () => ipcRenderer.invoke('sources:list'),
  openPst: (filePath) => ipcRenderer.invoke('sources:openPst', filePath),
  closeSource: (sourceId) => ipcRenderer.invoke('sources:closeSource', sourceId),
  createLocalStore: (name) => ipcRenderer.invoke('stores:createLocal', name),
  openLocalStore: () => ipcRenderer.invoke('stores:openLocal'),
  getRecentFiles: () => ipcRenderer.invoke('app:getRecentFiles'),
  getLastOpened: () => ipcRenderer.invoke('app:getLastOpened'),

  listSignatures: () => ipcRenderer.invoke('signatures:list'),
  upsertSignature: (sig) => ipcRenderer.invoke('signatures:upsert', sig),
  removeSignature: (id) => ipcRenderer.invoke('signatures:remove', id),

  listAccounts: () => ipcRenderer.invoke('accounts:list'),
  upsertAccount: (draft) => ipcRenderer.invoke('accounts:upsert', draft),
  removeAccount: (accountId, deleteStore) => ipcRenderer.invoke('accounts:remove', accountId, deleteStore),
  deleteLocalStore: (storePath) => ipcRenderer.invoke('stores:deleteLocal', storePath),
  testAccount: (draft) => ipcRenderer.invoke('accounts:test', draft),
  connectAccount: (accountId) => ipcRenderer.invoke('accounts:connect', accountId),
  sendMail: (input) => ipcRenderer.invoke('accounts:send', input),
  pickAttachments: () => ipcRenderer.invoke('dialog:pickAttachments'),
  pickImage: () => ipcRenderer.invoke('dialog:pickImage'),
  signatureFor: (accountId, kind) => ipcRenderer.invoke('accounts:signatureFor', accountId, kind),
  renameLocalStore: (path, newName) => ipcRenderer.invoke('stores:renameLocal', path, newName),
  createLocalStoreAuto: (name) => ipcRenderer.invoke('stores:createLocalAuto', name),
  createLocalStoreIn: (name, parentDir) =>
    ipcRenderer.invoke('stores:createLocalIn', name, parentDir),
  pickStoreDir: () => ipcRenderer.invoke('stores:pickStoreDir'),
  migratePst: (sourceId, destParentDir, name, existingStorePath) =>
    ipcRenderer.invoke('stores:migratePst', sourceId, destParentDir, name, existingStorePath),
  onMigrateProgress: (cb) => {
    const listener = (
      _e: unknown,
      p: { phase: 'folders' | 'items' | 'done'; currentFolder: string; done: number; total: number },
    ) => cb(p);
    ipcRenderer.on('migrate:progress', listener);
    return () => ipcRenderer.removeListener('migrate:progress', listener);
  },
  getLoadImages: () => ipcRenderer.invoke('settings:getLoadImages'),
  setLoadImages: (v) => ipcRenderer.invoke('settings:setLoadImages', v),
  getVirusScan: () => ipcRenderer.invoke('settings:getVirusScan'),
  setVirusScan: (v) => ipcRenderer.invoke('settings:setVirusScan', v),
  getAvStatus: () => ipcRenderer.invoke('antivirus:status'),
  getDefaultStore: () => ipcRenderer.invoke('settings:getDefaultStore'),
  setDefaultStore: (p) => ipcRenderer.invoke('settings:setDefaultStore', p),
  getTheme: () => ipcRenderer.invoke('settings:getTheme'),
  setTheme: (t) => ipcRenderer.invoke('settings:setTheme', t),
  getSetupComplete: () => ipcRenderer.invoke('settings:getSetupComplete'),
  setSetupComplete: (v) => ipcRenderer.invoke('settings:setSetupComplete', v),
  moveItem: (srcId, srcFolderId, index, dstId, dstFolderId) =>
    ipcRenderer.invoke('items:move', srcId, srcFolderId, index, dstId, dstFolderId),
  isLocalItem: (sourceId, folderId, index) =>
    ipcRenderer.invoke('items:isLocal', sourceId, folderId, index),
  deleteItem: (sourceId, folderId, index) =>
    ipcRenderer.invoke('items:delete', sourceId, folderId, index),
  deleteItems: (sourceId, folderId, indexes) =>
    ipcRenderer.invoke('items:deleteMany', sourceId, folderId, indexes),
  moveItems: (srcId, srcFolderId, indexes, dstId, dstFolderId) =>
    ipcRenderer.invoke('items:moveMany', srcId, srcFolderId, indexes, dstId, dstFolderId),
  setItemRead: (sourceId, folderId, index, read) =>
    ipcRenderer.invoke('items:setRead', sourceId, folderId, index, read),
  getCategories: () => ipcRenderer.invoke('cats:list'),
  saveCategories: (cats) => ipcRenderer.invoke('cats:save', cats),
  getItemCategories: (sourceId, folderId, index) =>
    ipcRenderer.invoke('items:getCategories', sourceId, folderId, index),
  setItemCategories: (sourceId, folderId, indexes, cats) =>
    ipcRenderer.invoke('items:setCategories', sourceId, folderId, indexes, cats),
  moveFolder: (sourceId, folderId, newParentFolderId) =>
    ipcRenderer.invoke('folders:move', sourceId, folderId, newParentFolderId),
  createFolder: (sourceId, name, parentFolderId) =>
    ipcRenderer.invoke('folders:create', sourceId, name, parentFolderId),
  createSpecialItem: (sourceId, item) =>
    ipcRenderer.invoke('items:createSpecial', sourceId, item),
  renameFolder: (sourceId, folderId, newName) =>
    ipcRenderer.invoke('folders:rename', sourceId, folderId, newName),
  deleteFolder: (sourceId, folderId) => ipcRenderer.invoke('folders:delete', sourceId, folderId),
  setFolderHidden: (sourceId, folderId, hidden) =>
    ipcRenderer.invoke('folders:setHidden', sourceId, folderId, hidden),
  copyFolder: (sourceId, folderId) => ipcRenderer.invoke('folders:copy', sourceId, folderId),
  emptyFolder: (sourceId, folderId) => ipcRenderer.invoke('folders:empty', sourceId, folderId),
  markAllRead: (sourceId, folderId) =>
    ipcRenderer.invoke('folders:markAllRead', sourceId, folderId),
  getFolderOrders: () => ipcRenderer.invoke('settings:getFolderOrders'),
  setFolderOrder: (key, orderedIds) =>
    ipcRenderer.invoke('settings:setFolderOrder', key, orderedIds),
  respondInvite: (input) => ipcRenderer.invoke('mail:respondInvite', input),
  onSyncDone: (cb) => {
    const listener = (_e: unknown, results: SyncResult[]) => cb(results);
    ipcRenderer.on('sync:done', listener);
    return () => ipcRenderer.removeListener('sync:done', listener);
  },
  onSyncProgress: (cb) => {
    const listener = (_e: unknown, ev: SyncProgress) => cb(ev);
    ipcRenderer.on('sync:progress', listener);
    return () => ipcRenderer.removeListener('sync:progress', listener);
  },
  onSendResult: (cb) => {
    const listener = (
      _e: unknown,
      r: { ok: boolean; subject: string; error?: string },
    ) => cb(r);
    ipcRenderer.on('mail:sendResult', listener);
    return () => ipcRenderer.removeListener('mail:sendResult', listener);
  },
  onMailto: (cb) => {
    const listener = (_e: unknown, url: string) => cb(url);
    ipcRenderer.on('mail:openMailto', listener);
    return () => ipcRenderer.removeListener('mail:openMailto', listener);
  },
  onCalendarReminder: (cb) => {
    const listener = (
      _e: unknown,
      ev: { subject: string; startTime: string; location: string },
    ) => cb(ev);
    ipcRenderer.on('calendar:reminder', listener);
    return () => ipcRenderer.removeListener('calendar:reminder', listener);
  },
  syncAccount: (accountId) => ipcRenderer.invoke('accounts:sync', accountId),
  syncAll: () => ipcRenderer.invoke('accounts:syncAll'),
  suggestContacts: (q) => ipcRenderer.invoke('contacts:suggest', q),
  syncCancel: () => ipcRenderer.invoke('accounts:syncCancel'),
  syncCancelTask: (taskId) => ipcRenderer.invoke('accounts:syncCancelTask', taskId),
  getSyncInterval: () => ipcRenderer.invoke('settings:getSyncInterval'),
  setSyncInterval: (min) => ipcRenderer.invoke('settings:setSyncInterval', min),
  listRules: () => ipcRenderer.invoke('rules:list'),
  upsertRule: (rule) => ipcRenderer.invoke('rules:upsert', rule),
  removeRule: (id) => ipcRenderer.invoke('rules:remove', id),
  applyRulesToFolder: (sourceId, folderId, ruleIds) =>
    ipcRenderer.invoke('rules:applyToFolder', sourceId, folderId, ruleIds),

  isOutlookImportAvailable: () => ipcRenderer.invoke('outlookImport:available'),
  scanOutlookAccounts: () => ipcRenderer.invoke('outlookImport:scan'),
  scanThunderbirdAccounts: () => ipcRenderer.invoke('thunderbirdImport:scan'),
  scanThunderbirdMail: () => ipcRenderer.invoke('thunderbirdImport:scanMail'),
  importThunderbird: (profilePath, destParentDir, name, existingStorePath) =>
    ipcRenderer.invoke('stores:importThunderbird', profilePath, destParentDir, name, existingStorePath),
  saveOutlookExport: (guesses) => ipcRenderer.invoke('outlookImport:saveExport', guesses),
  openOutlookExportFile: () => ipcRenderer.invoke('outlookImport:openExportFile'),
  exportBackup: () => ipcRenderer.invoke('backup:export'),
  importBackup: () => ipcRenderer.invoke('backup:import'),
  getAppInfo: () => ipcRenderer.invoke('app:info'),
  openDataDir: () => ipcRenderer.invoke('app:openDataDir'),
  openExternal: (url) => ipcRenderer.invoke('app:openExternal', url),
  getDataDir: () => ipcRenderer.invoke('app:getDataDir'),
  pickDataDirPath: () => ipcRenderer.invoke('app:pickDataDirPath'),
  applyDataDir: (dir) => ipcRenderer.invoke('app:applyDataDir', dir),
  restartApp: () => ipcRenderer.invoke('app:restart'),

  getFolderTree: (sourceId) => ipcRenderer.invoke('mail:folderTree', sourceId),
  getItemCount: (sourceId, folderId) => ipcRenderer.invoke('mail:itemCount', sourceId, folderId),
  getItems: (sourceId, folderId, offset, limit) =>
    ipcRenderer.invoke('mail:items', sourceId, folderId, offset, limit),
  getItem: (sourceId, folderId, index) => ipcRenderer.invoke('mail:item', sourceId, folderId, index),
  search: (sourceId, query, containerClass, folderId) =>
    ipcRenderer.invoke('mail:search', sourceId, query, containerClass, folderId),
  onSearchBatch: (cb) => {
    const listener = (_e: unknown, sourceId: string, hits: SearchResult[]) =>
      cb(sourceId, hits);
    ipcRenderer.on('mail:searchBatch', listener);
    return () => ipcRenderer.removeListener('mail:searchBatch', listener);
  },
  saveAttachment: (sourceId, folderId, index, attachIndex) =>
    ipcRenderer.invoke('mail:saveAttachment', sourceId, folderId, index, attachIndex),
  openAttachment: (sourceId, folderId, index, attachIndex) =>
    ipcRenderer.invoke('mail:openAttachment', sourceId, folderId, index, attachIndex),
  saveAllAttachments: (sourceId, folderId, index) =>
    ipcRenderer.invoke('mail:saveAllAttachments', sourceId, folderId, index),
  saveItemAs: (sourceId, folderId, index) =>
    ipcRenderer.invoke('items:saveAs', sourceId, folderId, index),
  printItem: (sourceId, folderId, index) =>
    ipcRenderer.invoke('items:print', sourceId, folderId, index),
  getAttachmentData: (sourceId, folderId, index, attachIndex) =>
    ipcRenderer.invoke('mail:attachmentData', sourceId, folderId, index, attachIndex),
};

contextBridge.exposeInMainWorld('outlook', api);
