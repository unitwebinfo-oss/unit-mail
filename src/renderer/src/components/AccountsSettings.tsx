import React, { useEffect, useRef, useState } from 'react';
import { Icon } from '../icons';
import { ConfirmDialog } from './ConfirmDialog';
import { RichTextEditor } from './RichTextEditor';
import { guessServers, applyTheme } from '../utils';
import type {
  AccountConfig,
  AccountDraft,
  FolderNode,
  ImportedAccountGuess,
  MailRule,
  MailSource,
  Signature,
  TbMailProfile,
} from '../../../shared/types';

interface AccountsSettingsProps {
  sources: MailSource[];
  onClose: () => void;
  onCloseSource: (sourceId: string) => void;
  onAccountsChanged: () => void;
  /** Aba inicial ao abrir (ex.: 'data' vindo do botão "Abrir PST") */
  initialTab?: SettingsTab;
  /** Conta para abrir já em edição (Arquivo → Configurações de Conta). */
  editAccountId?: string;
  /** Ação para disparar ao abrir (Arquivo → Importar .pst / Thunderbird). */
  initialAction?: 'importPst' | 'importTb';
}

const EMPTY_DRAFT: AccountDraft = {
  label: '',
  email: '',
  protocol: 'pop3',
  imapHost: '',
  imapPort: 995,
  imapSecure: true,
  imapUser: '',
  smtpHost: '',
  smtpPort: 587,
  smtpSecure: false,
  smtpUser: '',
  signature: '',
  leaveOnServer: true,
  leaveDays: 0,
  password: '',
};

export type SettingsTab = 'email' | 'data' | 'rules' | 'sigs' | 'geral';

function Field({
  label,
  children,
}: {
  label: string;
  children: React.ReactNode;
}) {
  return (
    <label className="fld">
      <span>{label}</span>
      {children}
    </label>
  );
}

/** Regra em edição — pode ser nova (sem id) ou existente. */
type RuleDraft = Omit<MailRule, 'id'> & { id?: string };

export function AccountsSettings({
  sources,
  onClose,
  onCloseSource,
  onAccountsChanged,
  initialTab,
  editAccountId,
  initialAction,
}: AccountsSettingsProps) {
  const [accounts, setAccounts] = useState<AccountConfig[]>([]);
  const [draft, setDraft] = useState<AccountDraft>(EMPTY_DRAFT);
  const [editing, setEditing] = useState(false);
  const [acctTab, setAcctTab] = useState<'conta' | 'servidores'>('conta');
  const [secureStore, setSecureStore] = useState(true);
  useEffect(() => {
    window.outlook.getAppInfo?.().then((i) => setSecureStore(i.secureStorage)).catch(() => {});
  }, []);
  const [testMsg, setTestMsg] = useState('');
  const [busy, setBusy] = useState(false);
  const [importAvailable, setImportAvailable] = useState(false);
  const [guesses, setGuesses] = useState<ImportedAccountGuess[] | null>(null);
  const [scanMsg, setScanMsg] = useState('');
  /** Assistente de importação: etapa "choose" (escolhe a conta) ou "form" (configura+testa). */
  const [wizStep, setWizStep] = useState<'choose' | 'form'>('choose');
  /** Email da conta marcada na etapa de escolha. */
  const [wizPick, setWizPick] = useState('');
  /** Nota exibida dentro do assistente (resultado da busca, aviso de senhas…). */
  const [wizNote, setWizNote] = useState('');
  /** Busca de contas em andamento — o assistente já abre mostrando loading. */
  const [wizScanning, setWizScanning] = useState(false);
  const [impTest, setImpTest] = useState('');
  /** Resultado da busca → abre o assistente; vazio só registra a mensagem. */
  const loadGuesses = (list: ImportedAccountGuess[]) => {
    setGuesses(list.length ? list : null);
    setWizStep('choose');
    setWizPick('');
    setImp(null);
  };
  const [rules, setRules] = useState<MailRule[]>([]);
  // "Executar regras agora" — escolhe a conta, as regras e a pasta a varrer
  const [runFolder, setRunFolder] = useState('');
  const [runFolders, setRunFolders] = useState<Array<{ id: string; label: string }>>([]);
  const [runAccount, setRunAccount] = useState('');
  const [runRuleIds, setRunRuleIds] = useState<string[]>([]);
  const [runMsg, setRunMsg] = useState('');
  const [runBusy, setRunBusy] = useState(false);
  const [showRun, setShowRun] = useState(false);
  /** Regras de uma conta (inclui as "todas as contas") — uma conta por vez, para não pesar. */
  const runAccountRules = rules.filter((r) => r.accountId === runAccount || r.accountId === 'all');
  useEffect(() => {
    if (!runAccount && accounts.length) setRunAccount(accounts[0].id);
  }, [runAccount, accounts]);
  // Ao trocar de conta, vêm marcadas as regras habilitadas dela; o usuário pode
  // marcar outras (inclusive desativadas) — execução manual, como no Outlook.
  useEffect(() => {
    setRunRuleIds(rules.filter((r) => r.enabled && (r.accountId === runAccount || r.accountId === 'all')).map((r) => r.id));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [runAccount, rules]);
  /** Assinaturas do Outlook encontradas na importação — o usuário decide criar ou não. */
  const [importedSigs, setImportedSigs] = useState<Array<{ name: string; html: string }>>([]);
  // Importação conta a conta — modal com as opções antes de criar a conta.
  const [imp, setImp] = useState<{
    g: ImportedAccountGuess;
    label: string;
    password: string;
    path: string;
    newData: boolean;
    newName: string;
    leave: boolean;
    leaveDays: number;
  } | null>(null);
  const [impBusy, setImpBusy] = useState(false);
  const [impErr, setImpErr] = useState('');
  /** Emails já importados nesta sessão da lista — viram "importada" e saem da seleção. */
  const [doneEmails, setDoneEmails] = useState<Set<string>>(new Set());
  /** Conta em edição: cria um arquivo de dados novo (true) ou usa um existente (false). */
  const [acctNewData, setAcctNewData] = useState(false);
  const [acctNewDataName, setAcctNewDataName] = useState('');
  const [defaultStore, setDefaultStore] = useState('');
  const runSources = sources.filter((s) => s.kind === 'local' || s.kind === 'pst');
  // A conta define o arquivo de dados: o da entrega dela, senão o principal.
  const runAccountObj = accounts.find((a) => a.id === runAccount);
  const runDeliverySrc = runAccountObj?.deliveryPath
    ? runSources.find((s) => s.filePath === runAccountObj.deliveryPath)
    : undefined;
  const runSrcObj = runDeliverySrc || runSources.find((s) => s.filePath === defaultStore) || runSources[0];
  const runSource = runSrcObj?.id || '';
  // Só carrega a árvore quando o diálogo "Executar regras agora" está aberto.
  useEffect(() => {
    if (!showRun || !runSource) return;
    window.outlook
      .getFolderTree(runSource)
      .then((tree: FolderNode[]) => {
        const flat: Array<{ id: string; label: string }> = [];
        const walk = (ns: FolderNode[], depth: number) => {
          for (const n of ns) {
            flat.push({ id: n.id, label: `${'\u00a0\u00a0\u00a0'.repeat(depth)}${n.name}` });
            walk(n.children, depth + 1);
          }
        };
        walk(tree, 0);
        setRunFolders(flat);
        // Caixa de Entrada é a pasta padrão (Inbox em bases importadas do Outlook)
        const inbox = flat.find((f) => /caixa de entrada|inbox/i.test(f.label.trim()));
        setRunFolder((inbox ?? flat[0])?.id ?? '');
      })
      .catch(() => setRunFolders([]));
  }, [runSource, showRun]);
  const runNow = async () => {
    if (!runSource || !runFolder) return;
    if (runRuleIds.length === 0) {
      setRunMsg('Marque ao menos uma regra para executar.');
      return;
    }
    setRunBusy(true);
    setRunMsg('');
    try {
      const r = await window.outlook.applyRulesToFolder(runSource, runFolder, runRuleIds);
      setRunMsg(
        r.applied > 0
          ? `${r.applied} mensagem(ns) combinaram com as regras e foram movidas.`
          : 'Nenhuma mensagem da pasta combinou com as regras.',
      );
    } catch (e) {
      setRunMsg(`Erro ao executar: ${(e as Error).message}`);
    } finally {
      setRunBusy(false);
    }
  };
  const [ruleDraft, setRuleDraft] = useState<RuleDraft | null>(null);
  const [customFolder, setCustomFolder] = useState(false);
  const [signatures, setSignatures] = useState<Signature[]>([]);
  const [sigDraft, setSigDraft] = useState<(Omit<Signature, 'id'> & { id?: string }) | null>(null);
  /** Contas que devem usar a assinatura em edição — aplicadas ao salvar. */
  const [sigAssign, setSigAssign] = useState<Record<string, { new: boolean; reply: boolean }>>({});
  const [newStoreName, setNewStoreName] = useState('');
  const [newStoreAccount, setNewStoreAccount] = useState('');
  const [showNewStore, setShowNewStore] = useState(false);
  const [tab, setTab] = useState<SettingsTab>(initialTab || 'email');
  const [folderChoices, setFolderChoices] = useState<Array<{ path: string; label: string }>>([]);
  const [renaming, setRenaming] = useState<{ path: string; name: string } | null>(null);
  const [loadImages, setLoadImages] = useState(false);
  const [syncMin, setSyncMin] = useState(30);
  const [virusScan, setVirusScan] = useState(true);
  const [avEngine, setAvEngine] = useState('');
  const [theme, setThemeState] = useState<'system' | 'light' | 'dark'>('system');
  const [geralSaved, setGeralSaved] = useState(false);
  const [dataDirPath, setDataDirPath] = useState('');
  const [confirm, setConfirm] = useState<{
    title: string;
    message: string;
    onOk: () => void;
  } | null>(null);

  // ---------- Migrar PST → base própria do app ----------
  const [migrateSrc, setMigrateSrc] = useState<MailSource | null>(null);
  const [migrateName, setMigrateName] = useState('');
  const [migrateDir, setMigrateDir] = useState('');
  const [migrateProgress, setMigrateProgress] = useState<{
    phase: 'folders' | 'items' | 'done';
    currentFolder: string;
    done: number;
    total: number;
  } | null>(null);
  const [migrateError, setMigrateError] = useState('');
  const [migrateDone, setMigrateDone] = useState<string | null>(null);
  // '' = criar base nova; senão = caminho da base local existente de destino
  const [migrateTarget, setMigrateTarget] = useState('');
  // true quando o PST foi aberto só para esta importação (botão "Importar
  // arquivo .pst…") — ao concluir, fecha-o de novo para não deixar um PST
  // "pendurado" na lista, já que o objetivo é só a base convertida.
  const [migrateCloseAfter, setMigrateCloseAfter] = useState(false);

  // ---------- Importar Thunderbird (.mbox do perfil) ----------
  const [tbProfiles, setTbProfiles] = useState<TbMailProfile[] | null>(null);
  const [tbSrc, setTbSrc] = useState('');
  const [tbName, setTbName] = useState('');
  const [tbDir, setTbDir] = useState('');
  const [tbTarget, setTbTarget] = useState('');
  const [tbError, setTbError] = useState('');
  const [tbBusy, setTbBusy] = useState(false);
  const [tbDone, setTbDone] = useState<string | null>(null);

  useEffect(() => {
    if (migrateSrc) {
      setMigrateName(migrateCloseAfter ? migrateSrc.label : `${migrateSrc.label} (migrado)`);
      setMigrateDir('');
      setMigrateTarget('');
      setMigrateProgress(null);
      setMigrateError('');
      setMigrateDone(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [migrateSrc]);

  useEffect(() => {
    return window.outlook.onMigrateProgress?.(setMigrateProgress);
  }, []);

  /** Botão "Importar arquivo .pst…" — abre o arquivo e já pede os dados da migração. */
  const importPst = async () => {
    const path = await window.outlook.openPstDialog();
    if (!path) return;
    try {
      const info = await window.outlook.openPst(path);
      setMigrateCloseAfter(true);
      setMigrateSrc(info);
    } catch (e) {
      setScanMsg(`Erro ao abrir .pst: ${(e as Error).message}`);
    }
  };

  const runMigration = async () => {
    if (!migrateSrc) return;
    if (!migrateTarget && !migrateName.trim()) return;
    setMigrateError('');
    try {
      // migrateDir vazio → o main cria na pasta de dados padrão (Mail/).
      const dest = await window.outlook.migratePst(
        migrateSrc.id,
        migrateDir,
        migrateName.trim(),
        migrateTarget || undefined,
      );
      // Fluxo de importação direta: fecha o PST temporário — só a base
      // convertida fica visível em Arquivos de Dados. O arquivo .pst em si
      // nunca é apagado, só deixa de aparecer como fonte aberta no app.
      if (migrateCloseAfter) await onCloseSource(migrateSrc.id);
      onAccountsChanged();
      // Concluiu: fecha TODOS os modais (este + Configurações) — a base nova
      // já aparece na árvore de pastas e em Arquivos de Dados ao fundo.
      setMigrateSrc(null);
      onClose();
    } catch (e) {
      setMigrateError((e as Error).message);
    }
  };

  /** Botão "Importar do Thunderbird…" — procura os perfis com .mbox e abre o diálogo. */
  const importTb = async () => {
    setTbBusy(true);
    setTbError('');
    try {
      const profiles = await window.outlook.scanThunderbirdMail();
      setTbProfiles(profiles);
      setTbSrc(profiles[0]?.path || '');
      setTbName(profiles[0] ? `Thunderbird — ${profiles[0].name}` : '');
      setTbDir('');
      setTbTarget('');
      setTbDone(null);
      setMigrateProgress(null);
    } catch (e) {
      setTbError((e as Error).message);
      setTbProfiles([]);
    } finally {
      setTbBusy(false);
    }
  };

  const runTbImport = async () => {
    if (!tbSrc) return;
    if (!tbTarget && !tbName.trim()) return;
    setTbError('');
    try {
      const dest = await window.outlook.importThunderbird(
        tbSrc,
        tbDir,
        tbName.trim(),
        tbTarget || undefined,
      );
      setTbDone(dest.label);
      onAccountsChanged();
      // Fecha tudo — a base nova já aparece na árvore ao fundo.
      setTbProfiles(null);
      onClose();
    } catch (e) {
      setTbError((e as Error).message);
    }
  };

  // Ação vinda do Arquivo (Abrir → Importar .pst / Thunderbird) — dispara o
  // fluxo correspondente uma única vez ao montar.
  const actionRan = useRef(false);
  useEffect(() => {
    if (actionRan.current || !initialAction) return;
    actionRan.current = true;
    if (initialAction === 'importPst') void importPst();
    else if (initialAction === 'importTb') void importTb();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const reloadAccounts = () => window.outlook.listAccounts().then(setAccounts);
  const reloadRules = () => window.outlook.listRules().then(setRules);
  const reloadSignatures = () =>
    Promise.resolve(window.outlook.listSignatures?.() ?? [])
      .then(setSignatures)
      .catch(() => {});

  useEffect(() => {
    reloadAccounts();
    reloadRules();
    reloadSignatures();
    window.outlook.isOutlookImportAvailable().then(setImportAvailable);
    window.outlook.getLoadImages?.().then((v) => setLoadImages(!!v));
    window.outlook.getSyncInterval?.().then((v) => setSyncMin(v));
    window.outlook.getVirusScan?.().then((v) => setVirusScan(!!v));
    window.outlook.getAvStatus?.().then((s) => setAvEngine(s.engine || ''));
    window.outlook.getDefaultStore?.().then((v) => setDefaultStore(v || ''));
    window.outlook.getTheme?.().then((t) => setThemeState(t || 'system'));
    window.outlook.getDataDir?.().then((v) => setDataDirPath(v || ''));
  }, []);

  const pstSources = sources.filter((s) => s.kind === 'pst');
  const dataSources = sources.filter(
    (s) => (s.kind === 'pst' || s.kind === 'local') && s.filePath,
  );
  // Dedup por caminho — a mesma base pode aparecer 2x se foi aberta e importada
  const seenPaths = new Set<string>();
  const uniqueSources = dataSources.filter((s) => {
    const key = s.filePath!.toLowerCase();
    if (seenPaths.has(key)) return false;
    seenPaths.add(key);
    return true;
  });

  /** Vincula a conta escolhida à base recém-criada (upsert mantém a senha
   *  salva — password vazio não sobrescreve o segredo no accountStore). */
  const bindAccount = async (accountId: string, storePath: string) => {
    const acct = accounts.find((a) => a.id === accountId);
    if (!acct || !storePath) return;
    await window.outlook.upsertAccount({ ...acct, deliveryPath: storePath, password: '' });
  };

  const createStore = async (inCustomDir: boolean) => {
    if (!newStoreName.trim()) return;
    let src: MailSource | null = null;
    if (inCustomDir) {
      const dir = await window.outlook.pickStoreDir();
      if (!dir) return;
      src = await window.outlook.createLocalStoreIn(newStoreName.trim(), dir);
    } else {
      src = await window.outlook.createLocalStoreAuto(newStoreName.trim());
    }
    // Vínculo opcional: só entrega as mensagens da conta nesta base se o
    // usuário escolheu uma conta no dropdown.
    if (src?.filePath && newStoreAccount) await bindAccount(newStoreAccount, src.filePath);
    setNewStoreName('');
    if (src) onAccountsChanged();
  };

  // Nomes de pastas existentes na base alvo da regra — dropdown com sugestões
  useEffect(() => {
    const path = ruleDraft?.targetPstPath;
    const src = sources.find((s) => s.filePath === path);
    if (!path || !src) {
      setFolderChoices([]);
      return;
    }
    window.outlook
      .getFolderTree(src.id)
      .then((nodes) => {
        // Árvore hierárquica: o valor é o caminho "Pai/Filho" e o rótulo mostra
        // a indentação — como o seletor de pastas do Outlook, que permite
        // escolher uma pasta dentro de outra.
        const names: Array<{ path: string; label: string }> = [];
        const walk = (ns: typeof nodes, prefix: string, depth: number) =>
          ns.forEach((n) => {
            const clean = n.name.replace(/ \(local\)$/, '');
            const path = prefix ? `${prefix}/${clean}` : clean;
            names.push({ path, label: `${'    '.repeat(depth)}${clean}` });
            walk(n.children, path, depth + 1);
          });
        walk(nodes, '', 0);
        setFolderChoices(names);
      })
      .catch(() => setFolderChoices([]));
  }, [ruleDraft?.targetPstPath, sources]);

  const startNew = () => {
    setDraft(EMPTY_DRAFT);
    setAcctTab('conta');
    setTestMsg('');
    setAcctNewData(false);
    setAcctNewDataName('');
    setEditing(true);
  };

  const startEdit = (a: AccountConfig) => {
    setDraft({
      ...a,
      password: '',
      // Normaliza o marcador legado "(nenhuma)" para campo vazio
      signatureNewId: a.signatureNewId === '__none' ? undefined : a.signatureNewId,
      signatureReplyId: a.signatureReplyId === '__none' ? undefined : a.signatureReplyId,
    });
    setAcctTab('conta');
    setTestMsg('');
    setAcctNewData(false);
    setAcctNewDataName('');
    setEditing(true);
  };

  // Arquivo → "Configurações de Conta" com uma conta selecionada: abre a janela
  // já editando essa conta (assim que a lista carregar). Consumido uma vez.
  const pendingEdit = useRef(editAccountId);
  useEffect(() => {
    if (!pendingEdit.current || accounts.length === 0) return;
    const a = accounts.find((x) => x.id === pendingEdit.current);
    if (a) {
      pendingEdit.current = undefined;
      setTab('email');
      startEdit(a);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accounts]);

  // Se a conta importada apontava um PST de entrega, só vincula se essa base
  // estiver ABERTA aqui (casa por caminho exato ou nome do arquivo) — nunca
  // cria uma base nova silenciosamente.
  const matchDeliveryPath = (g: ImportedAccountGuess): string | undefined => {
    const baseName = g.pstPath?.split(/[\\/]/).pop()?.toLowerCase();
    return g.pstPath
      ? uniqueSources.find(
          (s) =>
            s.filePath?.toLowerCase() === g.pstPath!.toLowerCase() ||
            s.filePath?.split(/[\\/]/).pop()?.toLowerCase() === baseName,
        )?.filePath
      : undefined;
  };

  /** Converte um resultado da busca em conta a salvar (arquivo de dados + senha). */
  const draftFromGuess = (g: ImportedAccountGuess): AccountDraft => ({
    ...EMPTY_DRAFT,
    label: g.displayName || g.email,
    email: g.email,
    protocol: g.protocol,
    imapHost: g.imapHost,
    imapPort: g.imapPort,
    imapSecure: g.imapSecure,
    imapUser: g.imapUser || g.email,
    smtpHost: g.smtpHost,
    smtpPort: g.smtpPort,
    smtpSecure: g.smtpSecure,
    smtpUser: g.smtpUser || g.email,
    deliveryPath: matchDeliveryPath(g),
    leaveOnServer: g.leaveOnServer ?? true,
    leaveDays: g.leaveDays ?? 0,
    password: g.password || '',
    smtpPassword: g.smtpPassword && g.smtpPassword !== g.password ? g.smtpPassword : undefined,
  });

  /** Assinaturas da importação viram sugestão na aba Assinaturas — nada é criado sozinho. */
  const collectSigs = (g: ImportedAccountGuess) => {
    const found: Array<{ name: string; html: string }> = [];
    if (g.signatureHtml) found.push({ name: `Assinatura — ${g.email}`, html: g.signatureHtml });
    if (g.signatureReplyHtml && g.signatureReplyHtml !== g.signatureHtml)
      found.push({ name: `Resposta — ${g.email}`, html: g.signatureReplyHtml });
    if (found.length) setImportedSigs((prev) => [...prev, ...found]);
  };

  /** Abre a etapa de configuração para a conta escolhida no assistente. */
  const openImport = (g: ImportedAccountGuess) => {
    const d = draftFromGuess(g);
    setImp({
      g,
      label: d.label || g.email,
      password: g.password || '',
      path: d.deliveryPath || defaultStore || '',
      newData: false,
      newName: '',
      leave: d.leaveOnServer ?? true,
      leaveDays: d.leaveDays ?? 0,
    });
    setImpErr('');
    setImpTest('');
    setWizStep('form');
  };

  /** Testa a conta com os dados da etapa de configuração, antes de importar. */
  const doTest = async () => {
    if (!imp) return;
    setImpBusy(true);
    setImpTest('');
    // Erro de autenticação → explica que a senha recuperada pode ser antiga
    // (o Outlook só guarda a última usada) e mostra o usuário tentado.
    const friendly = (m: string) =>
      /auth|pass|senha|credential|535|err\b/i.test(m)
        ? `${m} — usuário tentado: "${imp.g.imapUser || imp.g.email}". Se a senha recuperada estiver desatualizada, digite a senha atual do e-mail.`
        : m;
    try {
      const d = draftFromGuess(imp.g);
      const res = await window.outlook.testAccount({
        ...d,
        label: imp.label,
        password: imp.password,
        smtpPassword:
          imp.g.smtpPassword && imp.g.smtpPassword !== imp.password ? imp.g.smtpPassword : undefined,
      });
      setImpTest(res.ok ? `✓ ${res.message}` : `✗ ${friendly(res.message)}`);
    } catch (e) {
      setImpTest(`✗ ${friendly((e as Error).message)}`);
    }
    setImpBusy(false);
  };

  /** Cria a conta com as opções escolhidas no modal de importação. */
  const doImport = async () => {
    if (!imp) return;
    setImpBusy(true);
    setImpErr('');
    try {
      const d = draftFromGuess(imp.g);
      let deliveryPath = imp.path || undefined;
      if (imp.newData) {
        const src = await window.outlook.createLocalStoreAuto(
          imp.newName.trim() || `Emails — ${imp.label}`,
        );
        if (!src?.filePath) throw new Error('Não foi possível criar o arquivo de dados');
        deliveryPath = src.filePath;
      }
      const saved = await window.outlook.upsertAccount({
        ...d,
        label: imp.label,
        deliveryPath,
        deliveryFolder: undefined,
        password: imp.password,
        smtpPassword:
          imp.g.smtpPassword && imp.g.smtpPassword !== imp.password ? imp.g.smtpPassword : undefined,
        leaveOnServer: imp.leave,
        leaveDays: imp.leaveDays,
      });
      await window.outlook.connectAccount(saved.id);
      collectSigs(imp.g);
      const em = imp.g.email.toLowerCase();
      setDoneEmails((prev) => new Set(prev).add(em));
      reloadAccounts();
      onAccountsChanged();
      // Volta para a tela de contas — o assistente fecha após cada importação.
      setImp(null);
      setWizPick('');
      setWizStep('choose');
      setGuesses(null);
      setScanMsg(`Conta ${imp.g.email} importada.`);
    } catch (e) {
      setImpErr((e as Error).message);
    }
    setImpBusy(false);
  };

  const applyGuess = (g: ImportedAccountGuess) => {
    const matchPath = matchDeliveryPath(g);
    collectSigs(g);
    setAcctNewData(false);
    setAcctNewDataName('');
    setDraft({
      ...draftFromGuess(g),
      deliveryPath: matchPath,
      leaveOnServer: g.leaveOnServer ?? true,
      leaveDays: g.leaveDays ?? 0,
      // Senha recuperada do Registro (DPAPI) nesta sessão/máquina — nunca
      // sai para o arquivo de exportação, só preenche o formulário aqui.
      password: g.password || '',
    });
    setTestMsg(
      (g.password
        ? `Configuração e senha recuperadas automaticamente (${g.protocol.toUpperCase()}). Revise e clique em Salvar.`
        : `Configuração importada (${g.protocol.toUpperCase()}). Digite a senha para concluir.`) +
        (g.pstPath
          ? matchPath
            ? ` Entrega vinculada à base "${pstLabel(matchPath)}".`
            : ` O PST de entrega do Outlook (${g.pstPath.split(/[\\/]/).pop()}) não foi importado — importe-o na aba "Arquivos de Dados" ou escolha outro arquivo de dados na seção "Arquivo de dados" da conta.`
          : ''),
    );
    setAcctTab('conta');
    setEditing(true);
  };

  /** Abre o editor de assinatura com os vínculos de conta já carregados. */
  const openSignature = (s: Signature) => {
    setSigDraft({ ...s });
    setSigAssign(
      Object.fromEntries(
        accounts.map((a) => [
          a.id,
          { new: a.signatureNewId === s.id, reply: a.signatureReplyId === s.id },
        ]),
      ),
    );
  };

  /** Salva a assinatura e aplica os vínculos com as contas marcados na matriz. */
  const saveSignature = async () => {
    if (!sigDraft) return;
    const saved = await window.outlook.upsertSignature(sigDraft);
    for (const a of accounts) {
      const asg = sigAssign[a.id];
      if (!asg) continue;
      const wantNew = asg.new ? saved.id : a.signatureNewId === saved.id ? undefined : a.signatureNewId;
      const wantReply = asg.reply ? saved.id : a.signatureReplyId === saved.id ? undefined : a.signatureReplyId;
      if (wantNew !== a.signatureNewId || wantReply !== a.signatureReplyId) {
        await window.outlook.upsertAccount({
          ...a,
          signatureNewId: wantNew,
          signatureReplyId: wantReply,
          password: '',
        });
      }
    }
    setSigDraft(null);
    setSigAssign({});
    reloadSignatures();
    reloadAccounts();
  };

  const onEmailBlur = () => {
    if (!draft.imapHost && !draft.smtpHost && draft.email.includes('@')) {
      const g = guessServers(draft.email);
      setDraft((d) => ({
        ...d,
        imapHost: d.imapHost || g.imapHost,
        smtpHost: d.smtpHost || g.smtpHost,
        imapUser: d.imapUser || d.email,
        smtpUser: d.smtpUser || d.email,
        label: d.label || d.email,
      }));
    }
  };

  const test = async () => {
    setBusy(true);
    setTestMsg('Testando…');
    try {
      const res = await window.outlook.testAccount(draft);
      setTestMsg(res.ok ? `✓ ${res.message}` : `✗ ${res.message}`);
    } finally {
      setBusy(false);
    }
  };

  const save = async () => {
    setBusy(true);
    try {
      // "Criar novo arquivo de dados" — cria a base local e vincula só a esta conta.
      // Sem escolha explícita, deliveryPath vazio = arquivo principal do app.
      // A entrega é sempre na Caixa de Entrada; a escolha de pasta fica com as regras.
      let toSave: AccountDraft = { ...draft, deliveryPath: draft.deliveryPath || undefined, deliveryFolder: undefined };
      if (acctNewData) {
        const src = await window.outlook.createLocalStoreAuto(
          acctNewDataName.trim() || `Emails — ${draft.label || draft.email}`,
        );
        if (!src?.filePath) throw new Error('Não foi possível criar o arquivo de dados');
        toSave = { ...toSave, deliveryPath: src.filePath };
      }
      const saved = await window.outlook.upsertAccount(toSave);
      await window.outlook.connectAccount(saved.id);
      setEditing(false);
      reloadAccounts();
      onAccountsChanged();
    } catch (e) {
      setTestMsg(`Erro ao salvar: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  };

  const remove = (id: string) => {
    const acc = accounts.find((a) => a.id === id);
    // Base local vinculada SÓ a esta conta — por padrão é apagada junto (um
    // .pst nunca é excluído: é o arquivo original do usuário).
    const boundLocal = uniqueSources.find((s) => s.kind === 'local' && s.filePath === acc?.deliveryPath);
    setConfirm({
      title: 'Remover conta',
      message: boundLocal
        ? `Remover a conta "${acc?.email || id}"? A base de dados vinculada "${boundLocal.label}" também será excluída — todas as mensagens nela serão apagadas do disco.`
        : `Remover a conta "${acc?.email || id}"? ${acc?.deliveryPath ? 'A base de dados vinculada (PST) não é apagada.' : ''}`,
      onOk: () => {
        setConfirm(null);
        void window.outlook.removeAccount(id, !!boundLocal).then(() => {
          reloadAccounts();
          onAccountsChanged();
        });
      },
    });
  };

  /** "Re-sincronizar": esquece o que a conta já processou e re-baixa tudo do
   *  servidor. Recuperação de mensagens marcadas como entregues por engano
   *  (base restaurada de outro SO com caminhos inválidos, por exemplo). */
  const resetSync = (id: string) => {
    const acc = accounts.find((a) => a.id === id);
    setConfirm({
      title: 'Re-sincronizar do servidor',
      message:
        `Re-baixar todas as mensagens de "${acc?.email || id}" que ainda estiverem no servidor? ` +
        'O que já está na base não duplica, mas a sincronização pode demorar se a caixa for grande.',
      onOk: () => {
        setConfirm(null);
        void window.outlook.resetSyncState(id).then(() => {
          setTestMsg('Sincronização reiniciada — use Enviar/Receber para baixar novamente.');
        });
      },
    });
  };

  const deleteStore = (s: MailSource) => {
    setConfirm({
      title: 'Excluir base de dados',
      message: `Excluir "${s.label}"? Todas as mensagens guardadas nela serão apagadas do disco — isso não pode ser desfeito.`,
      onOk: () => {
        setConfirm(null);
        void window.outlook.deleteLocalStore(s.filePath!).then(() => onAccountsChanged());
      },
    });
  };

  const scanAccounts = async (src: 'outlook' | 'thunderbird') => {
    const name = src === 'outlook' ? 'Outlook' : 'Thunderbird';
    // Abre o assistente na hora com "procurando…" na área da lista — a busca
    // no Registro/perfis pode levar segundos e, sem feedback imediato,
    // parecia que o botão (ou o app) tinha travado.
    setGuesses([]);
    setWizStep('choose');
    setWizPick('');
    setImp(null);
    setWizNote('');
    setScanMsg('');
    setWizScanning(true);
    try {
      const found =
        src === 'outlook' ? await window.outlook.scanOutlookAccounts() : await window.outlook.scanThunderbirdAccounts();
      loadGuesses(found);
      const pwdNote = found.some((g) => g.password)
        ? 'Senhas recuperadas onde estavam salvas no Windows.'
        : src === 'thunderbird'
          ? 'O Thunderbird guarda as senhas cifradas: digite a senha ao usar os dados.'
          : 'Digite a senha ao usar os dados.';
      if (found.length) {
        // O resultado abre no assistente — nada é escrito na tela de fundo.
        setScanMsg('');
        setWizNote(`${found.length} conta(s) encontrada(s) no ${name}. ${pwdNote}`);
      } else {
        setWizNote('');
        setScanMsg(`Nenhuma conta encontrada no ${name}.`);
      }
    } catch (e) {
      // Erro na busca: fecha o assistente e mostra a mensagem na tela de fundo
      setGuesses(null);
      setScanMsg(`Erro: ${(e as Error).message}`);
    } finally {
      setWizScanning(false);
    }
  };

  const saveExport = async () => {
    if (!guesses) return;
    const path = await window.outlook.saveOutlookExport(guesses);
    if (path) setWizNote(`Exportado para ${path}. Copie este arquivo para o Linux e use "Arquivo" lá.`);
  };

  const openExport = async () => {
    try {
      const found = await window.outlook.openOutlookExportFile();
      if (found) {
        loadGuesses(found);
        setScanMsg('');
        setWizNote(`${found.length} conta(s) carregada(s) do arquivo de exportação — digite a senha de cada uma ao importar.`);
      }
    } catch (e) {
      setScanMsg(`Erro: ${(e as Error).message}`);
    }
  };

  // ---------- Regras ----------
  const EMPTY_RULE: RuleDraft = {
    name: '',
    accountId: '',
    enabled: true,
    conditions: {},
    targetPstPath: '',
    targetFolder: 'Caixa de Entrada',
  };

  const saveRule = async () => {
    if (!ruleDraft) return;
    await window.outlook.upsertRule(ruleDraft);
    setRuleDraft(null);
    reloadRules();
  };

  const editRule = (r: MailRule) => {
    setCustomFolder(false);
    setRuleDraft({ ...r, conditions: { ...r.conditions } });
  };

  /** Duplica a regra como uma NOVA (sem id) e abre o editor — a conta fica
   *  vazia de propósito para obrigar a escolha do email de chegada. */
  const duplicateRule = (r: MailRule) => {
    setCustomFolder(false);
    setRuleDraft({
      name: `${r.name || 'Regra'} (cópia)`,
      accountId: '',
      enabled: r.enabled,
      conditions: { ...r.conditions },
      targetPstPath: r.targetPstPath,
      targetFolder: r.targetFolder,
      stopProcessing: r.stopProcessing,
    });
  };

  const toggleRule = async (r: MailRule) => {
    await window.outlook.upsertRule({ ...r, enabled: !r.enabled });
    reloadRules();
  };

  const removeRule = (id: string) => {
    const rule = rules.find((r) => r.id === id);
    setConfirm({
      title: 'Excluir regra',
      message: `Excluir a regra "${rule?.name || '(sem nome)'}"?`,
      onOk: () => {
        setConfirm(null);
        void window.outlook.removeRule(id).then(reloadRules);
      },
    });
  };

  const accountLabel = (id: string) =>
    id === 'all' ? 'Qualquer conta' : accounts.find((a) => a.id === id)?.email || id;
  const dataFiles = uniqueSources.filter((s) => (s.kind === 'local' || s.kind === 'pst') && s.filePath);
  const accountMissing = (id: string) => id !== 'all' && !accounts.some((a) => a.id === id);
  const pstLabel = (p: string) => sources.find((s) => s.filePath === p)?.label || p.split(/[\\/]/).pop() || p;

  const saveRename = async () => {
    if (!renaming?.name.trim()) return;
    try {
      await window.outlook.renameLocalStore(renaming.path, renaming.name.trim());
      setRenaming(null);
      onAccountsChanged(); // reabre lista de fontes com o novo label
    } catch (e) {
      setScanMsg(`Erro ao renomear: ${(e as Error).message}`);
    }
  };

  const tabIcons: Record<SettingsTab, string> = {
    email: 'envelope',
    data: 'archive',
    rules: 'filter',
    sigs: 'pen',
    geral: 'settings',
  };

  return (
    <div className="set-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="set-window" onClick={(e) => e.stopPropagation()}>
        <div className="set-head">
          <div className="set-head-icon">
            <Icon name="settings" size={18} />
          </div>
          <div>
            <h2>Configurações</h2>
            <div className="sub">Contas, arquivos de dados, regras e preferências</div>
          </div>
          <button className="set-close" onClick={onClose} title="Fechar">
            <Icon name="close" size={16} />
          </button>
        </div>

        {/* Navegação em pílulas */}
        <div className="set-tabs">
          {(
            [
              ['email', 'Email'],
              ['data', 'Arquivos de Dados'],
              ['rules', 'Regras e Alertas'],
              ['sigs', 'Assinaturas'],
              ['geral', 'Geral'],
            ] as Array<[SettingsTab, string]>
          ).map(([id, label]) => (
            <button
              key={id}
              onClick={() => setTab(id)}
              className={`set-tab${tab === id ? ' active' : ''}`}
            >
              <Icon name={tabIcons[id]} size={14} />
              {label}
            </button>
          ))}
        </div>

        {/* Conteúdo das abas */}
        <div className="set-body">

        {/* Bases de dados (PST + locais) */}
        {tab === 'data' && (
        <>
        <section className="set-card fill">
          <h3 className="set-card-title"><Icon name="archive" size={15} /> Arquivos de dados</h3>
          <p className="set-hint">
            Onde as mensagens são guardadas. Marque uma como <strong>Principal</strong> — mensagens
            sem regra e contas sem base vinculada são entregues na Caixa de Entrada dela.
          </p>

          {/* Lista de bases abertas — cartões (rola dentro do card) */}
          <div className="fill-scroll">
          {uniqueSources.length === 0 && (
            <div className="empty-state">
              Nenhuma base aberta. Crie uma base local ou importe um .pst abaixo.
            </div>
          )}
          {uniqueSources.map((s) => (
            <div key={s.id} className="acct-card">
              <div className="acct-avatar" style={s.kind === 'local' ? { background: 'linear-gradient(135deg,#107c10,#0b5a0b)' } : undefined}>
                <Icon name={s.kind === 'pst' ? 'archive' : 'folder'} size={17} />
              </div>
              {renaming !== null && renaming.path === s.filePath ? (
                <>
                  <input
                    className="inp"
                    style={{ flex: 1 }}
                    value={renaming.name}
                    autoFocus
                    onChange={(e) => setRenaming({ path: s.filePath!, name: e.target.value })}
                    onKeyDown={(e) => e.key === 'Enter' && saveRename()}
                  />
                  <button onClick={saveRename} className="btn btn-primary">Salvar</button>
                  <button onClick={() => setRenaming(null)} className="btn">Cancelar</button>
                </>
              ) : (
                <>
                  <span className="acct-info">
                    <span className="acct-name">
                      {s.label}
                      <span className={`badge ${s.kind === 'pst' ? 'badge-pst' : 'badge-local'}`}>
                        {s.kind === 'pst' ? 'PST · somente leitura' : 'Base local'}
                      </span>
                      {defaultStore === s.filePath && (
                        <span className="badge badge-main">Principal</span>
                      )}
                    </span>
                    <span className="acct-sub" title={s.filePath}>
                      {(() => {
                        const bound = accounts.find((a) => a.deliveryPath === s.filePath);
                        return (
                          <>
                            {bound ? (
                              <span style={{ color: 'var(--blue)' }}>vinculada a {bound.email} · </span>
                            ) : (
                              <span style={{ color: 'var(--text-dim)' }}>sem conta vinculada · </span>
                            )}
                            {s.filePath}
                          </>
                        );
                      })()}
                    </span>
                  </span>
                  {defaultStore !== s.filePath && (
                    <button
                      className="btn"
                      title="Tornar esta a base principal (destino padrão das mensagens)"
                      onClick={() => {
                        setDefaultStore(s.filePath!);
                        void window.outlook.setDefaultStore?.(s.filePath!);
                      }}
                    >
                      Definir principal
                    </button>
                  )}
                  {s.kind === 'local' && (
                    <button
                      onClick={() => setRenaming({ path: s.filePath!, name: s.label })}
                      className="btn btn-ghost"
                      title="Renomear a base de dados"
                    >
                      Renomear
                    </button>
                  )}
                  {s.kind === 'pst' && (
                    <button
                      onClick={() => {
                        setMigrateCloseAfter(false);
                        setMigrateSrc(s);
                      }}
                      className="btn btn-ghost"
                      title="Copia todo o conteúdo do PST para uma base própria do app — lá, excluir apaga de verdade e libera espaço"
                    >
                      Migrar…
                    </button>
                  )}
                  {s.kind === 'local' ? (
                    <button
                      onClick={() => deleteStore(s)}
                      className="btn btn-ghost danger"
                      title="Apaga a pasta e todas as mensagens do disco — não pode ser desfeito"
                    >
                      Excluir
                    </button>
                  ) : (
                    <button
                      onClick={() => onCloseSource(s.id)}
                      className="btn btn-ghost"
                      title="Fecha o PST no app — o arquivo original nunca é apagado"
                    >
                      Fechar
                    </button>
                  )}
                </>
              )}
            </div>
          ))}
          </div>
        </section>

        <section className="set-card">
          <h3 className="set-card-title"><Icon name="new" size={15} /> Adicionar arquivo de dados</h3>
          <p className="set-hint" style={{ marginTop: 0 }}>
            Dois jeitos de adicionar dados ao app — os dois criam bases no formato próprio
            do Unit Mail (pasta com mensagens .eml, portável para o Linux):
          </p>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button
              className="btn btn-primary"
              onClick={() => {
                setNewStoreName('');
                setNewStoreAccount('');
                setShowNewStore(true);
              }}
            >
              <Icon name="new" size={14} /> Novo arquivo de dados…
            </button>
            <button className="btn" onClick={importPst}>
              <Icon name="download" size={14} /> Importar arquivo .pst…
            </button>
            <button className="btn" onClick={importTb} disabled={tbBusy}>
              <Icon name="download" size={14} /> {tbBusy ? 'Procurando perfis…' : 'Importar do Thunderbird…'}
            </button>
          </div>
          <p className="set-hint" style={{ marginTop: 8, marginBottom: 0 }}>
            <strong>Novo arquivo de dados</strong> cria uma base vazia — pode receber mensagens de
            várias contas na mesma Caixa de Entrada (use Regras para separar em pastas).{' '}
            <strong>Importar .pst</strong> traz TUDO que estiver no arquivo do Outlook — todas as
            pastas e mensagens — para a base que você escolher; o .pst original nunca é alterado.{' '}
            <strong>Importar do Thunderbird</strong> traz todas as pastas e mensagens do perfil
            instalado neste computador — o perfil original não é alterado.
          </p>
        </section>

        <section className="set-card">
          <h3 className="set-card-title"><Icon name="folder" size={15} /> Pasta de dados</h3>
          <p className="set-hint">
            Todos os seus dados ficam nesta pasta — contas, regras, assinaturas e as bases de
            email dentro de <code>Mail/</code> (como o perfil do Thunderbird). Para mover para
            outro computador, copie a pasta inteira e aponte o app para ela aqui: tudo volta a
            funcionar sem recadastrar nada.
          </p>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <input className="inp" style={{ flex: 1 }} readOnly value={dataDirPath} title={dataDirPath} />
            <button
              className="btn"
              style={{ whiteSpace: 'nowrap' }}
              onClick={async () => {
                const res = await window.outlook.pickDataDirPath?.();
                if (!res) return;
                setConfirm({
                  title: 'Alterar pasta de dados',
                  message: res.isProfile
                    ? `"${res.dir}" já contém um perfil do Unit Mail — contas, regras e bases serão restaurados de lá. Confirmar?`
                    : `Mover seus dados para "${res.dir}"? As bases locais serão copiadas para a subpasta Mail/ e as referências atualizadas.`,
                  onOk: () => {
                    setConfirm(null);
                    void window.outlook.applyDataDir?.(res.dir).then((applied) => {
                      if (!applied) return;
                      setDataDirPath(applied.dir);
                      if (applied.adopted) {
                        // Perfil restaurado — reinicia sozinho
                        void window.outlook.restartApp?.();
                        return;
                      }
                      setConfirm({
                        title: 'Pasta de dados alterada',
                        message: 'Dados copiados. Reinicie o app para concluir.',
                        onOk: () => void window.outlook.restartApp?.(),
                      });
                    });
                  },
                });
              }}
            >
              Alterar pasta…
            </button>
            <button
              className="btn btn-ghost"
              title="Abrir a pasta de dados no explorador"
              onClick={() => void window.outlook.openDataDir()}
            >
              Abrir
            </button>
          </div>
        </section>
        </>
        )}

        {/* Contas de email */}
        {tab === 'email' && (
        <>
        <section className="set-card fill">
          <div className="set-card-head">
            <h3 className="set-card-title"><Icon name="envelope" size={15} /> Contas de email</h3>
            <div className="acct-toolbar">
              {importAvailable && (
                <button
                  className="btn btn-primary btn-sm"
                  title="Procurar contas do Outlook instalado no Windows"
                  onClick={() => scanAccounts('outlook')}
                >
                  <Icon name="search" size={14} /> Outlook
                </button>
              )}
              <button
                className="btn btn-primary btn-sm"
                title="Procurar contas nos perfis do Thunderbird"
                onClick={() => scanAccounts('thunderbird')}
              >
                <Icon name="mail" size={14} /> Thunderbird
              </button>
              <button
                className="btn btn-primary btn-sm"
                title="Importar arquivo de exportação de contas (.json)"
                onClick={openExport}
              >
                <Icon name="download" size={14} /> Arquivo
              </button>
            </div>
          </div>
          {scanMsg && <div className="scan-msg">{scanMsg}</div>}
          <p className="set-hint">Recebimento por POP3 ou IMAP; envio por SMTP — igual ao Outlook.</p>
          {!secureStore && (
            <div className="set-warning" role="alert">
              O sistema não oferece armazenamento seguro de senhas (keyring indisponível).
              As senhas das contas ficam sem criptografia no arquivo de dados. Instale e
              ative um cofre de senhas (GNOME Keyring ou KWallet) e recadastre as contas.
            </div>
          )}
          {accounts.length === 0 && (
            <div className="empty-state">Nenhuma conta cadastrada. Clique em "Adicionar conta" para começar.</div>
          )}
          {accounts.length > 0 && (
            <div className="acct-table fill-scroll" role="table">
              <div className="acct-row acct-head" role="row">
                <span role="columnheader">Conta</span>
                <span role="columnheader">Recebimento</span>
                <span role="columnheader">Entrega</span>
                <span role="columnheader">Situação</span>
                <span role="columnheader" />
              </div>
              {accounts.map((a) => {
                const src = sources.find((s) => s.id === `acct:${a.id}`);
                const state = src?.connecting ? 'wait' : src?.error ? 'err' : 'ok';
                const stateLabel = src?.connecting ? 'Conectando…' : src?.error ? 'Erro de conexão' : 'Pronta';
                return (
                  <div key={a.id} className="acct-row" role="row">
                    <div className="acct-cell acct-who" role="cell">
                      <div className="acct-avatar">{(a.label || a.email || '?').trim().charAt(0).toUpperCase()}</div>
                      <div className="acct-info">
                        <div className="acct-name">{a.label || a.email}</div>
                        <div className="acct-sub">{a.email}</div>
                      </div>
                    </div>
                    <div className="acct-cell" role="cell">
                      <span>
                        <span className={`badge ${a.protocol === 'pop3' ? 'badge-pop3' : 'badge-imap'}`}>
                          {a.protocol.toUpperCase()}
                        </span>
                      </span>
                      <span className="acct-sub">{a.imapHost}:{a.imapPort}</span>
                    </div>
                    <div className="acct-cell" role="cell">
                      {/* Resolve a base efetiva: a vinculada à conta ou a
                          principal — assim a mesma base sempre mostra o mesmo
                          nome, e "(principal)" indica a padrão. */}
                      {(() => {
                        const eff = a.deliveryPath || defaultStore;
                        if (!eff) return <span className="acct-sub">Base padrão</span>;
                        return (
                          <span className="acct-sub acct-strong" title={eff}>
                            {pstLabel(eff)}
                            {eff === defaultStore ? ' (principal)' : ''}
                          </span>
                        );
                      })()}
                    </div>
                    <div className="acct-cell" role="cell">
                      <span className={`acct-state acct-state-${state}`} title={src?.error || ''}>
                        <span className={`status-dot ${state === 'wait' ? 'status-idle' : state === 'err' ? 'status-err' : 'status-ok'}`} />
                        {stateLabel}
                      </span>
                    </div>
                    <div className="acct-cell acct-actions" role="cell">
                      <button onClick={() => startEdit(a)} className="btn">Editar</button>
                      <button
                        onClick={() => resetSync(a.id)}
                        className="btn"
                        title="Esquece o que já foi baixado e re-sincroniza tudo do servidor"
                      >
                        Re-sincronizar
                      </button>
                      <button onClick={() => remove(a.id)} className="btn btn-ghost danger">Remover</button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
          <button className="btn btn-primary" style={{ marginTop: 4 }} onClick={startNew}>
            <Icon name="new" size={14} /> Adicionar conta de email
          </button>

          {editing && (
            <div className="modal-overlay" style={{ zIndex: 140 }} onMouseDown={(e) => e.target === e.currentTarget && setEditing(false)}>
              <div
                className="set-card acct-modal"
                style={{ display: 'flex', flexDirection: 'column', gap: 12 }}
                onClick={(e) => e.stopPropagation()}
              >
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <strong style={{ fontSize: 14 }}>
                  {draft.id ? `Editar conta — ${draft.email || ''}` : 'Nova conta de email'}
                </strong>
                <button className="set-close" onClick={() => setEditing(false)} title="Fechar">
                  <Icon name="close" size={15} />
                </button>
              </div>
              {/* Abas do formulário de conta — divide a tela em partes menores */}
              <div className="set-tabs" style={{ marginBottom: -2 }}>
                {(
                  [
                    ['conta', 'Conta'],
                    ['servidores', 'Servidores'],
                  ] as Array<[typeof acctTab, string]>
                ).map(([id, label]) => (
                  <button
                    key={id}
                    onClick={() => setAcctTab(id)}
                    className={`set-tab${acctTab === id ? ' active' : ''}`}
                  >
                    {label}
                  </button>
                ))}
              </div>

              {acctTab === 'conta' && (
              <>
              <div className="grp-title">Informações do usuário</div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
                <Field label="Nome de exibição">
                  <input className="inp" value={draft.label} onChange={(e) => setDraft({ ...draft, label: e.target.value })} />
                </Field>
                <Field label="Endereço de email">
                  <input
                    className="inp"
                    value={draft.email}
                    onChange={(e) => setDraft({ ...draft, email: e.target.value })}
                    onBlur={onEmailBlur}
                  />
                </Field>
                <Field label="Protocolo">
                  <select
                    className="inp"
                    value={draft.protocol}
                    onChange={(e) =>
                      setDraft({
                        ...draft,
                        protocol: e.target.value as 'imap' | 'pop3',
                        imapPort: e.target.value === 'pop3' ? (draft.imapSecure ? 995 : 110) : draft.imapSecure ? 993 : 143,
                      })
                    }
                  >
                    <option value="pop3">POP3</option>
                    <option value="imap">IMAP</option>
                  </select>
                </Field>
              </div>
              <div className="grp-title">Logon</div>
              <Field label="Senha (armazenada criptografada localmente)">
                <input
                  type="password"
                  className="inp"
                  value={draft.password}
                  placeholder={draft.id && draft.hasPassword ? '•••••••• (senha salva)' : draft.id ? '(deixe em branco para manter a atual)' : ''}
                  onChange={(e) => setDraft({ ...draft, password: e.target.value })}
                />
                {draft.id && draft.hasPassword && (
                  <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                    ✓ Senha salva — preencha apenas se quiser trocar.
                  </span>
                )}
              </Field>
              <div className="grp-title">Arquivo de dados</div>
              <p className="set-hint" style={{ margin: 0 }}>
                Onde as mensagens desta conta são guardadas. Você pode usar um arquivo que já existe
                ou criar um novo só para esta conta.
              </p>
              <div className="acct-data-choice" role="radiogroup">
                <label className={`acct-data-opt${!acctNewData ? ' on' : ''}`}>
                  <input type="radio" checked={!acctNewData} onChange={() => setAcctNewData(false)} />
                  <span>
                    <strong>Usar um arquivo de dados existente</strong>
                    <em>Sem escolha, a conta usa o arquivo principal. Entrega na Caixa de Entrada.</em>
                  </span>
                </label>
                <label className={`acct-data-opt${acctNewData ? ' on' : ''}`}>
                  <input type="radio" checked={acctNewData} onChange={() => setAcctNewData(true)} />
                  <span>
                    <strong>Criar um novo arquivo de dados</strong>
                    <em>Cria uma base própria ao salvar. Se for o primeiro, vira o principal.</em>
                  </span>
                </label>
              </div>
              {!acctNewData ? (
                <Field label="Arquivo de dados">
                  <select
                    className="inp"
                    value={draft.deliveryPath || defaultStore || ''}
                    disabled={dataFiles.length === 0}
                    onChange={(e) => setDraft({ ...draft, deliveryPath: e.target.value || undefined })}
                  >
                    {dataFiles.length === 0 && <option value="">Nenhum arquivo de dados aberto</option>}
                    {dataFiles.map((s) => (
                      <option key={s.id} value={s.filePath}>
                        {s.label}
                        {defaultStore === s.filePath ? ' (principal)' : ''}
                      </option>
                    ))}
                  </select>
                </Field>
              ) : (
                <Field label="Nome do novo arquivo de dados">
                  <input
                    className="inp"
                    value={acctNewDataName}
                    placeholder={`Emails — ${draft.label || draft.email || 'conta'}`}
                    onChange={(e) => setAcctNewDataName(e.target.value)}
                  />
                </Field>
              )}
              </>
              )}

              {acctTab === 'servidores' && (
              <>
              <div className="grp-title">Servidor de entrada — {draft.protocol.toUpperCase()}</div>
              <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr 1fr', gap: 10 }}>
                <Field label={`Servidor ${draft.protocol.toUpperCase()} (recebimento)`}>
                  <input className="inp" value={draft.imapHost} onChange={(e) => setDraft({ ...draft, imapHost: e.target.value })} />
                </Field>
                <Field label="Porta">
                  <input
                    type="number"
                    className="inp"
                    value={draft.imapPort}
                    onChange={(e) => setDraft({ ...draft, imapPort: Number(e.target.value) })}
                  />
                </Field>
                <Field label="Criptografia">
                  <select
                    className="inp"
                    value={draft.imapSecure ? '1' : '0'}
                    onChange={(e) =>
                      setDraft({
                        ...draft,
                        imapSecure: e.target.value === '1',
                        imapPort:
                          e.target.value === '1'
                            ? draft.protocol === 'pop3' ? 995 : 993
                            : draft.protocol === 'pop3' ? 110 : 143,
                      })
                    }
                  >
                    <option value="1">SSL/TLS</option>
                    <option value="0">Nenhum</option>
                  </select>
                </Field>
              </div>
              <div className="grp-title">Servidor de saída — SMTP</div>
              <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr 1fr', gap: 10 }}>
                <Field label="Servidor SMTP (envio)">
                  <input className="inp" value={draft.smtpHost} onChange={(e) => setDraft({ ...draft, smtpHost: e.target.value })} />
                </Field>
                <Field label="Porta">
                  <input
                    type="number"
                    className="inp"
                    value={draft.smtpPort}
                    onChange={(e) => setDraft({ ...draft, smtpPort: Number(e.target.value) })}
                  />
                </Field>
                <Field label="TLS">
                  <select
                    className="inp"
                    value={draft.smtpSecure ? '1' : '0'}
                    onChange={(e) => setDraft({ ...draft, smtpSecure: e.target.value === '1' })}
                  >
                    <option value="1">SSL/TLS</option>
                    <option value="0">STARTTLS/Nenhum</option>
                  </select>
                </Field>
              </div>
              {draft.protocol === 'pop3' && (
                <>
                  <div className="grp-title">Entrega no servidor</div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, padding: '8px 10px', background: 'var(--surface-alt)', borderRadius: 3 }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        checked={draft.leaveOnServer ?? true}
                        onChange={(e) => setDraft({ ...draft, leaveOnServer: e.target.checked })}
                      />
                      Deixar uma cópia das mensagens no servidor
                    </label>
                    <Field label="Remover do servidor após (dias — 0 = nunca)">
                      <input
                        type="number"
                        min={0}
                        className="inp"
                        disabled={!(draft.leaveOnServer ?? true)}
                        value={draft.leaveDays ?? 0}
                        onChange={(e) => setDraft({ ...draft, leaveDays: Math.max(0, Number(e.target.value) || 0) })}
                      />
                    </Field>
                  </div>
                </>
              )}
              </>
              )}

              {testMsg && <div style={{ fontSize: 12 }}>{testMsg}</div>}
              <div style={{ display: 'flex', gap: 8 }}>
                <button disabled={busy} onClick={test} className="btn">
                  Testar conexão
                </button>
                <button disabled={busy} onClick={save} className="btn btn-primary">
                  Salvar
                </button>
                <button disabled={busy} onClick={() => setEditing(false)} className="btn">
                  Cancelar
                </button>
              </div>
              </div>
            </div>
          )}
        </section>

        {/* Assistente de importação — uma conta por vez, em duas etapas.
            As informações só aparecem depois de clicar nos botões do topo. */}
        {guesses !== null && (
            <div className="set-overlay" style={{ zIndex: 120 }} onMouseDown={(e) => e.target === e.currentTarget && !impBusy && setGuesses(null)}>
              <div className="set-window" style={{ width: 580, height: 'auto' }} onClick={(e) => e.stopPropagation()}>
                <div className="set-head">
                  <div className="set-head-icon"><Icon name="envelope" size={18} /></div>
                  <div>
                    <h2>Importar conta</h2>
                    <div className="sub">
                      {wizStep === 'choose'
                        ? wizScanning
                          ? 'Etapa 1 de 2 — procurando contas…'
                          : `Etapa 1 de 2 — escolha a conta (${guesses.length} encontrada(s))`
                        : `Etapa 2 de 2 — ${imp?.g.email || ''}`}
                    </div>
                  </div>
                  <button className="set-close" onClick={() => !impBusy && setGuesses(null)}>
                    <Icon name="close" size={16} />
                  </button>
                </div>
                <div className="set-body" style={{ background: 'var(--surface)' }}>
                  {wizStep === 'choose' ? (
                    <>
                      <p className="set-hint" style={{ marginBottom: 4 }}>
                        Escolha a conta que quer importar e clique em <strong>Próximo</strong>. Uma por
                        vez — as demais ficam aqui para importar depois.
                      </p>
                      {wizNote && (
                        <p className="set-hint" style={{ marginBottom: 8 }}>{wizNote}</p>
                      )}
                      <div className="guess-rows" style={{ border: '1px solid var(--border)', borderRadius: 8 }}>
                        {wizScanning && (
                          <div className="guess-row" style={{ cursor: 'default' }}>
                            <span className="guess-main">
                              <span className="guess-name">Procurando contas…</span>
                              <span className="guess-meta">
                                Lendo o Registro do Windows / perfis — pode levar alguns segundos.
                              </span>
                            </span>
                          </div>
                        )}
                        {!wizScanning &&
                          guesses.map((g) => {
                          const done =
                            doneEmails.has(g.email.toLowerCase()) ||
                            accounts.some((a) => a.email.toLowerCase() === g.email.toLowerCase());
                          return (
                            <label
                              key={g.email}
                              className={`guess-row${done ? ' done' : ''}${wizPick === g.email ? ' pick' : ''}`}
                            >
                              <input
                                type="radio"
                                name="wiz-pick"
                                disabled={done}
                                checked={wizPick === g.email}
                                onChange={() => setWizPick(g.email)}
                              />
                              <span
                                className="guess-avatar"
                                style={{ background: `hsl(${avatarHue(g.email)} 50% 40%)` }}
                              >
                                {guessInitials(g)}
                              </span>
                              <span className="guess-main">
                                <span className="guess-name">
                                  {g.displayName}
                                  <span className="guess-email"> · {g.email}</span>
                                  {doneEmails.has(g.email.toLowerCase()) && (
                                    <span className="badge badge-local">importada</span>
                                  )}
                                  {!doneEmails.has(g.email.toLowerCase()) && done && (
                                    <span className="badge badge-off">já cadastrada</span>
                                  )}
                                  {g.password && !done && (
                                    <span className="badge badge-local">senha recuperada</span>
                                  )}
                                </span>
                                <span className="guess-meta">
                                  <span className="guess-chip">{g.protocol.toUpperCase()}</span>
                                  <span className="guess-chip">
                                    {g.imapHost}:{g.imapPort}
                                    {g.imapSecure ? ' · SSL' : ''}
                                  </span>
                                  <span className="guess-chip">SMTP {g.smtpHost}:{g.smtpPort}</span>
                                  {g.pstPath && (
                                    <span className="guess-chip">
                                      entrega em {g.pstPath.split(/[\\/]/).pop()}
                                    </span>
                                  )}
                                </span>
                              </span>
                            </label>
                          );
                        })}
                      </div>
                      <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
                        <button
                          className="btn btn-primary"
                          disabled={!wizPick}
                          onClick={() => {
                            const g = guesses.find((x) => x.email === wizPick);
                            if (g) openImport(g);
                          }}
                        >
                          Próximo →
                        </button>
                        <button className="btn" onClick={() => setGuesses(null)}>
                          Cancelar
                        </button>
                        {importAvailable && (
                          <button
                            className="btn btn-ghost"
                            title="Grava as contas encontradas num .json para levar ao Linux"
                            onClick={saveExport}
                          >
                            Salvar exportação para levar ao Linux
                          </button>
                        )}
                      </div>
                    </>
                  ) : (
                    imp && (
                      <>
                        <div className="guess-meta" style={{ marginBottom: 10 }}>
                          <span className="guess-chip">{imp.g.protocol.toUpperCase()}</span>
                          <span className="guess-chip">
                            {imp.g.imapHost}:{imp.g.imapPort}{imp.g.imapSecure ? ' · SSL' : ''}
                          </span>
                          <span className="guess-chip">SMTP {imp.g.smtpHost}:{imp.g.smtpPort}</span>
                          {imp.g.imapUser && imp.g.imapUser.toLowerCase() !== imp.g.email.toLowerCase() && (
                            <span className="guess-chip">usuário: {imp.g.imapUser}</span>
                          )}
                          {imp.g.smtpPassword && imp.g.smtpPassword !== imp.password && (
                            <span className="guess-chip">senha de SMTP separada recuperada</span>
                          )}
                          {imp.g.pstPath && (
                            <span className="guess-chip">
                              entrega em {imp.g.pstPath.split(/[\\/]/).pop()}
                            </span>
                          )}
                        </div>
                        <Field label="Nome de exibição">
                          <input
                            className="inp"
                            value={imp.label}
                            onChange={(e) => setImp({ ...imp, label: e.target.value })}
                          />
                        </Field>
                        <Field label="Senha">
                          <input
                            className="inp"
                            type="password"
                            value={imp.password}
                            placeholder={imp.g.password ? 'Senha recuperada do Windows' : 'Digite a senha da conta'}
                            onChange={(e) => setImp({ ...imp, password: e.target.value })}
                          />
                        </Field>
                        {imp.g.password && (
                          <span className="guess-hint">
                            A senha foi recuperada — revise ou troque, se necessário.
                          </span>
                        )}

                        <div className="grp-title" style={{ marginTop: 6 }}>Arquivo de dados</div>
                        <div className="acct-data-choice" role="radiogroup">
                          <label className={`acct-data-opt${!imp.newData ? ' on' : ''}`}>
                            <input type="radio" checked={!imp.newData} onChange={() => setImp({ ...imp, newData: false })} />
                            <span>
                              <strong>Usar um arquivo de dados existente</strong>
                              <em>As mensagens desta conta entram na Caixa de Entrada dele.</em>
                            </span>
                          </label>
                          <label className={`acct-data-opt${imp.newData ? ' on' : ''}`}>
                            <input type="radio" checked={imp.newData} onChange={() => setImp({ ...imp, newData: true })} />
                            <span>
                              <strong>Criar um novo arquivo de dados</strong>
                              <em>Base própria só para esta conta.</em>
                            </span>
                          </label>
                        </div>
                        {!imp.newData ? (
                          <Field label="Arquivo de dados">
                            <select
                              className="inp"
                              value={imp.path}
                              disabled={dataFiles.length === 0}
                              onChange={(e) => setImp({ ...imp, path: e.target.value })}
                            >
                              {dataFiles.length === 0 && <option value="">Nenhum arquivo de dados aberto</option>}
                              {dataFiles.map((s) => (
                                <option key={s.id} value={s.filePath}>
                                  {s.label}{defaultStore === s.filePath ? ' (principal)' : ''}
                                </option>
                              ))}
                            </select>
                          </Field>
                        ) : (
                          <Field label="Nome do novo arquivo de dados">
                            <input
                              className="inp"
                              value={imp.newName}
                              placeholder={`Emails — ${imp.label || imp.g.email}`}
                              onChange={(e) => setImp({ ...imp, newName: e.target.value })}
                            />
                          </Field>
                        )}

                        {imp.g.protocol === 'pop3' && (
                          <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12.5, marginTop: 4 }}>
                            <input
                              type="checkbox"
                              checked={imp.leave}
                              onChange={(e) => setImp({ ...imp, leave: e.target.checked })}
                            />
                            Deixar cópia das mensagens no servidor
                            {imp.leave && imp.leaveDays > 0 && ` por ${imp.leaveDays} dias`}
                          </label>
                        )}
                        {(imp.g.signatureHtml || imp.g.signatureReplyHtml) && (
                          <span className="guess-hint" style={{ display: 'block', marginTop: 6 }}>
                            Esta conta tem assinatura do Outlook — depois da importação ela aparece como
                            sugestão na aba Assinaturas.
                          </span>
                        )}
                        {impTest && (
                          <div
                            className="run-now-msg"
                            style={{ marginTop: 8, color: impTest.startsWith('✓') ? 'var(--ok, #107c41)' : '#a4262c' }}
                          >
                            {impTest}
                          </div>
                        )}
                        {impErr && <div className="rule-warn">{impErr}</div>}
                        <div style={{ display: 'flex', gap: 8, marginTop: 10, flexWrap: 'wrap' }}>
                          <button
                            className="btn"
                            disabled={impBusy}
                            onClick={() => setWizStep('choose')}
                          >
                            ← Voltar
                          </button>
                          <button className="btn" disabled={impBusy} onClick={() => void doTest()}>
                            {impBusy ? 'Testando…' : 'Testar conta'}
                          </button>
                          <button className="btn btn-primary" disabled={impBusy} onClick={() => void doImport()}>
                            {impBusy ? 'Importando…' : 'Importar esta conta'}
                          </button>
                          <button
                            className="btn"
                            title="Abrir o formulário completo para revisar todos os campos"
                            disabled={impBusy}
                            onClick={() => {
                              const g = imp.g;
                              setGuesses(null);
                              setImp(null);
                              applyGuess(g);
                            }}
                          >
                            Editar detalhes
                          </button>
                        </div>
                      </>
                    )
                  )}
                </div>
              </div>
            </div>
          )}

        </>
        )}

        {/* Assinaturas — lista na aba; a edição acontece num diálogo amplo */}
        {tab === 'sigs' && (
        <section className="set-card fill">
          <h3 className="set-card-title"><Icon name="pen" size={15} /> Assinaturas</h3>
          <p className="set-hint">
            Clique em <strong>Editar</strong> para abrir o editor e marque
            <strong> quais contas usam</strong> cada assinatura — a mesma assinatura
            pode ser aplicada a vários emails, em novas mensagens e/ou respostas.
          </p>
          {importedSigs.length > 0 && (
            <div className="sig-import">
              <div className="sig-import-title">Assinaturas encontradas na importação do Outlook</div>
              {importedSigs.map((s, i) => (
                <div key={i} className="acct-card sig-card import">
                  <div className="sig-avatar import"><Icon name="envelope" size={15} /></div>
                  <div className="acct-info">
                    <div className="acct-name">{s.name}</div>
                    <div className="acct-sub">Nenhuma conta usa ainda</div>
                  </div>
                  <button
                    className="btn"
                    onClick={() => {
                      setSigDraft({ name: s.name, html: s.html });
                      setSigAssign({});
                      setImportedSigs((prev) => prev.filter((_, j) => j !== i));
                    }}
                  >
                    Criar assinatura
                  </button>
                  <button
                    className="btn btn-ghost"
                    onClick={() => setImportedSigs((prev) => prev.filter((_, j) => j !== i))}
                  >
                    Descartar
                  </button>
                </div>
              ))}
            </div>
          )}
          <div className="sig-rows fill-scroll">
            {signatures.length === 0 && (
              <div className="empty-state">
                Nenhuma assinatura — clique em "Nova assinatura" para criar.
              </div>
            )}
            {signatures.map((s) => {
              const usedNew = accounts
                .filter((a) => a.signatureNewId === s.id)
                .map((a) => a.label || a.email);
              const usedReply = accounts
                .filter((a) => a.signatureReplyId === s.id)
                .map((a) => a.label || a.email);
              const used = [
                usedNew.length ? `novas: ${usedNew.join(', ')}` : '',
                usedReply.length ? `respostas: ${usedReply.join(', ')}` : '',
              ]
                .filter(Boolean)
                .join(' · ');
              return (
                <div key={s.id} className="acct-card sig-card">
                  <div className="sig-avatar"><Icon name="pen" size={15} /></div>
                  <div className="acct-info">
                    <div className="acct-name">
                      {s.name}
                      {usedNew.length > 0 && (
                        <span className="badge badge-pst">novas · {usedNew.length}</span>
                      )}
                      {usedReply.length > 0 && (
                        <span className="badge badge-local">respostas · {usedReply.length}</span>
                      )}
                    </div>
                    <div className="acct-sub" title={used}>
                      {used || 'Não usada por nenhuma conta'}
                    </div>
                  </div>
                  <button className="btn btn-ghost" onClick={() => openSignature(s)}>Editar</button>
                  <button
                    className="btn btn-ghost danger"
                    onClick={() =>
                      setConfirm({
                        title: 'Excluir assinatura',
                        message: `Excluir a assinatura "${s.name}"? Contas que a usam ficarão sem assinatura.`,
                        onOk: () => {
                          setConfirm(null);
                          void window.outlook.removeSignature(s.id).then(() => {
                            if (sigDraft?.id === s.id) setSigDraft(null);
                            reloadSignatures();
                          });
                        },
                      })
                    }
                  >
                    Excluir
                  </button>
                </div>
              );
            })}
          </div>
          <div style={{ paddingTop: 10 }}>
            <button
              className="btn btn-primary"
              onClick={() => {
                setSigDraft({ name: `Assinatura ${signatures.length + 1}`, html: '' });
                setSigAssign({});
              }}
            >
              <Icon name="new" size={14} /> Nova assinatura
            </button>
          </div>

        </section>
        )}

        {/* Editor de assinatura em diálogo — tela ampla, como no Outlook */}
        {tab === 'sigs' && sigDraft && (
          <div
            className="modal-overlay"
            onMouseDown={(e) => {
              if (e.target === e.currentTarget) setSigDraft(null);
            }}
          >
            <div className="sig-dlg">
              <div className="sig-dlg-head">
                <h2>{sigDraft.id ? 'Editar assinatura' : 'Nova assinatura'}</h2>
                <button className="sig-dlg-close" title="Fechar" onClick={() => setSigDraft(null)}>✕</button>
              </div>
              <div className="sig-dlg-body">
                <Field label="Nome da assinatura">
                  <input
                    className="inp"
                    value={sigDraft.name}
                    onChange={(e) => setSigDraft({ ...sigDraft, name: e.target.value })}
                  />
                </Field>
                {/* div.fld (não <label>): o clique numa imagem dentro do editor
                    não deve disparar o comportamento padrão de label */}
                <div className="fld sig-dlg-editor">
                  <span>Conteúdo</span>
                  <RichTextEditor
                    key={sigDraft.id || 'nova'}
                    value={sigDraft.html}
                    onChange={(html) => setSigDraft({ ...sigDraft, html })}
                    placeholder="Seu Nome — Cargo, Empresa, telefone…"
                  />
                </div>
                <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                  Para ajustar uma imagem, clique nela e use "Tamanho da imagem…"
                  (ou dê duplo clique) — largura e altura são editáveis no diálogo.
                </span>

                {/* Vínculo assinatura → contas: a mesma assinatura pode servir a
                    várias contas, em novas mensagens e/ou respostas. */}
                <div style={{ border: '1px solid var(--border)', borderRadius: 8, padding: '8px 10px' }}>
                  <div className="grp-title" style={{ marginTop: 0 }}>Contas que usam esta assinatura</div>
                  {accounts.length === 0 && (
                    <span style={{ fontSize: 12, color: 'var(--text-dim)' }}>
                      Nenhuma conta cadastrada — crie uma na aba Email.
                    </span>
                  )}
                  {accounts.length > 0 && (
                    <>
                      {accounts.map((a) => {
                        const asg = sigAssign[a.id] || { new: false, reply: false };
                        const cell: React.CSSProperties = {
                          display: 'flex',
                          flexDirection: 'column',
                          alignItems: 'center',
                          gap: 2,
                          minWidth: 110,
                        };
                        const lbl: React.CSSProperties = {
                          fontSize: 10.5,
                          color: 'var(--text-dim)',
                          textAlign: 'center',
                          lineHeight: 1.2,
                        };
                        return (
                          <div
                            key={a.id}
                            style={{
                              display: 'grid',
                              gridTemplateColumns: '1fr auto auto',
                              gap: '4px 16px',
                              alignItems: 'center',
                              padding: '4px 0',
                              borderTop: '1px dashed var(--border)',
                            }}
                          >
                            <span
                              style={{
                                fontSize: 12,
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                                whiteSpace: 'nowrap',
                              }}
                              title={a.email}
                            >
                              {a.label || a.email}
                              <span style={{ color: 'var(--text-dim)' }}> — {a.email}</span>
                            </span>
                            <label style={{ ...cell, cursor: 'pointer' }}>
                              <span style={lbl}>Novas mensagens</span>
                              <input
                                type="checkbox"
                                style={{ margin: 0 }}
                                checked={asg.new}
                                onChange={(e) =>
                                  setSigAssign({
                                    ...sigAssign,
                                    [a.id]: { ...asg, new: e.target.checked },
                                  })
                                }
                              />
                            </label>
                            <label style={{ ...cell, cursor: 'pointer' }}>
                              <span style={lbl}>Respostas /<br />encaminhamentos</span>
                              <input
                                type="checkbox"
                                style={{ margin: 0 }}
                                checked={asg.reply}
                                onChange={(e) =>
                                  setSigAssign({
                                    ...sigAssign,
                                    [a.id]: { ...asg, reply: e.target.checked },
                                  })
                                }
                              />
                            </label>
                          </div>
                        );
                      })}
                      <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                        Os vínculos são aplicados ao clicar em Salvar.
                      </span>
                    </>
                  )}
                </div>
              </div>
              <div className="sig-dlg-foot">
                <button className="btn" onClick={() => setSigDraft(null)}>Cancelar</button>
                <button
                  className="btn btn-primary"
                  disabled={!sigDraft.name.trim()}
                  onClick={saveSignature}
                >
                  Salvar
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Regras de recebimento */}
        {tab === 'rules' && (
        <section className="set-card fill">
          <h3 className="set-card-title"><Icon name="filter" size={15} /> Regras de recebimento</h3>
          <p className="set-hint">
            Quando chegar email novo, as regras são avaliadas em ordem — a primeira que combinar
            arquiva a mensagem na pasta escolhida. <strong>Regra sem condições combina com TODAS as mensagens</strong> —
            preencha pelo menos um campo "contém" para filtrar.
          </p>
          <div className="run-now">
            <div className="run-now-text">
              <div className="grp-title" style={{ marginTop: 0 }}>Executar regras</div>
              <p className="set-hint" style={{ margin: 0 }}>
                Aplica regras às mensagens de uma pasta existente, sem esperar um novo recebimento.
              </p>
            </div>
            <button
              className="btn btn-primary"
              disabled={accounts.length === 0}
              onClick={() => {
                setRunMsg('');
                setShowRun(true);
              }}
            >
              Executar regras
            </button>
          </div>
          {showRun && (
            <div className="modal-overlay" onMouseDown={(e) => e.target === e.currentTarget && setShowRun(false)}>
              <div className="app-dlg" role="dialog" aria-label="Executar regras">
                <div className="app-dlg-head">
                  <h2>Executar regras</h2>
                  <button className="app-dlg-close" title="Fechar" onClick={() => setShowRun(false)}>✕</button>
                </div>
                <div className="app-dlg-body">
                  <p className="set-hint" style={{ margin: 0 }}>
                    Escolha a conta, marque as regras e a pasta. Regras desativadas também podem ser executadas.
                  </p>
                  <Field label="Conta">
                    <select className="inp" value={runAccount} onChange={(e) => setRunAccount(e.target.value)}>
                      {accounts.map((a) => (
                        <option key={a.id} value={a.id}>{a.label || a.email} — {a.email}</option>
                      ))}
                    </select>
                  </Field>
                  <div className="run-now-rules">
                    {runAccountRules.length === 0 && (
                      <span className="run-now-empty">Esta conta não tem regras. Crie uma na lista abaixo.</span>
                    )}
                    {runAccountRules.map((r) => (
                      <label key={r.id} className="run-now-rule">
                        <input
                          type="checkbox"
                          checked={runRuleIds.includes(r.id)}
                          onChange={(e) =>
                            setRunRuleIds(
                              e.target.checked
                                ? [...runRuleIds, r.id]
                                : runRuleIds.filter((id) => id !== r.id),
                            )
                          }
                        />
                        <span className="run-now-rule-name">
                          {r.name || '(sem nome)'}
                          {!r.enabled && <em className="run-now-off"> — desativada</em>}
                        </span>
                        <span className="run-now-rule-target">
                          → {pstLabel(r.targetPstPath)} › {r.targetFolder}
                        </span>
                      </label>
                    ))}
                  </div>
                  <div className="run-now-grid">
                    <Field label="Arquivo de dados">
                      <div className="run-now-static">
                        {runSrcObj ? runSrcObj.label : '—'}
                      </div>
                      {runAccountObj?.deliveryPath && !runDeliverySrc && (
                        <span className="run-now-warn">
                          O arquivo de entrega desta conta não está aberto — usando o principal.
                        </span>
                      )}
                    </Field>
                    <Field label="Pasta">
                      <select className="inp" value={runFolder} onChange={(e) => setRunFolder(e.target.value)}>
                        {runFolders.map((f) => (
                          <option key={f.id} value={f.id}>{f.label}</option>
                        ))}
                      </select>
                    </Field>
                  </div>
                  {runSources.length === 0 && (
                    <div className="empty-state">Abra um arquivo de dados para executar as regras.</div>
                  )}
                  {runMsg && <div className="run-now-msg">{runMsg}</div>}
                </div>
                <div className="app-dlg-foot">
                  <button className="btn" onClick={() => setShowRun(false)}>Fechar</button>
                  <button
                    className="btn btn-primary"
                    disabled={runBusy || !runFolder || !runSource || runRuleIds.length === 0}
                    onClick={runNow}
                  >
                    {runBusy ? 'Executando…' : 'Executar agora'}
                  </button>
                </div>
              </div>
            </div>
          )}
          <div className="fill-scroll">
          {rules.length === 0 && (
            <div className="empty-state">Nenhuma regra cadastrada — mensagens novas vão para a pasta de entrega da conta.</div>
          )}
          {rules.map((r) => {
            const missing = accountMissing(r.accountId);
            const noConds = !r.conditions.fromContains && !r.conditions.toContains && !r.conditions.subjectContains;
            const isAuto = r.name?.startsWith('Entrega padrão');
            return (
              <div
                key={r.id}
                className={`rule-card${!r.enabled ? ' off' : ''}${missing ? ' broken' : ''}`}
              >
                <input
                  type="checkbox"
                  checked={r.enabled}
                  onChange={() => toggleRule(r)}
                  title={r.enabled ? 'Desativar regra' : 'Ativar regra'}
                />
                <span className="acct-info">
                  <span className="acct-name">
                    {r.name || '(sem nome)'}
                    {isAuto && <span className="badge badge-off">automática</span>}
                    {missing && (
                      <span className="badge badge-warn" title="A conta de email desta regra foi removida. Edite e escolha uma conta existente.">
                        conta ausente
                      </span>
                    )}
                  </span>
                  <span className="rule-desc">
                    em <em>{accountLabel(r.accountId)}</em>:
                    {noConds ? (
                      <span className="cond-chip cond-chip-any">
                        qualquer mensagem
                      </span>
                    ) : (
                      <>
                        {r.conditions.fromContains && <span className="cond-chip">De contém "{r.conditions.fromContains}"</span>}
                        {r.conditions.toContains && <span className="cond-chip">Para contém "{r.conditions.toContains}"</span>}
                        {r.conditions.subjectContains && <span className="cond-chip">Assunto contém "{r.conditions.subjectContains}"</span>}
                      </>
                    )}
                    {' → '}arquivar em <strong>{pstLabel(r.targetPstPath)}</strong> › {r.targetFolder}
                  </span>
                </span>
                <button onClick={() => editRule(r)} className="btn btn-ghost">Editar</button>
                <button
                  onClick={() => duplicateRule(r)}
                  className="btn btn-ghost"
                  title="Duplicar regra — abre o editor para escolher a conta de email"
                >
                  Duplicar
                </button>
                <button onClick={() => removeRule(r.id)} className="btn btn-ghost danger">Remover</button>
              </div>
            );
          })}
          </div>

          <button
            className="btn btn-primary"
            style={{ marginTop: 4 }}
            onClick={() => {
              setCustomFolder(false);
              setRuleDraft({ ...EMPTY_RULE });
            }}
          >
            <Icon name="new" size={14} /> Nova regra
          </button>
        </section>
        )}

        {/* Modal de edição de regra */}
        {ruleDraft && (
          <div className="modal-overlay" style={{ zIndex: 140 }} onMouseDown={(e) => { if (e.target === e.currentTarget) { setRuleDraft(null); setCustomFolder(false); } }}>
            <div
              className="set-card acct-modal"
              style={{ display: 'flex', flexDirection: 'column', gap: 12 }}
              onClick={(e) => e.stopPropagation()}
            >
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <strong style={{ fontSize: 14 }}>{ruleDraft.id ? 'Editar regra' : 'Nova regra'}</strong>
                <button className="set-close" onClick={() => { setRuleDraft(null); setCustomFolder(false); }} title="Fechar">
                  <Icon name="close" size={15} />
                </button>
              </div>
              <Field label="Nome da regra">
                <input className="inp" value={ruleDraft.name} onChange={(e) => setRuleDraft({ ...ruleDraft, name: e.target.value })} />
              </Field>
              <Field label="Aplicar quando chegar email em">
                <select
                  className="inp"
                  style={{
                    borderColor: ruleDraft.accountId !== '' && accountMissing(ruleDraft.accountId) ? '#c00' : 'var(--border)',
                  }}
                  value={ruleDraft.accountId}
                  onChange={(e) => setRuleDraft({ ...ruleDraft, accountId: e.target.value })}
                >
                  <option value="">— selecione a conta (obrigatório) —</option>
                  {accounts.map((a) => (
                    <option key={a.id} value={a.id}>{a.email}</option>
                  ))}
                  {ruleDraft.accountId !== '' && accountMissing(ruleDraft.accountId) && (
                    <option value={ruleDraft.accountId}>⚠ {ruleDraft.accountId} (conta removida)</option>
                  )}
                </select>
                {ruleDraft.accountId !== '' && accountMissing(ruleDraft.accountId) && (
                  <span style={{ fontSize: 11, color: '#a4262c' }}>
                    A conta desta regra foi removida — selecione uma conta existente para corrigir.
                  </span>
                )}
                {ruleDraft.accountId === '' && (
                  <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>É obrigatório escolher uma conta existente.</span>
                )}
              </Field>
              <div className="grp-title">Condições — preencha pelo menos uma para filtrar</div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
                <Field label="De contém (opcional)">
                  <input
                    className="inp"
                    value={ruleDraft.conditions.fromContains || ''}
                    onChange={(e) => setRuleDraft({ ...ruleDraft, conditions: { ...ruleDraft.conditions, fromContains: e.target.value } })}
                  />
                </Field>
                <Field label="Para contém (opcional)">
                  <input
                    className="inp"
                    value={ruleDraft.conditions.toContains || ''}
                    onChange={(e) => setRuleDraft({ ...ruleDraft, conditions: { ...ruleDraft.conditions, toContains: e.target.value } })}
                  />
                </Field>
                <Field label="Assunto contém (opcional)">
                  <input
                    className="inp"
                    value={ruleDraft.conditions.subjectContains || ''}
                    onChange={(e) => setRuleDraft({ ...ruleDraft, conditions: { ...ruleDraft.conditions, subjectContains: e.target.value } })}
                  />
                </Field>
              </div>
              {!ruleDraft.conditions.fromContains && !ruleDraft.conditions.toContains && !ruleDraft.conditions.subjectContains && (
                <div className="rule-warn">
                  ⚠ Sem condições, esta regra vale para <strong>todas as mensagens</strong> que chegarem na conta —
                  elas serão arquivadas na pasta abaixo em vez da entrega padrão.
                </div>
              )}
              <div className="grp-title">Destino — onde a mensagem é arquivada</div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                <Field label="Arquivar na base de dados">
                  <select
                    className="inp"
                    value={ruleDraft.targetPstPath}
                    onChange={(e) => {
                      setRuleDraft({ ...ruleDraft, targetPstPath: e.target.value, targetFolder: '' });
                      setCustomFolder(false);
                    }}
                  >
                    <option value="">— selecione a base —</option>
                    {uniqueSources.map((s) => (
                      <option key={s.id} value={s.filePath}>{s.label} ({s.filePath})</option>
                    ))}
                  </select>
                </Field>
                <Field label="Pasta de destino">
                  {folderChoices.length > 0 && !customFolder ? (
                    <select
                      className="inp"
                      value={ruleDraft.targetFolder}
                      onChange={(e) => {
                        if (e.target.value === '__new') {
                          setCustomFolder(true);
                          setRuleDraft({ ...ruleDraft, targetFolder: '' });
                        } else {
                          setRuleDraft({ ...ruleDraft, targetFolder: e.target.value });
                        }
                      }}
                    >
                      <option value="">— selecione a pasta —</option>
                      {folderChoices.map((f) => (
                        <option key={f.path} value={f.path}>{f.label}</option>
                      ))}
                      <option value="__new">＋ Nova pasta…</option>
                    </select>
                  ) : (
                    <input
                      className="inp"
                      placeholder={
                        ruleDraft.targetPstPath
                          ? 'Nome da pasta (será criada se não existir)'
                          : 'Selecione a base primeiro'
                      }
                      disabled={!ruleDraft.targetPstPath}
                      value={ruleDraft.targetFolder}
                      onChange={(e) => setRuleDraft({ ...ruleDraft, targetFolder: e.target.value })}
                    />
                  )}
                  {customFolder && folderChoices.length > 0 && (
                    <button
                      style={{ border: 'none', background: 'none', color: 'var(--blue)', cursor: 'pointer', fontSize: 11, textAlign: 'left' }}
                      onClick={() => {
                        setCustomFolder(false);
                        setRuleDraft({ ...ruleDraft, targetFolder: folderChoices[0]?.path || '' });
                      }}
                    >
                      ← escolher pasta existente
                    </button>
                  )}
                </Field>
              </div>
              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={!!ruleDraft.stopProcessing}
                  onChange={(e) => setRuleDraft({ ...ruleDraft, stopProcessing: e.target.checked })}
                />
                Parar de processar mais regras
              </label>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  onClick={saveRule}
                  disabled={
                    !ruleDraft.accountId ||
                    accountMissing(ruleDraft.accountId) ||
                    !ruleDraft.targetPstPath ||
                    !ruleDraft.targetFolder.trim()
                  }
                  className="btn btn-primary"
                >
                  Salvar regra
                </button>
                <button
                  onClick={() => {
                    setRuleDraft(null);
                    setCustomFolder(false);
                  }}
                  className="btn"
                >
                  Cancelar
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Geral — privacidade e exibição */}
        {tab === 'geral' && (
        <section className="set-card">
          <h3 className="set-card-title"><Icon name="settings" size={15} /> Geral</h3>
          <div className="grp-title">Aparência</div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6, fontSize: 12 }}>
            <span style={{ color: 'var(--text-dim)' }}>Tema da interface:</span>
            <select
              className="inp"
              style={{ width: 'auto' }}
              value={theme}
              onChange={(e) => {
                const t = e.target.value as 'system' | 'light' | 'dark';
                setThemeState(t);
                void window.outlook.setTheme?.(t).then(() => applyTheme(t));
              }}
            >
              <option value="system">Igual ao sistema</option>
              <option value="light">Claro</option>
              <option value="dark">Escuro</option>
            </select>
          </div>
          <div className="grp-title" style={{ marginTop: 14 }}>Envio e recebimento</div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 6, fontSize: 12 }}>
            <span style={{ color: 'var(--text-dim)' }}>Verificar novas mensagens a cada:</span>
            <select
              className="inp"
              style={{ width: 'auto' }}
              value={syncMin}
              onChange={(e) => {
                const v = Number(e.target.value);
                setSyncMin(v);
                void window.outlook.setSyncInterval?.(v);
              }}
            >
              <option value={0}>desligado</option>
              <option value={1}>1 minuto</option>
              <option value={5}>5 minutos</option>
              <option value={10}>10 minutos</option>
              <option value={15}>15 minutos</option>
              <option value={30}>30 minutos</option>
              <option value={60}>60 minutos</option>
            </select>
          </div>
          <div className="grp-title">Privacidade</div>
          <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: 12, cursor: 'pointer', maxWidth: 560 }}>
            <input
              type="checkbox"
              style={{ marginTop: 2 }}
              checked={loadImages}
              onChange={(e) => {
                setLoadImages(e.target.checked);
                void window.outlook.setLoadImages?.(e.target.checked);
              }}
            />
            <span>
              <strong>Baixar imagens externas automaticamente</strong>
              <br />
              <span style={{ color: 'var(--text-dim)' }}>
                Quando desativado (padrão, igual ao Outlook), imagens hospedadas na internet
                dentro de emails HTML ficam bloqueadas — remetentes não podem rastrear quando
                você abriu a mensagem. O HTML do email continua sendo interpretado.
              </span>
            </span>
          </label>
          <div className="grp-title" style={{ marginTop: 14 }}>Segurança</div>
          <label style={{ display: 'flex', alignItems: 'flex-start', gap: 8, fontSize: 12, cursor: 'pointer', maxWidth: 560 }}>
            <input
              type="checkbox"
              style={{ marginTop: 2 }}
              checked={virusScan}
              onChange={(e) => {
                setVirusScan(e.target.checked);
                void window.outlook.setVirusScan?.(e.target.checked);
              }}
            />
            <span>
              <strong>Verificar vírus nas mensagens recebidas</strong>
              <br />
              <span style={{ color: 'var(--text-dim)' }}>
                Cada mensagem baixada é verificada pelo antivírus do computador
                {avEngine ? ` (${avEngine})` : ''}. Mensagens infectadas vão direto para a
                pasta <b>Itens infectados</b> da base de dados, sem passar pela Caixa de Entrada.
              </span>
            </span>
          </label>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginTop: 18 }}>
            <button
              className="btn btn-primary"
              onClick={() => {
                // Os campos já persistem na troca — o botão regrava tudo e
                // confirma, como o "OK" das Opções do Outlook.
                void window.outlook.setTheme?.(theme).then(() => applyTheme(theme));
                void window.outlook.setSyncInterval?.(syncMin);
                void window.outlook.setLoadImages?.(loadImages);
                void window.outlook.setVirusScan?.(virusScan);
                setGeralSaved(true);
                setTimeout(() => setGeralSaved(false), 4000);
              }}
            >
              <Icon name="check" size={14} /> Salvar
            </button>
            {geralSaved && (
              <span style={{ fontSize: 12, color: 'var(--ok)' }}>Configurações salvas.</span>
            )}
          </div>
        </section>
        )}

        </div>
      </div>
      {confirm && (
        <ConfirmDialog
          title={confirm.title}
          message={confirm.message}
          onConfirm={confirm.onOk}
          onCancel={() => setConfirm(null)}
        />
      )}
      {/* Modal: novo arquivo de dados */}
      {showNewStore && (
        <div className="set-overlay" style={{ zIndex: 110 }} onMouseDown={(e) => e.target === e.currentTarget && setShowNewStore(false)}>
          <div className="set-window" style={{ width: 480, height: 'auto' }} onClick={(e) => e.stopPropagation()}>
            <div className="set-head">
              <div className="set-head-icon"><Icon name="new" size={18} /></div>
              <div>
                <h2>Novo arquivo de dados</h2>
                <div className="sub">Base própria do Unit Mail — mensagens em .eml</div>
              </div>
              <button className="set-close" onClick={() => setShowNewStore(false)}>
                <Icon name="close" size={16} />
              </button>
            </div>
            <div className="set-body" style={{ background: 'var(--surface)' }}>
              <Field label="Nome do arquivo de dados">
                <input
                  className="inp"
                  placeholder="Ex.: Arquivo 2024, Emails — Trabalho…"
                  value={newStoreName}
                  autoFocus
                  onChange={(e) => setNewStoreName(e.target.value)}
                />
              </Field>
              <Field label="Vincular a uma conta (opcional)">
                <select
                  className="inp"
                  value={newStoreAccount}
                  onChange={(e) => setNewStoreAccount(e.target.value)}
                >
                  <option value="">Sem vínculo — qualquer conta pode usar</option>
                  {accounts.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.label || a.email} — {a.email}
                      {a.deliveryPath ? ' (já tem base)' : ''}
                    </option>
                  ))}
                </select>
              </Field>
              <p className="set-hint">
                Sem vínculo, a base fica disponível para receber mensagens de várias contas —
                todas caem na mesma Caixa de Entrada (use Regras para separar em pastas).
                Com vínculo, só as mensagens daquela conta são entregues aqui.
              </p>
              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  className="btn btn-primary"
                  disabled={!newStoreName.trim()}
                  onClick={() => {
                    setShowNewStore(false);
                    void createStore(false);
                  }}
                >
                  Criar na pasta de dados
                </button>
                <button
                  className="btn"
                  disabled={!newStoreName.trim()}
                  onClick={() => {
                    setShowNewStore(false);
                    void createStore(true);
                  }}
                >
                  Criar em outra pasta…
                </button>
                <button className="btn" onClick={() => setShowNewStore(false)}>Cancelar</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {migrateSrc && (
        <div className="set-overlay" style={{ zIndex: 110 }} onMouseDown={(e) => e.target === e.currentTarget && !migrateProgress && setMigrateSrc(null)}>
          <div className="set-window" style={{ width: 480, height: 'auto' }} onClick={(e) => e.stopPropagation()}>
            <div className="set-head">
              <div className="set-head-icon"><Icon name="archive" size={18} /></div>
              <div>
                <h2>{migrateCloseAfter ? 'Importar arquivo .pst' : 'Migrar para base própria'}</h2>
                <div className="sub">{migrateSrc.label}</div>
              </div>
              {!migrateProgress && (
                <button
                  className="set-close"
                  onClick={() => {
                    // Cancelar a importação: fecha o PST temporário que só foi
                    // aberto para esta conversão — não deixa resíduo na lista.
                    if (migrateCloseAfter && !migrateDone) void onCloseSource(migrateSrc.id);
                    setMigrateSrc(null);
                  }}
                >
                  <Icon name="close" size={16} />
                </button>
              )}
            </div>
            <div className="set-body" style={{ background: 'var(--surface)' }}>
              {!migrateProgress && !migrateDone && (
                <>
                  <p className="set-hint">
                    Copia TODAS as pastas e mensagens do PST para uma base nova, 100% do Unit Mail —
                    sem a limitação de "somente leitura" do formato .pst. Nessa base, excluir apaga o
                    arquivo de verdade e libera espaço em disco de imediato.
                  </p>
                  <div className="rule-warn" style={{ marginBottom: 4 }}>
                    ⚠ O PST original <strong>não é alterado nem apagado</strong> —
                    {migrateCloseAfter
                      ? ' só foi aberto para ler os dados e será fechado ao final; o arquivo continua no disco igual estava.'
                      : ' continua intacto e aberto no app.'}{' '}
                    Itens de calendário/contatos/tarefas são convertidos para um registro
                    com assunto e corpo (o formato ICS/vCard original não é preservado).
                  </div>
                  <div className="rule-warn" style={{ marginBottom: 4 }}>
                    ⏱ <strong>PSTs grandes podem demorar</strong> — um arquivo de vários GB leva
                    bastante tempo para copiar item a item. O progresso é mostrado durante a
                    importação; não feche o app enquanto ela roda.
                  </div>
                  <Field label="Importar para">
                    <select
                      className="inp"
                      value={migrateTarget}
                      onChange={(e) => setMigrateTarget(e.target.value)}
                    >
                      <option value="">Nova base de dados do Unit Mail</option>
                      {uniqueSources
                        .filter((s) => s.kind === 'local' && s.id !== migrateSrc.id)
                        .map((s) => (
                          <option key={s.id} value={s.filePath}>
                            {s.label} — base existente
                          </option>
                        ))}
                    </select>
                  </Field>
                  {migrateTarget ? (
                    <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                      As pastas e mensagens do PST serão mescladas dentro dessa base —
                      itens com nomes de pasta iguais se juntam.
                    </span>
                  ) : (
                    <>
                      <Field label="Nome da nova base">
                        <input className="inp" value={migrateName} onChange={(e) => setMigrateName(e.target.value)} />
                      </Field>
                      <Field label="Pasta onde criar">
                        <div style={{ display: 'flex', gap: 8 }}>
                          <input
                            className="inp"
                            style={{ flex: 1 }}
                            readOnly
                            value={migrateDir || 'Pasta de dados padrão (Mail/) — escolher outra ao lado'}
                          />
                          <button
                            className="btn"
                            onClick={async () => {
                              const d = await window.outlook.pickStoreDir();
                              if (d) setMigrateDir(d);
                            }}
                          >
                            Escolher…
                          </button>
                        </div>
                      </Field>
                    </>
                  )}
                  {migrateError && <div className="rule-warn">{migrateError}</div>}
                  <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
                    <button
                      className="btn btn-primary"
                      disabled={!migrateTarget && !migrateName.trim()}
                      onClick={runMigration}
                    >
                      {migrateCloseAfter ? 'Importar' : 'Iniciar migração'}
                    </button>
                    <button
                      className="btn"
                      onClick={() => {
                        if (migrateCloseAfter) void onCloseSource(migrateSrc.id);
                        setMigrateSrc(null);
                      }}
                    >
                      Cancelar
                    </button>
                  </div>
                </>
              )}
              {migrateProgress && migrateProgress.phase !== 'done' && (
                <>
                  <p className="set-hint" style={{ marginBottom: 10 }}>
                    Migrando "{migrateProgress.currentFolder || migrateSrc.label}"… não feche o app.
                  </p>
                  <div style={{ height: 10, borderRadius: 6, background: 'var(--surface-mid)', overflow: 'hidden', border: '1px solid var(--border)' }}>
                    <div
                      style={{
                        height: '100%',
                        background: 'var(--blue)',
                        width: `${migrateProgress.total ? Math.min(100, (migrateProgress.done / migrateProgress.total) * 100) : 0}%`,
                        transition: 'width 0.2s',
                      }}
                    />
                  </div>
                  <div style={{ fontSize: 11, color: 'var(--text-dim)', marginTop: 6 }}>
                    {migrateProgress.done} / {migrateProgress.total} itens
                  </div>
                </>
              )}
              {migrateDone && (
                <>
                  <p className="set-hint">
                    ✓ {migrateCloseAfter ? 'Importação' : 'Migração'} concluída — base <strong>"{migrateDone}"</strong> criada
                    com todo o conteúdo.{' '}
                    {migrateCloseAfter
                      ? 'O arquivo .pst original não foi alterado — só não fica mais aberto na lista.'
                      : 'O PST original continua aberto e intocado; feche-o quando quiser (aba "Arquivos de Dados" → Fechar).'}{' '}
                    Defina a nova base como principal, se desejar.
                  </p>
                  <button className="btn btn-primary" onClick={() => setMigrateSrc(null)}>Concluir</button>
                </>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Importar do Thunderbird — perfis com arquivos .mbox */}
      {tbProfiles !== null && (
        <div
          className="set-overlay"
          style={{ zIndex: 110 }}
          onMouseDown={(e) => e.target === e.currentTarget && !migrateProgress && setTbProfiles(null)}
        >
          <div className="set-window" style={{ width: 480, height: 'auto' }} onClick={(e) => e.stopPropagation()}>
            <div className="set-head">
              <div className="set-head-icon"><Icon name="download" size={18} /></div>
              <div>
                <h2>Importar do Thunderbird</h2>
                <div className="sub">mensagens (.mbox) do perfil instalado neste computador</div>
              </div>
              {!migrateProgress && (
                <button className="set-close" onClick={() => setTbProfiles(null)}>
                  <Icon name="close" size={16} />
                </button>
              )}
            </div>
            <div className="set-body" style={{ background: 'var(--surface)' }}>
              {tbProfiles.length === 0 ? (
                <>
                  <p className="set-hint">
                    Nenhum perfil do Thunderbird com mensagens foi encontrado neste computador.
                  </p>
                  {tbError && <div className="rule-warn">{tbError}</div>}
                  <button className="btn" onClick={() => setTbProfiles(null)}>Fechar</button>
                </>
              ) : !migrateProgress && !tbDone ? (
                (() => {
                  const prof = tbProfiles.find((p) => p.path === tbSrc);
                  return (
                    <>
                      <p className="set-hint">
                        Copia TODAS as pastas e mensagens do perfil para uma base do Unit Mail —
                        o perfil do Thunderbird não é alterado. Perfis grandes podem demorar;
                        não feche o app durante a importação.
                      </p>
                      {tbProfiles.length > 1 && (
                        <Field label="Perfil do Thunderbird">
                          <select
                            className="inp"
                            value={tbSrc}
                            onChange={(e) => {
                              setTbSrc(e.target.value);
                              const p = tbProfiles.find((x) => x.path === e.target.value);
                              if (p) setTbName(`Thunderbird — ${p.name}`);
                            }}
                          >
                            {tbProfiles.map((p) => (
                              <option key={p.path} value={p.path}>
                                {p.name} — {p.path}
                              </option>
                            ))}
                          </select>
                        </Field>
                      )}
                      {prof && (
                        <div className="rule-warn" style={{ marginBottom: 4 }}>
                          Encontrado: {prof.stores.map((s) => s.label).join(', ')}.
                        </div>
                      )}
                      <Field label="Importar para">
                        <select className="inp" value={tbTarget} onChange={(e) => setTbTarget(e.target.value)}>
                          <option value="">Nova base de dados do Unit Mail</option>
                          {uniqueSources
                            .filter((s) => s.kind === 'local')
                            .map((s) => (
                              <option key={s.id} value={s.filePath}>
                                {s.label} — base existente
                              </option>
                            ))}
                        </select>
                      </Field>
                      {tbTarget ? (
                        <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                          As pastas e mensagens do perfil serão mescladas dentro dessa base —
                          itens com nomes de pasta iguais se juntam.
                        </span>
                      ) : (
                        <>
                          <Field label="Nome da nova base">
                            <input className="inp" value={tbName} onChange={(e) => setTbName(e.target.value)} />
                          </Field>
                          <Field label="Pasta onde criar">
                            <div style={{ display: 'flex', gap: 8 }}>
                              <input
                                className="inp"
                                style={{ flex: 1 }}
                                readOnly
                                value={tbDir || 'Pasta de dados padrão (Mail/) — escolher outra ao lado'}
                              />
                              <button
                                className="btn"
                                onClick={async () => {
                                  const d = await window.outlook.pickStoreDir();
                                  if (d) setTbDir(d);
                                }}
                              >
                                Escolher…
                              </button>
                            </div>
                          </Field>
                        </>
                      )}
                      {tbError && <div className="rule-warn">{tbError}</div>}
                      <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
                        <button
                          className="btn btn-primary"
                          disabled={!tbSrc || (!tbTarget && !tbName.trim())}
                          onClick={runTbImport}
                        >
                          Importar
                        </button>
                        <button className="btn" onClick={() => setTbProfiles(null)}>
                          Cancelar
                        </button>
                      </div>
                    </>
                  );
                })()
              ) : null}
              {migrateProgress && migrateProgress.phase !== 'done' && (
                <>
                  <p className="set-hint" style={{ marginBottom: 10 }}>
                    Importando "{migrateProgress.currentFolder || 'Thunderbird'}"… não feche o app.
                  </p>
                  <div style={{ height: 10, borderRadius: 6, background: 'var(--surface-mid)', overflow: 'hidden', border: '1px solid var(--border)' }}>
                    <div
                      style={{
                        height: '100%',
                        background: 'var(--blue)',
                        width: `${migrateProgress.total ? Math.min(100, (migrateProgress.done / migrateProgress.total) * 100) : 0}%`,
                        transition: 'width 0.2s',
                      }}
                    />
                  </div>
                  <div style={{ fontSize: 11, color: 'var(--text-dim)', marginTop: 6 }}>
                    {migrateProgress.done} / {migrateProgress.total} itens
                  </div>
                </>
              )}
              {tbDone && (
                <>
                  <p className="set-hint">
                    ✓ Importação concluída — base <strong>"{tbDone}"</strong> criada com todo o
                    conteúdo do perfil. O Thunderbird original não foi alterado.
                  </p>
                  <button className="btn btn-primary" onClick={() => setTbProfiles(null)}>Concluir</button>
                </>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function guessInitials(g: ImportedAccountGuess): string {
  const name = (g.displayName || g.email).replace(/<[^>]*>/g, '').trim();
  const parts = name.split(/\s+/).filter(Boolean);
  if (parts.length > 1) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  return (parts[0]?.[0] || '?').toUpperCase();
}

function avatarHue(name: string): number {
  let h = 0;
  for (const c of name) h = (h * 31 + c.charCodeAt(0)) % 360;
  return h;
}
