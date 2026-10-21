import React, { useEffect, useRef, useState } from 'react';
import { Icon } from '../icons';
import { guessServers } from '../utils';
import { getLang, setLang, t, useLang } from '../i18n';
import type { AccountDraft, ImportedAccountGuess, MailSource } from '../../../shared/types';

interface OnboardingWizardProps {
  /** Recarrega fontes/contas na tela principal */
  onChanged: () => void;
  /** Marca a jornada como concluída e fecha */
  onDone: () => void;
}

type Step = 'datadir' | 'choice' | 'import' | 'account' | 'done';

const inputStyle: React.CSSProperties = {
  padding: '8px 10px',
  border: '1px solid var(--border)',
  borderRadius: 6,
  fontFamily: 'inherit',
  fontSize: 13,
  width: '100%',
};

function guessProtocol(email: string): { protocol: 'pop3' | 'imap'; host: string; port: number; secure: boolean; smtpHost: string; smtpPort: number; smtpSecure: boolean } {
  const g = guessServers(email);
  const domain = email.split('@')[1]?.toLowerCase() || '';
  // Presets conhecidos POP3; fallback IMAP
  if (domain === 'gmail.com')
    return { protocol: 'pop3', host: 'pop.gmail.com', port: 995, secure: true, smtpHost: 'smtp.gmail.com', smtpPort: 465, smtpSecure: true };
  if (['outlook.com', 'hotmail.com', 'live.com'].includes(domain))
    return { protocol: 'imap', host: 'outlook.office365.com', port: 993, secure: true, smtpHost: 'smtp.office365.com', smtpPort: 587, smtpSecure: false };
  return { protocol: 'pop3', host: `mail.${domain}`, port: 995, secure: true, smtpHost: g.smtpHost.replace('smtp.', 'mail.'), smtpPort: 587, smtpSecure: false };
}

/** Dois últimos segmentos do caminho (ex.: 'Documents\Unit Mail') — rótulo
 *  amigável que não expõe o caminho completo da máquina na tela. */
function shortPathLabel(p: string): string {
  const segs = p.split(/[\\/]+/).filter(Boolean);
  return segs.slice(-2).join('\\') || p;
}

/** Rascunho do wizard salvo no localStorage — se o app fechar no meio da
 *  configuração, ao reabrir ele volta na mesma etapa com os campos preenchidos.
 *  Senha NÃO é persistida (segurança) — é digitada de novo. */
const DRAFT_KEY = 'unitmail-wizard-draft';

interface WizardDraft {
  step: Step;
  email: string;
  label: string;
  /** 'new' | 'newCustom' | 'none' | 'existing:<caminho da base>' */
  storeChoice: string;
}

function loadDraft(): WizardDraft | null {
  try {
    const raw = localStorage.getItem(DRAFT_KEY);
    return raw ? (JSON.parse(raw) as WizardDraft) : null;
  } catch {
    return null;
  }
}

export function OnboardingWizard({ onChanged, onDone }: OnboardingWizardProps) {
  useLang();
  const draft0 = useRef(loadDraft()).current;
  const [step, setStep] = useState<Step>(draft0?.step || 'datadir');
  const [importAvailable, setImportAvailable] = useState(false);
  const [guesses, setGuesses] = useState<ImportedAccountGuess[]>([]);
  const [addedEmails, setAddedEmails] = useState<Set<string>>(new Set());
  const [dataDirPath, setDataDirPath] = useState('');
  const [dirMsg, setDirMsg] = useState('');
  const [profileFound, setProfileFound] = useState(false);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState('');

  // Mini-formulário de conta
  const [email, setEmail] = useState(draft0?.email || '');
  const [label, setLabel] = useState(draft0?.label || '');
  const [password, setPassword] = useState('');
  const [storeChoice, setStoreChoice] = useState<string>(draft0?.storeChoice || 'new');
  /** Bases locais já abertas — oferecidas como destino existente na conta. */
  const [localStores, setLocalStores] = useState<MailSource[]>([]);

  useEffect(() => {
    window.outlook.isOutlookImportAvailable().then(setImportAvailable);
    window.outlook.getDataDir?.().then((v) => setDataDirPath(v || ''));
    window.outlook
      .listSources()
      .then((list) => setLocalStores(list.filter((s) => s.kind === 'local' && s.filePath)));
  }, []);

  // Persiste o rascunho a cada mudança de etapa/campo relevante
  useEffect(() => {
    try {
      localStorage.setItem(
        DRAFT_KEY,
        JSON.stringify({ step, email, label, storeChoice } satisfies WizardDraft),
      );
    } catch {
      // storage cheio/bloqueado — não crítico
    }
  }, [step, email, label, storeChoice]);

  // Pasta escolhida mas ainda não aplicada — só grava ao confirmar ("Continuar")
  const [pendingDir, setPendingDir] = useState('');
  const [pendingIsProfile, setPendingIsProfile] = useState(false);
  // Pasta que JÁ foi aplicada neste wizard — ao re-selecioná-la não pode ser
  // tratada como "perfil encontrado" (o marcador é nosso, recém-criado)
  const [appliedDir, setAppliedDir] = useState('');

  const pickDataDir = async () => {
    const res = await window.outlook.pickDataDirPath?.();
    if (!res) return;
    const same = (a: string, b: string) => a && b && a.toLowerCase() === b.toLowerCase();
    const isOwn = same(res.dir, appliedDir) || same(res.dir, dataDirPath);
    setPendingDir(res.dir);
    setPendingIsProfile(res.isProfile && !isOwn);
    // Trocou de pasta → sai do modo "perfil encontrado"
    setProfileFound(false);
    setDirMsg(
      isOwn
        ? t('✓ Esta já é a pasta de dados atual — pode continuar.')
        : res.isProfile
          ? t('⚠ Esta pasta já contém um perfil do Unit Mail — restaure ou escolha outra pasta.')
          : t('Pasta selecionada — ainda nada foi gravado. Confirme em "Continuar".'),
    );
  };

  // Aplica a escolha de pasta (só no "Continuar") e avança o wizard.
  // Pasta com perfil NÃO é aplicada aqui — exige decisão explícita.
  const confirmDataDir = async () => {
    if (pendingIsProfile) {
      // Trava o passo: ou restaura, ou escolhe outra pasta
      setProfileFound(true);
      return;
    }
    // Já é a pasta aplicada/atual — nada a fazer, só avança
    const same = (a: string, b: string) => !!a && !!b && a.toLowerCase() === b.toLowerCase();
    if (!pendingDir || same(pendingDir, appliedDir) || same(pendingDir, dataDirPath)) {
      setStep('choice');
      return;
    }
    setBusy(true);
    try {
      const res = await window.outlook.applyDataDir?.(pendingDir);
      if (res?.adopted) {
        setProfileFound(true);
        return;
      }
      if (res) {
        setDataDirPath(res.dir);
        setAppliedDir(res.dir);
      }
      setStep('choice');
    } finally {
      setBusy(false);
    }
  };

  // Restauração: só agora o ponteiro é gravado e o app reinicia
  const restoreProfile = async () => {
    setBusy(true);
    setDirMsg(t('✓ Restaurando perfil — o app vai reiniciar…'));
    try {
      await window.outlook.applyDataDir?.(pendingDir);
      await window.outlook.restartApp?.();
    } finally {
      setBusy(false);
    }
  };

  const scan = async (src: 'outlook' | 'thunderbird') => {
    const name = src === 'outlook' ? 'Outlook' : 'Thunderbird';
    setBusy(true);
    setMsg(t('Procurando contas do {name}…', { name }));
    try {
      const found =
        src === 'outlook' ? await window.outlook.scanOutlookAccounts() : await window.outlook.scanThunderbirdAccounts();
      setGuesses(found);
      setSelectedEmails(new Set(found.map((g) => g.email)));
      setMsg(found.length ? '' : t('Nenhuma conta encontrada no {name}.', { name }));
    } catch {
      setMsg(t('Não foi possível ler as contas do {name} neste sistema.', { name }));
    } finally {
      setBusy(false);
    }
  };

  const openExport = async () => {
    setBusy(true);
    try {
      const found = await window.outlook.openOutlookExportFile().catch(() => null);
      if (found) {
        setGuesses(found);
        setSelectedEmails(new Set(found.map((g) => g.email)));
      }
      setMsg(found ? '' : t('Nenhum arquivo selecionado.'));
    } finally {
      setBusy(false);
    }
  };

  const [selectedEmails, setSelectedEmails] = useState<Set<string>>(new Set());

  const addGuess = async (g: ImportedAccountGuess) => {
    // Conta POP3 importada ganha sua própria base de dados em Mail/ —
    // modelo Thunderbird, e igual ao que o app faz ao criar conta manual.
    // IMAP não: as mensagens ficam no servidor, sem arquivo de dados local.
    let deliveryPath: string | undefined;
    if (g.protocol !== 'imap') {
      try {
        const src = await window.outlook.createLocalStoreAuto(t('Emails — {name}', { name: g.email }));
        deliveryPath = src?.filePath;
      } catch {
        // sem base — a conta ainda entra, vincula depois
      }
    }
    const draft: AccountDraft = {
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
      signature: g.signatureHtml || '',
      leaveOnServer: g.leaveOnServer ?? true,
      leaveDays: g.leaveDays ?? 0,
      // Senha recuperada do Registro via DPAPI (só nesta máquina/sessão)
      password: g.password || '',
      smtpPassword: g.smtpPassword && g.smtpPassword !== g.password ? g.smtpPassword : undefined,
      deliveryPath,
    };
    await window.outlook.upsertAccount(draft);
    setAddedEmails((s) => new Set(s).add(g.email));
    onChanged();
  };

  const addSelected = async () => {
    setBusy(true);
    setMsg('');
    try {
      for (const g of guesses) {
        if (selectedEmails.has(g.email) && !addedEmails.has(g.email)) await addGuess(g);
      }
      setStep('done');
    } finally {
      setBusy(false);
    }
  };

  const createAccount = async () => {
    if (!email.includes('@')) {
      setMsg(t('Informe um email válido.'));
      return;
    }
    setBusy(true);
    setMsg('');
    try {
      const g = guessProtocol(email);
      // IMAP mantém tudo no servidor — não cria nem vincula arquivo de dados.
      let deliveryPath: string | undefined;
      if (g.protocol !== 'imap') {
        if (storeChoice.startsWith('existing:')) {
          deliveryPath = storeChoice.slice('existing:'.length);
        } else if (storeChoice === 'new') {
          // Local padrão: pasta Mail/ da pasta de dados
          const src = await window.outlook.createLocalStoreAuto(t('Emails — {name}', { name: label || email }));
          if (src?.filePath) {
            deliveryPath = src.filePath;
          }
        } else if (storeChoice === 'newCustom') {
          // Usuário escolhe a pasta onde o arquivo de dados será criado
          const dir = await window.outlook.pickStoreDir();
          if (!dir) {
            setMsg(t('Seleção de pasta cancelada — a conta não foi criada.'));
            return;
          }
          const src = await window.outlook.createLocalStoreIn(t('Emails — {name}', { name: label || email }), dir);
          if (src?.filePath) deliveryPath = src.filePath;
        }
      }
      const draft: AccountDraft = {
        label: label || email,
        email,
        protocol: g.protocol,
        imapHost: g.host,
        imapPort: g.port,
        imapSecure: g.secure,
        imapUser: email,
        smtpHost: g.smtpHost,
        smtpPort: g.smtpPort,
        smtpSecure: g.smtpSecure,
        smtpUser: email,
        signature: '',
        leaveOnServer: true,
        leaveDays: 0,
        password,
        deliveryPath,
      };
      const saved = await window.outlook.upsertAccount(draft);
      if (password) {
        try {
          await window.outlook.connectAccount(saved.id);
        } catch (e) {
          setMsg(t('Conta criada, mas a conexão falhou: {msg}. Confira a senha nas Configurações.', { msg: (e as Error).message }));
        }
      }
      onChanged();
      setStep('done');
    } catch (e) {
      setMsg(t('Erro ao criar conta: {msg}', { msg: (e as Error).message }));
    } finally {
      setBusy(false);
    }
  };

  const finish = () => {
    try {
      localStorage.removeItem(DRAFT_KEY);
    } catch {
      // ignora
    }
    void window.outlook.setSetupComplete(true);
    // Nenhum arquivo de dados é criado aqui — base só nasce na criação de
    // conta ou em Configurações → Arquivos de Dados. Quem pulou pode abrir
    // um PST/criar tudo depois.
    onChanged();
    onDone();
  };

  return (
    <div className="modal-overlay" style={{ zIndex: 200 }}>
      <div className="wizard-window">
        <div className="wizard-head">
          <Icon name="logo" size={40} />
          <div style={{ flex: 1 }}>
            <h1>{t('Bem-vindo ao Unit Mail')}</h1>
            <p>{t('Configure seu email em poucos passos')}</p>
          </div>
          <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: 'var(--text-dim)' }}>
            <Icon name="settings" size={14} />
            <select
              className="inp"
              style={{ width: 'auto', padding: '4px 8px', fontSize: 12 }}
              value={getLang()}
              onChange={(e) => setLang(e.target.value as 'pt' | 'en')}
            >
              <option value="pt">Português</option>
              <option value="en">English</option>
            </select>
          </label>
        </div>

        <div className="wizard-steps">
          {(['datadir', 'choice', 'import', 'account', 'done'] as Step[]).map((s, i) => {
            const labels: Record<Step, string> = {
              datadir: t('Dados'),
              choice: t('Início'),
              import: t('Importar'),
              account: t('Conta'),
              done: t('Pronto'),
            };
            const order: Step[] = ['datadir', 'choice', 'import', 'account', 'done'];
            const active = order.indexOf(step) >= i;
            const isCurrent = step === s;
            return (
              <span key={s} className={`wz-step ${isCurrent ? 'current' : ''} ${active ? 'active' : ''}`}>
                {labels[s]}
              </span>
            );
          })}
        </div>

        {step === 'datadir' && (
          <div className="wizard-body">
            <h2>{t('Onde guardar seus dados?')}</h2>
            <p className="wz-hint">
              {t('O Unit Mail guarda')} <strong>{t('tudo')}</strong>{' '}
              {t('numa única pasta — contas, regras, assinaturas e os emails, dentro de')}{' '}
              <code>Mail/</code> {t('(como o perfil do Thunderbird). Para backup ou mudança de computador, basta copiar essa pasta.')}
            </p>
            <p className="wz-hint" style={{ marginTop: -8 }}>
              {t('Voltando de outro PC?')} <strong>{t('Escolha a pasta copiada')}</strong> —{' '}
              {t('o app reconhece o perfil e restaura tudo sozinho.')}
            </p>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input
                style={{ ...inputStyle, color: 'var(--text-dim)', flex: 1 }}
                readOnly
                // Sem pasta escolhida ainda, mostra um rótulo amigável do
                // padrão (sem expor o caminho completo da máquina, ex.: o
                // nome do usuário em C:\Users\…).
                value={pendingDir || t('Pasta padrão — {label}', { label: shortPathLabel(dataDirPath) })}
                title={pendingDir || dataDirPath}
              />
              <button
                className="btn"
                style={{ whiteSpace: 'nowrap', flexShrink: 0 }}
                disabled={busy}
                onClick={pickDataDir}
              >
                <Icon name="folder" size={14} /> {t('Escolher pasta…')}
              </button>
            </div>
            {pendingDir && (
              <p style={{ fontSize: 11, marginTop: 4, color: 'var(--text-dim)' }}>
                {pendingIsProfile ? t('Perfil encontrado nesta pasta.') : t('Pasta nova — seus dados serão criados aqui.')}
                {' '}{t('Pode escolher outra antes de continuar.')}
              </p>
            )}
            {dirMsg && (
              <p style={{ fontSize: 12, marginTop: 8, color: '#107c41' }}>{dirMsg}</p>
            )}
            <div className="wizard-nav">
              {profileFound ? (
                <button className="primary-btn" disabled={busy} onClick={restoreProfile}>
                  {busy ? t('Restaurando…') : t('Restaurar e reiniciar →')}
                </button>
              ) : (
                <button className="primary-btn" disabled={busy} onClick={confirmDataDir}>
                  {busy ? t('Aplicando…') : t('Continuar →')}
                </button>
              )}
            </div>
            {profileFound && (
              <p className="wz-hint" style={{ marginTop: 10 }}>
                {t('Esta pasta já tem dados do Unit Mail — restaure tudo, ou escolha outra pasta para começar do zero.')}
              </p>
            )}
          </div>
        )}

        {step === 'choice' && (
          <div className="wizard-body">
            <h2>{t('Como você quer começar?')}</h2>
            <div className="wizard-cards">
              <button className="wz-card" onClick={() => setStep('account')}>
                <Icon name="contact" size={32} />
                <strong>{t('Configurar conta')}</strong>
                <span>{t('Adicionar uma conta POP3/IMAP e escolher onde ela entrega')}</span>
              </button>
              <button className="wz-card" onClick={() => setStep('import')}>
                <Icon name="download" size={32} />
                <strong>{t('Importar contas')}</strong>
                <span>
                  {importAvailable
                    ? t('Trazer as contas do Outlook ou do Thunderbird')
                    : t('Trazer as contas do Thunderbird deste sistema')}
                </span>
              </button>
            </div>
            <p className="wz-hint" style={{ marginTop: 4 }}>
              {t('O Unit Mail guarda os emails em bases de dados próprias — criadas automaticamente ao configurar a conta.')}
            </p>
            <div className="wizard-nav" style={{ justifyContent: 'space-between' }}>
              <button onClick={() => setStep('datadir')} style={{ ...inputStyle, width: 'auto' }}>
                {t('← Voltar')}
              </button>
              <button className="wz-skip" style={{ marginTop: 0 }} onClick={finish}>
                {t('Pular configuração inicial →')}
              </button>
            </div>
          </div>
        )}

        {step === 'import' && (
          <div className="wizard-body">
            <h2>{t('Importar contas')}</h2>
            <p className="wz-hint">
              {importAvailable
                ? t('Outlook: lê as contas instaladas neste Windows, com a senha salva. Thunderbird: lê as contas de todos os perfis deste sistema (as senhas você digita).')
                : t('Thunderbird: lê as contas de todos os perfis deste sistema. As senhas dele você digita. Também é possível importar um arquivo .json de exportação.')}
            </p>
            <p className="wz-hint" style={{ marginTop: -8 }}>
              {t('Já tem arquivos')} <strong>.pst</strong>?{' '}
              {t('Você pode importá-los depois em Configurações → Arquivos de Dados → "Importar arquivo .pst…".')}
            </p>
            {guesses.length === 0 ? (
              <div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  {importAvailable && (
                    <button className="primary-btn" disabled={busy} onClick={() => scan('outlook')}>
                      <Icon name="search" size={14} /> {busy ? t('Procurando…') : t('Procurar contas do Outlook')}
                    </button>
                  )}
                  <button className="primary-btn" disabled={busy} onClick={() => scan('thunderbird')}>
                    <Icon name="search" size={14} /> {busy ? t('Procurando…') : t('Procurar contas do Thunderbird')}
                  </button>
                  <button className="btn-ghost" style={{ padding: '6px 0' }} disabled={busy} onClick={openExport}>
                    <Icon name="open" size={14} /> {t('Importar arquivo .json…')}
                  </button>
                </div>
                <p className="wz-hint" style={{ marginTop: 10, marginBottom: 0 }}>
                  {t('Quem copiou a')} <strong>{t('pasta de dados')}</strong>{' '}
                  {t('na etapa anterior já tem tudo — não precisa importar nada aqui. O arquivo')} <strong>.json</strong>{' '}
                  {t('é só para quem não tem a pasta: ele é gerado pelo próprio Unit Mail no Windows (Configurações → Email → "Procurar contas" → "Exportar") e traz apenas as configurações de conta, sem senhas.')}
                </p>
              </div>
            ) : (
              <>
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, marginBottom: 6, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={guesses.every((g) => selectedEmails.has(g.email))}
                    onChange={(e) =>
                      setSelectedEmails(
                        e.target.checked ? new Set(guesses.map((g) => g.email)) : new Set(),
                      )
                    }
                  />
                  {t('Selecionar todas')}
                </label>
                <div className="wz-guess-list">
                  {guesses.map((g) => (
                    <label key={g.email} className="wz-guess-row" style={{ cursor: 'pointer' }}>
                      <input
                        type="checkbox"
                        disabled={addedEmails.has(g.email)}
                        checked={selectedEmails.has(g.email) || addedEmails.has(g.email)}
                        onChange={(e) =>
                          setSelectedEmails((s) => {
                            const n = new Set(s);
                            if (e.target.checked) n.add(g.email);
                            else n.delete(g.email);
                            return n;
                          })
                        }
                      />
                      <span style={{ flex: 1, fontSize: 12 }}>
                        {g.displayName} &lt;{g.email}&gt; — {g.protocol.toUpperCase()}: {g.imapHost}
                        {g.password && (
                          <span style={{ color: '#107c41' }}> · {t('senha recuperada')}</span>
                        )}
                      </span>
                      {addedEmails.has(g.email) && (
                        <span style={{ color: '#107c41', fontSize: 12 }}>✓ {t('adicionada')}</span>
                      )}
                    </label>
                  ))}
                </div>
                <button
                  className="primary-btn"
                  style={{ marginTop: 10, alignSelf: 'flex-start' }}
                  disabled={busy || guesses.every((g) => addedEmails.has(g.email) || !selectedEmails.has(g.email))}
                  onClick={addSelected}
                >
                  {busy
                    ? t('Adicionando…')
                    : t('Adicionar {n} selecionada(s) →', { n: [...selectedEmails].filter((e) => !addedEmails.has(e)).length })}
                </button>
              </>
            )}
            {msg && <div style={{ fontSize: 12, marginTop: 6 }}>{msg}</div>}
            <div className="wizard-nav">
              <button
                onClick={() => setStep(addedEmails.size ? 'done' : 'account')}
                className="primary-btn"
              >
                {addedEmails.size ? t('Concluir →') : t('Continuar →')}
              </button>
              <button onClick={() => setStep('choice')} style={{ ...inputStyle, width: 'auto' }}>
                {t('← Voltar')}
              </button>
            </div>
          </div>
        )}

        {step === 'account' && (
          <div className="wizard-body">
            <h2>{t('Sua primeira conta de email')}</h2>
            <p className="wz-hint">{t('Você pode pular e configurar depois em Configurações.')}</p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 420 }}>
              <input style={inputStyle} placeholder={t('Seu nome (ex.: Roger Oliveira)')} value={label} onChange={(e) => setLabel(e.target.value)} />
              <input style={inputStyle} placeholder={t('Email (ex.: voce@empresa.com)')} value={email} onChange={(e) => setEmail(e.target.value)} />
              <input
                style={inputStyle}
                type="password"
                placeholder={t('Senha (opcional agora — pode configurar depois)')}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
              />
              {email.includes('@') && guessProtocol(email).protocol === 'imap' ? (
                <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                  {t('Esta conta será configurada como IMAP — as mensagens ficam no servidor e nenhum arquivo de dados local é criado.')}
                </span>
              ) : (
                <>
                  <label style={{ fontSize: 12, color: 'var(--text-dim)' }}>
                    {t('Onde as mensagens serão entregues')}
                  </label>
                  <select
                    style={inputStyle}
                    value={storeChoice}
                    onChange={(e) => setStoreChoice(e.target.value)}
                  >
                    <option value="new">{t('Nova base na pasta de dados do app (padrão)')}</option>
                    <option value="newCustom">{t('Nova base em pasta à minha escolha…')}</option>
                    {localStores.map((s) => (
                      <option key={s.id} value={`existing:${s.filePath}`}>
                        {t('Usar base existente: {name}', { name: s.label })}
                      </option>
                    ))}
                    {/* Sem nenhuma base local, a entrega é obrigatória — só resta criar. */}
                    {localStores.length > 0 && (
                      <option value="none">{t('Sem base de dados por enquanto')}</option>
                    )}
                  </select>
                  <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                    {storeChoice === 'new'
                      ? t('Cria a base da conta dentro da pasta de dados (Mail/) — os emails ficam no formato próprio do app.')
                      : storeChoice === 'newCustom'
                        ? t('Você escolhe a pasta onde a nova base de dados será criada.')
                        : storeChoice.startsWith('existing:')
                          ? t('A conta entrega na base já aberta — nada novo é criado.')
                          : t('Sem entrega automática — você pode vincular uma base depois.')}
                  </span>
                  {localStores.length === 0 && (
                    <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                      {t('Não existe nenhum arquivo de dados ainda — é preciso criar um para esta conta (POP3 precisa de onde guardar as mensagens).')}
                    </span>
                  )}
                  <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                    {t('Tem um arquivo .pst do Outlook? Depois use Configurações → Arquivos de Dados → "Importar arquivo .pst…" para trazer tudo para o formato do app.')}
                  </span>
                </>
              )}
            </div>
            {msg && <div style={{ fontSize: 12, marginTop: 8, color: '#c00' }}>{msg}</div>}
            <div className="wizard-nav">
              <button className="primary-btn" disabled={busy || !email.includes('@')} onClick={createAccount}>
                {busy ? t('Criando…') : t('Criar conta')}
              </button>
              <button onClick={() => setStep('done')} style={{ ...inputStyle, width: 'auto' }}>
                {t('Pular →')}
              </button>
              <button onClick={() => setStep('choice')} style={{ ...inputStyle, width: 'auto' }}>
                {t('← Voltar')}
              </button>
            </div>
          </div>
        )}

        {step === 'done' && (
          <div className="wizard-body" style={{ textAlign: 'center' }}>
            <Icon name="task" size={56} />
            <h2>{t('Tudo pronto!')}</h2>
            <p className="wz-hint">
              {t('Use')} <strong>{t('Configurações')}</strong>{' '}
              {t('no ribbon para gerenciar contas, arquivos de dados, regras e assinaturas a qualquer momento.')}
            </p>
            <button className="primary-btn" style={{ margin: '12px auto 0' }} onClick={finish}>
              {t('Começar a usar')}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
