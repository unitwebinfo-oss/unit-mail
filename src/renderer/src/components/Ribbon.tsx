import React, { useEffect, useState } from 'react';
import { Icon } from '../icons';
import { t, useLang } from '../i18n';
import type { AccountConfig, MailCategory, ModuleId } from '../../../shared/types';

export type ReadingPaneLayout = 'right' | 'bottom' | 'off';

/** Paleta dos ícones da barra (estilo Fluent) — cada ação tem sua cor de acento. */
const IC: Record<string, string> = {
  envelope: '#0078d4',
  mail: '#0078d4',
  new: '#107c10',
  junk: '#ca5010',
  delete: '#d13438',
  trash: '#d13438',
  archive: '#8764b8',
  reply: '#0078d4',
  replyAll: '#038387',
  forward: '#107c10',
  flag: '#e81123',
  refresh: '#038387',
  printer: '#69797e',
  sendReceive: '#0078d4',
  download: '#038387',
  sent: '#107c10',
  calendar: '#0078d4',
  contact: '#e3008c',
  task: '#8764b8',
  folder: '#eaa300',
  open: '#eaa300',
  journal: '#69797e',
  close: '#d13438',
  settings: '#69797e',
  pen: '#107c10',
  note: '#eaa300',
};

const Ic = ({ name, size }: { name: string; size?: number }) => (
  <Icon name={name} size={size} color={IC[name]} />
);

/** Links exibidos na tela Ajuda — preencher quando o repositório/site estiver no ar. */
const APP_LINKS = {
  github: 'https://github.com/unitwebinfo-oss/unit-mail',
  site: '', // ex.: 'https://www.empresa.com.br'
  company: '', // nome da empresa desenvolvedora
  donate: 'https://www.paypal.com/donate/?hosted_button_id=X95ATUT9TXGNW',
};

interface RibbonProps {
  module: ModuleId;
  pstOpen: boolean;
  hasAccounts: boolean;
  canReply: boolean;
  accounts: AccountConfig[];
  sourcesCount: number;
  rpLayout: ReadingPaneLayout;
  onOpenSettings: () => void;
  onRefresh: () => void;
  onNew: () => void;
  onNewAppt: () => void;
  onNewContact: () => void;
  onNewTask: () => void;
  onReply: () => void;
  onReplyAll: () => void;
  onForward: () => void;
  onSendReceive: () => void;
  onSyncAccount: (accountId: string) => void;
  /** Abre o diálogo "Progresso de envio/recebimento" (não dispara o sync). */
  onShowProgress: () => void;
  onLayoutChange: (l: ReadingPaneLayout) => void;
  /** Excluir o item selecionado (só itens locais são removidos de verdade). */
  onDeleteItem: () => void;
  /** Mover o item selecionado para uma pasta por nome (ex.: 'Lixo eletrônico'). */
  onMoveItemTo: (folderName: string) => void;
  /** Abre as Configurações direto numa aba específica; accountId abre já em edição;
   *  action dispara um fluxo (ex.: importação .pst/Thunderbird) ao abrir. */
  onOpenSettingsTab: (
    tab: 'email' | 'data' | 'rules' | 'sigs' | 'geral',
    accountId?: string,
    action?: 'importPst' | 'importTb',
  ) => void;
  /** Pasta selecionada (para a aba Pasta). */
  selFolder: { sourceId: string; folderId: string; name: string; isLocal: boolean } | null;
  /** Destinos possíveis para "Mover Pasta" (árvore achatada da base). */
  folderTargets: Array<{ id: string | null; label: string }>;
  onFolderNew: (name: string) => void;
  onFolderRename: (name: string) => void;
  onFolderDelete: () => void;
  onFolderCopy: () => void;
  onFolderMoveTo: (destId: string | null) => void;
  onMarkAllReadSel: () => void;
  onRunRules: () => void;
  /** Exporta o backup de configurações (contas, regras, assinaturas). */
  onExportBackup: () => void;
  /** Importa um backup de configurações. */
  onImportBackup: () => void;
  /** Salva a mensagem aberta como .eml (Arquivo → Salvar como). */
  onSaveAs: () => void;
  /** Imprime a mensagem aberta (Arquivo → Imprimir). */
  onPrint: () => void;
  /** Lista-mestre de categorias (grupo Marcas → Categorizar). */
  categories: MailCategory[];
  /** Aplica/remove uma categoria nos itens selecionados. */
  onCategorize: (catId: string) => void;
  /** Abre o diálogo "Todas as Categorias". */
  onCatsManage: () => void;
}

type TabId = 'arquivo' | 'inicio' | 'enviar' | 'pasta' | 'exibicao';

function Group({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="ribbon-group">
      <div className="ribbon-buttons">{children}</div>
      <div className="ribbon-group-label">{t(label)}</div>
    </div>
  );
}

function BigBtn({
  icon,
  label,
  onClick,
  disabled,
  title,
}: {
  icon: string;
  label: string;
  onClick?: () => void;
  disabled?: boolean;
  title?: string;
}) {
  return (
    <button className="ribbon-btn" onClick={onClick} disabled={disabled} title={t(title ?? label)}>
      <Ic name={icon} />
      <span>{t(label)}</span>
    </button>
  );
}

function SmallBtn({
  icon,
  label,
  onClick,
  disabled,
  title,
}: {
  icon: string;
  label: string;
  onClick?: () => void;
  disabled?: boolean;
  title?: string;
}) {
  return (
    <button className="ribbon-btn small" onClick={onClick} disabled={disabled} title={t(title ?? label)}>
      <Ic name={icon} />
      <span>{t(label)}</span>
    </button>
  );
}

export function Ribbon(props: RibbonProps) {
  useLang();
  const [tab, setTab] = useState<TabId>('inicio');
  const [syncMin, setSyncMin] = useState(30);
  const [newMenu, setNewMenu] = useState(false);
  const [grpMenu, setGrpMenu] = useState(false);
  const [bsPage, setBsPage] = useState<'info' | 'open' | 'help'>('info');
  /** Conta escolhida no seletor de "Informações sobre Contas". */
  const [bsAccount, setBsAccount] = useState('');
  const [appInfo, setAppInfo] = useState<Awaited<
    ReturnType<typeof window.outlook.getAppInfo>
  > | null>(null);
  const [pastaAction, setPastaAction] = useState<'new' | 'rename' | 'move' | null>(null);
  const [pastaName, setPastaName] = useState('');
  /** Popover do botão Categorizar (aba Página Inicial, grupo Marcas). */
  const [catsMenu, setCatsMenu] = useState(false);

  useEffect(() => {
    Promise.resolve(window.outlook.getSyncInterval?.() ?? 30)
      .then(setSyncMin)
      .catch(() => {});
    window.outlook.getAppInfo?.().then(setAppInfo).catch(() => {});
  }, []);
  const {
    module, pstOpen, hasAccounts, canReply, accounts, rpLayout,
    onRefresh, onNew, onNewAppt, onNewContact, onNewTask, onReply, onReplyAll, onForward,
    onSendReceive, onSyncAccount, onShowProgress, onLayoutChange, onDeleteItem, onMoveItemTo,
    onOpenSettingsTab,
    selFolder, folderTargets,
    onFolderNew, onFolderRename, onFolderDelete, onFolderCopy, onFolderMoveTo,
    onMarkAllReadSel, onRunRules,
    onExportBackup, onImportBackup, onSaveAs, onPrint,
    categories, onCategorize, onCatsManage,
  } = props;

  const tabs: Array<{ id: TabId; label: string; file?: boolean }> = [
    { id: 'arquivo', label: 'Arquivo', file: true },
    { id: 'inicio', label: 'Página Inicial' },
    { id: 'enviar', label: 'Enviar/Receber' },
    { id: 'pasta', label: 'Pasta' },
    { id: 'exibicao', label: 'Exibição' },
  ];

  return (
    <div
      className="ribbon"
      onClick={() => {
        if (newMenu) setNewMenu(false);
        if (grpMenu) setGrpMenu(false);
        if (catsMenu) setCatsMenu(false);
      }}
    >
      <div className="ribbon-tabs">
        {tabs.map((tb) => (
          <div
            key={tb.id}
            className={`ribbon-tab ${tb.file ? 'file-tab' : ''} ${tab === tb.id && !tb.file ? 'active' : ''}`}
            onClick={() => setTab(tb.id)}
          >
            {t(tb.label)}
          </div>
        ))}
      </div>

      {/* Backstage (aba Arquivo) — tela cheia estilo Outlook */}
      {tab === 'arquivo' && (
        <div className="backstage" onClick={(e) => e.stopPropagation()}>
          <div className="bs-side">
            <button className="bs-back" onClick={() => setTab('inicio')} title={t('Voltar')}>
              <Icon name="reply" size={18} />
            </button>
            <button
              className={`bs-item ${bsPage === 'info' ? 'active' : ''}`}
              onClick={() => setBsPage('info')}
            >
              <Icon name="contact" size={16} /> {t('Informações')}
            </button>
            <button
              className={`bs-item ${bsPage === 'open' ? 'active' : ''}`}
              onClick={() => setBsPage('open')}
            >
              <Icon name="open" size={16} /> {t('Abrir e Exportar')}
            </button>
            <div className="bs-sep" />
            <button className="bs-item" disabled={!canReply} onClick={onSaveAs} title={canReply ? t('Salvar a mensagem aberta como .eml') : t('Abra uma mensagem para salvar')}>
              <Icon name="archive" size={16} /> {t('Salvar como')}
            </button>
            <button className="bs-item" disabled={!canReply} onClick={onPrint} title={canReply ? t('Imprimir a mensagem aberta') : t('Abra uma mensagem para imprimir')}>
              <Icon name="printer" size={16} /> {t('Imprimir')}
            </button>
            <div className="bs-sep" />
            <button
              className={`bs-item ${bsPage === 'help' ? 'active' : ''}`}
              onClick={() => setBsPage('help')}
            >
              <Icon name="envelope" size={16} /> {t('Ajuda')}
            </button>
            <div className="bs-sep" />
            <button
              className="bs-item"
              onClick={() => {
                setTab('inicio');
                onOpenSettingsTab('geral');
              }}
              title={t('Opções e preferências do aplicativo')}
            >
              <Icon name="settings" size={16} /> {t('Configurações')}
            </button>
          </div>
          <div className="bs-main">
            {bsPage === 'info' && (
              <>
                <h1>{t('Informações sobre Contas')}</h1>
                <select
                  className="bs-account"
                  disabled={accounts.length === 0}
                  value={bsAccount || accounts[0]?.id || ''}
                  onChange={(e) => setBsAccount(e.target.value)}
                >
                  {accounts.length === 0 && <option>{t('Nenhuma conta configurada')}</option>}
                  {accounts.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.email} — {a.protocol.toUpperCase()}
                    </option>
                  ))}
                </select>
                <button
                  className="bs-add"
                  onClick={() => {
                    setTab('inicio');
                    onOpenSettingsTab('email');
                  }}
                >
                  {t('+ Adicionar Conta')}
                </button>
                <div className="bs-tiles">
                  <button
                    className="bs-tile"
                    onClick={() => {
                      setTab('inicio');
                      // A conta escolhida no seletor acima abre já em edição
                      onOpenSettingsTab('email', bsAccount || accounts[0]?.id);
                    }}
                  >
                    <Ic name="contact" size={28} />
                    <div>
                      <strong>{t('Configurações de Conta')}</strong>
                      <span>{t('Altere configurações para esta conta ou configure mais conexões.')}</span>
                    </div>
                  </button>
                  <button
                    className="bs-tile"
                    onClick={() => {
                      setTab('inicio');
                      onOpenSettingsTab('data');
                    }}
                  >
                    <Ic name="archive" size={28} />
                    <div>
                      <strong>{t('Configurações de Caixa de Correio')}</strong>
                      <span>{t('Gerencie arquivos de dados, bases locais e onde as mensagens são entregues.')}</span>
                    </div>
                  </button>
                  <button
                    className="bs-tile"
                    onClick={() => {
                      setTab('inicio');
                      onOpenSettingsTab('rules');
                    }}
                  >
                    <Ic name="flag" size={28} />
                    <div>
                      <strong>{t('Gerenciar Regras e Alertas')}</strong>
                      <span>{t('Use Regras e Alertas para organizar as mensagens de email de entrada.')}</span>
                    </div>
                  </button>
                </div>
                <div className="bs-foot">
                  {t('{n} base(s) de dados aberta(s)', { n: props.sourcesCount })}
                </div>
              </>
            )}
            {bsPage === 'open' && (
              <>
                <h1>{t('Abrir')}</h1>
                <div className="bs-tiles">
                  <button
                    className="bs-tile"
                    onClick={() => {
                      setTab('inicio');
                      onOpenSettingsTab('data', undefined, 'importPst');
                    }}
                  >
                    <Ic name="open" size={28} />
                    <div>
                      <strong>{t('Importar .pst (Outlook)')}</strong>
                      <span>
                        {t('Traz todas as pastas e mensagens de um arquivo .pst para uma base do Unit Mail — o arquivo original nunca é alterado.')}
                      </span>
                    </div>
                  </button>
                  <button
                    className="bs-tile"
                    onClick={() => {
                      setTab('inicio');
                      onOpenSettingsTab('data', undefined, 'importTb');
                    }}
                  >
                    <Ic name="download" size={28} />
                    <div>
                      <strong>{t('Importar do Thunderbird')}</strong>
                      <span>
                        {t('Traz todas as pastas e mensagens do perfil do Thunderbird instalado neste computador — o perfil original nunca é alterado.')}
                      </span>
                    </div>
                  </button>
                  <button
                    className="bs-tile"
                    onClick={() => {
                      setTab('inicio');
                      onOpenSettingsTab('data');
                    }}
                  >
                    <Ic name="folder" size={28} />
                    <div>
                      <strong>{t('Novo Arquivo de Dados')}</strong>
                      <span>{t('Crie uma base local gravável para receber e organizar mensagens.')}</span>
                    </div>
                  </button>
                </div>

                <h1 style={{ marginTop: 32 }}>{t('Exportar')}</h1>
                <div className="bs-tiles">
                  <button className="bs-tile" onClick={onExportBackup}>
                    <Ic name="download" size={28} />
                    <div>
                      <strong>{t('Exportar configurações (backup)')}</strong>
                      <span>
                        {t('Salva contas, regras e assinaturas num arquivo .json para levar para outra máquina. O arquivo contém as senhas — guarde com segurança.')}
                      </span>
                    </div>
                  </button>
                  <button className="bs-tile" onClick={onImportBackup}>
                    <Ic name="archive" size={28} />
                    <div>
                      <strong>{t('Importar configurações')}</strong>
                      <span>
                        {t('Restaura contas, regras e assinaturas de um backup. Processo recomendado: primeiro adicione os arquivos de dados, depois importe o backup e re-vincule as contas em Configurações → Email.')}
                      </span>
                    </div>
                  </button>
                </div>
              </>
            )}
            {bsPage === 'help' && (
              <div className="help-page">
                {/* Hero — identidade do app */}
                <div className="help-hero">
                  <div className="help-hero-icon">
                    <Ic name="logo" size={40} />
                  </div>
                  <div className="help-hero-info">
                    <div className="help-hero-name">
                      Unit Mail <span className="help-ver">v{appInfo?.version ?? ''}</span>
                    </div>
                    <div className="help-hero-tag">
                      Email, calendário e contatos no desktop — com leitura de arquivos .pst.
                    </div>
                    <div className="help-badges">
                      <span className="help-badge">{appInfo?.license ?? 'GPL-3.0-only'}</span>
                      <span className="help-badge">Código aberto</span>
                    </div>
                  </div>
                </div>

                {/* Aviso único e importante */}
                <div className="help-note">
                  <Ic name="archive" size={16} />
                  <span>
                    Arquivos <strong>.pst</strong> são abertos <strong>somente para leitura</strong> —
                    o original nunca é alterado. Para excluir e organizar de verdade, use{' '}
                    <strong>Migrar</strong> em Configurações → Arquivos de Dados.
                  </span>
                </div>

                {/* Sobre o aplicativo */}
                <div className="help-about">
                  <div className="help-about-row">
                    <span>Versão</span>
                    <strong>{appInfo?.version ?? '—'}</strong>
                  </div>
                  <div className="help-about-row">
                    <span>Licença</span>
                    <strong>{appInfo?.license ?? 'GPL-3.0-only'} — código aberto</strong>
                  </div>
                  <div className="help-about-row">
                    <span>Plataforma</span>
                    <strong>{appInfo?.platform ?? ''}</strong>
                  </div>
                  {APP_LINKS.company && (
                    <div className="help-about-row">
                      <span>Desenvolvido por</span>
                      <strong>{APP_LINKS.company}</strong>
                    </div>
                  )}
                </div>

                {(APP_LINKS.github || APP_LINKS.site || APP_LINKS.donate) && (
                  <div className="help-links">
                    {APP_LINKS.donate && (
                      <button
                        className="bs-add help-donate"
                        onClick={() => window.outlook.openExternal?.(APP_LINKS.donate)}
                      >
                        Apoiar o projeto — doar via PayPal
                      </button>
                    )}
                    {APP_LINKS.github && (
                      <button
                        className="bs-add"
                        onClick={() => window.outlook.openExternal?.(APP_LINKS.github)}
                      >
                        GitHub — código-fonte e empresa
                      </button>
                    )}
                    {APP_LINKS.site && (
                      <button
                        className="bs-add"
                        onClick={() => window.outlook.openExternal?.(APP_LINKS.site)}
                      >
                        Site oficial
                      </button>
                    )}
                  </div>
                )}

                <div className="help-foot">
                  <span>
                    Electron {appInfo?.electron} · Chromium {appInfo?.chrome} · Node {appInfo?.node}
                  </span>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {tab === 'inicio' && (
        <div className="ribbon-body" style={{ position: 'relative' }}>
          {module === 'mail' && (
            <>
              <Group label="Novo">
                <BigBtn icon="envelope" label="Novo Email" onClick={onNew} disabled={!hasAccounts} />
                <div className="ribbon-col">
                  <div style={{ position: 'relative' }}>
                    <button
                      className="ribbon-btn small"
                      onClick={(e) => {
                        e.stopPropagation();
                        setNewMenu((v) => !v);
                      }}
                    >
                      <Ic name="new" />
                      <span>{t('Novos Itens')} ▾</span>
                    </button>
                    {newMenu && (
                      <div className="ribbon-menu" onClick={(e) => e.stopPropagation()}>
                        <button
                          onClick={() => {
                            setNewMenu(false);
                            onNew();
                          }}
                          disabled={!hasAccounts}
                        >
                          <Ic name="envelope" size={14} /> {t('Novo Email')}
                        </button>
                        <button
                          onClick={() => {
                            setNewMenu(false);
                            onNewAppt();
                          }}
                          disabled={!props.sourcesCount}
                        >
                          <Ic name="calendar" size={14} /> {t('Novo Compromisso')}
                        </button>
                        <button
                          onClick={() => {
                            setNewMenu(false);
                            onNewContact();
                          }}
                          disabled={!props.sourcesCount}
                        >
                          <Ic name="contact" size={14} /> {t('Novo Contato')}
                        </button>
                        <button
                          onClick={() => {
                            setNewMenu(false);
                            onNewTask();
                          }}
                          disabled={!props.sourcesCount}
                        >
                          <Ic name="task" size={14} /> {t('Nova Tarefa')}
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              </Group>
              <Group label="Excluir">
                <BigBtn
                  icon="junk"
                  label="Lixo Eletrônico"
                  onClick={() => onMoveItemTo('Lixo eletrônico')}
                  disabled={!canReply}
                />
                <BigBtn icon="delete" label="Excluir" onClick={onDeleteItem} disabled={!canReply} />
                <BigBtn
                  icon="archive"
                  label="Arquivar"
                  onClick={() => onMoveItemTo('Arquivo Morto')}
                  disabled={!canReply}
                />
              </Group>
              <Group label="Responder">
                <BigBtn icon="reply" label="Responder" onClick={onReply} disabled={!canReply} />
                <BigBtn icon="replyAll" label="Responder a Todos" onClick={onReplyAll} disabled={!canReply} />
                <BigBtn icon="forward" label="Encaminhar" onClick={onForward} disabled={!canReply} />
              </Group>
              <Group label="Marcas">
                <div style={{ position: 'relative' }}>
                  <button
                    className="ribbon-btn"
                    disabled={!canReply}
                    title={
                      canReply
                        ? t('Atribuir categorias coloridas aos itens selecionados')
                        : t('Abra uma mensagem para categorizar')
                    }
                    onClick={(e) => {
                      e.stopPropagation();
                      setCatsMenu((v) => !v);
                    }}
                  >
                    <Ic name="flag" />
                    <span>{t('Categorizar')} ▾</span>
                  </button>
                  {catsMenu && (
                    <div className="ribbon-menu" onClick={(e) => e.stopPropagation()}>
                      {categories.map((c) => (
                        <button
                          key={c.id}
                          onClick={() => {
                            setCatsMenu(false);
                            onCategorize(c.id);
                          }}
                        >
                          <span className="cat-dot" style={{ background: c.color }} />
                          {c.name}
                        </button>
                      ))}
                      <div className="ribbon-menu-sep" />
                      <button
                        onClick={() => {
                          setCatsMenu(false);
                          onCatsManage();
                        }}
                      >
                        <Ic name="flag" size={14} /> {t('Todas as Categorias…')}
                      </button>
                    </div>
                  )}
                </div>
              </Group>
            </>
          )}
          {module === 'calendar' && (
            <Group label="Compromisso">
              <BigBtn icon="new" label="Novo Compromisso" onClick={onNewAppt} disabled={!props.sourcesCount} />
            </Group>
          )}
          {module === 'people' && (
            <Group label="Contato">
              <BigBtn icon="new" label="Novo Contato" onClick={onNewContact} disabled={!props.sourcesCount} />
            </Group>
          )}
          {module === 'tasks' && (
            <Group label="Tarefa">
              <BigBtn icon="new" label="Nova Tarefa" onClick={onNewTask} disabled={!props.sourcesCount} />
            </Group>
          )}
          <Group label="Ferramentas">
            <BigBtn icon="refresh" label="Atualizar" onClick={onRefresh} disabled={!pstOpen} />
            <BigBtn
              icon="printer"
              label="Imprimir"
              onClick={onPrint}
              disabled={!canReply}
              title={canReply ? 'Imprimir a mensagem aberta' : 'Abra uma mensagem para imprimir'}
            />
          </Group>
          <Group label="Enviar e Receber">
            <BigBtn
              icon="sendReceive"
              label="Enviar/Receber"
              onClick={onSendReceive}
              disabled={!hasAccounts}
              title="Enviar e receber mensagens de todas as contas"
            />
          </Group>
        </div>
      )}

      {tab === 'enviar' && (
        <div className="ribbon-body">
          <Group label="Enviar e Receber">
            <BigBtn
              icon="download"
              label="Enviar/Receber Todas as Pastas"
              onClick={onSendReceive}
              disabled={!hasAccounts}
            />
            <div className="ribbon-col">
              <SmallBtn icon="refresh" label="Atualizar Pasta" onClick={onRefresh} disabled={!pstOpen} />
              <SmallBtn icon="sent" label="Enviar Tudo" onClick={onSendReceive} disabled={!hasAccounts} />
              <div style={{ position: 'relative' }}>
                <button
                  className="ribbon-btn small"
                  onClick={(e) => {
                    e.stopPropagation();
                    setGrpMenu((v) => !v);
                  }}
                >
                  <Ic name="envelope" />
                  <span>{t('Grupos de Envio/Recebimento')} ▾</span>
                </button>
                {grpMenu && (
                  <div className="ribbon-menu" onClick={(e) => e.stopPropagation()}>
                    {accounts.length === 0 && (
                      <button disabled>{t('Nenhuma conta configurada')}</button>
                    )}
                    {accounts.map((a) => (
                      <button
                        key={a.id}
                        onClick={() => {
                          setGrpMenu(false);
                          onSyncAccount(a.id);
                        }}
                      >
                        <Ic name="envelope" size={14} /> {a.email}
                      </button>
                    ))}
                    {accounts.length > 0 && <div className="ribbon-menu-sep" />}
                    <div className="ribbon-menu-row">
                      {t('Sincronizar a cada')}
                      <select
                        value={syncMin}
                        onChange={(e) => {
                          const v = Number(e.target.value);
                          setSyncMin(v);
                          window.outlook.setSyncInterval(v).catch(() => {});
                        }}
                      >
                        <option value={0}>{t('desligado')}</option>
                        {[1, 5, 10, 15, 30, 60].map((v) => (
                          <option key={v} value={v}>
                            {t('{n} min', { n: v })}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </Group>
          <Group label="Baixar">
            <BigBtn icon="journal" label="Mostrar Progresso" onClick={onShowProgress} disabled={!hasAccounts} />
          </Group>
        </div>
      )}

      {tab === 'pasta' && (
        <div className="ribbon-body" style={{ position: 'relative' }}>
          <Group label="Novo">
            <BigBtn
              icon="folder"
              label="Nova Pasta"
              disabled={!selFolder}
              onClick={() => {
                setPastaName('');
                setPastaAction('new');
              }}
            />
          </Group>
          <Group label="Ações">
            <BigBtn
              icon="draft"
              label="Renomear Pasta"
              disabled={!selFolder?.isLocal}
              title={selFolder && !selFolder.isLocal ? 'Pasta do PST é somente leitura' : undefined}
              onClick={() => {
                setPastaName(selFolder?.name || '');
                setPastaAction('rename');
              }}
            />
            <div className="ribbon-col">
              <SmallBtn
                icon="folder"
                label="Copiar Pasta"
                onClick={onFolderCopy}
                disabled={!selFolder?.isLocal}
              />
              <SmallBtn
                icon="open"
                label="Mover Pasta"
                onClick={() => setPastaAction('move')}
                disabled={!selFolder?.isLocal}
              />
              <SmallBtn
                icon="delete"
                label="Excluir Pasta"
                onClick={onFolderDelete}
                disabled={!selFolder?.isLocal}
              />
            </div>
          </Group>
          <Group label="Limpar">
            <BigBtn
              icon="envelope"
              label="Marcar Todas como Lidas"
              onClick={onMarkAllReadSel}
              disabled={!selFolder}
            />
            <BigBtn
              icon="flag"
              label="Executar Regras Agora"
              onClick={onRunRules}
              disabled={!selFolder}
            />
          </Group>
          <Group label="Pasta atual">
            <BigBtn icon="refresh" label="Atualizar" onClick={onRefresh} disabled={!pstOpen} />
          </Group>

          {/* Popover de ação da aba Pasta */}
          {pastaAction && (
            <div className="ctx-menu ribbon-pop" onClick={(e) => e.stopPropagation()}>
              {pastaAction === 'move' ? (
                <div className="ctx-move">
                  <div className="ctx-move-title">
                    {t('Mover "{nome}" para:', { nome: selFolder?.name || '' })}
                  </div>
                  {folderTargets.map((ft) => (
                    <button
                      key={ft.id ?? '__root'}
                      onClick={() => {
                        setPastaAction(null);
                        onFolderMoveTo(ft.id);
                      }}
                    >
                      {ft.label}
                    </button>
                  ))}
                  <div className="ctx-sep" />
                  <button onClick={() => setPastaAction(null)}>{t('Cancelar')}</button>
                </div>
              ) : (
                <div className="ctx-input">
                  <label>
                    {pastaAction === 'new' ? t('Nome da nova pasta') : t('Novo nome da pasta')}
                  </label>
                  <input
                    autoFocus
                    value={pastaName}
                    onChange={(e) => setPastaName(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && pastaName.trim()) {
                        const v = pastaName.trim();
                        setPastaAction(null);
                        if (pastaAction === 'new') onFolderNew(v);
                        else onFolderRename(v);
                      }
                      if (e.key === 'Escape') setPastaAction(null);
                    }}
                  />
                  <div className="ctx-input-actions">
                    <button onClick={() => setPastaAction(null)}>{t('Cancelar')}</button>
                    <button
                      className="primary"
                      onClick={() => {
                        if (!pastaName.trim()) return;
                        const v = pastaName.trim();
                        setPastaAction(null);
                        if (pastaAction === 'new') onFolderNew(v);
                        else onFolderRename(v);
                      }}
                    >
                      OK
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {tab === 'exibicao' && (
        <div className="ribbon-body">
          <Group label="Painel de Leitura">
            <BigBtn
              icon="chevronR"
              label="À Direita"
              onClick={() => onLayoutChange('right')}
              disabled={rpLayout === 'right'}
            />
            <BigBtn
              icon="chevronD"
              label="Inferior"
              onClick={() => onLayoutChange('bottom')}
              disabled={rpLayout === 'bottom'}
            />
            <BigBtn
              icon="close"
              label="Desativado"
              onClick={() => onLayoutChange('off')}
              disabled={rpLayout === 'off'}
            />
          </Group>
        </div>
      )}
    </div>
  );
}
