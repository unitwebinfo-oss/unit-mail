// Tipos compartilhados entre processo principal (Electron) e renderer

import { folderKindOf, defaultFolderNames, type Lang } from './folderNames';

/**
 * Pastas padrão de uma base de dados — criadas automaticamente e usadas por
 * fluxos internos (ex.: excluir → "Itens Excluídos", antivírus → "Itens
 * infectados"). Não podem ser excluídas, renomeadas nem movidas.
 * Mantida em pt-BR por compatibilidade — criação usa defaultFolderNames(lang).
 */
export const DEFAULT_FOLDERS = defaultFolderNames('pt');

/** A pasta é uma das padrão? Só vale para ids "arc:<Nome>" de primeiro nível
 *  — subpastas criadas pelo usuário dentro delas seguem livres. Reconhece os
 *  nomes padrão em qualquer idioma suportado. */
export function isDefaultFolderId(folderId: string): boolean {
  return folderId.startsWith('arc:') && folderKindOf(folderId.slice(4)) !== null;
}

export type { Lang };

export interface PstOpenResult {
  fileName: string;
  storeName: string;
  filePath: string;
}

export interface FolderNode {
  /** Caminho de índices a partir da raiz, ex: "0/2/1" */
  id: string;
  name: string;
  /** Classe do contêiner MAPI: IPF.Note, IPF.Appointment, IPF.Contact, IPF.Task, etc. */
  containerClass: string;
  contentCount: number;
  unreadCount: number;
  children: FolderNode[];
  /**
   * Nome do diretório-sombra (.eml) mesclado a esta pasta do PST — mensagens
   * entregues pelo Enviar/Receber aparecem dentro da própria pasta, como no Outlook.
   */
  archiveName?: string;
  /** Pasta ocultada pelo usuário ("não uso esta pasta") — some da árvore
   *  sem apagar conteúdo; reversível em "Mostrar pastas ocultas". */
  hidden?: boolean;
  /** IMAP SPECIAL-USE (\Sent, \Trash…) — pasta especial do servidor,
   *  não pode ser renomeada nem excluída. */
  specialUse?: string;
}

/** Categoria colorida estilo Outlook — lista-mestre em categories.json. */
export interface MailCategory {
  id: string;
  name: string;
  /** Cor hex do selo (ex.: '#d13438'). */
  color: string;
}

/** As seis categorias padrão do Outlook. */
export const DEFAULT_CATEGORIES: MailCategory[] = [
  { id: 'red', name: 'Categoria Vermelha', color: '#d13438' },
  { id: 'orange', name: 'Categoria Laranja', color: '#f7630c' },
  { id: 'yellow', name: 'Categoria Amarela', color: '#eaa300' },
  { id: 'green', name: 'Categoria Verde', color: '#107c10' },
  { id: 'blue', name: 'Categoria Azul', color: '#0078d4' },
  { id: 'purple', name: 'Categoria Roxa', color: '#5c2e91' },
];

export interface AttachmentMeta {
  index: number;
  filename: string;
  size: number;
  mimeTag: string;
  contentId: string;
  /** Anexo é uma mensagem embutida (ex: email encaminhado) */
  isEmbeddedMessage: boolean;
  /** Imagem inline referenciada via cid: no corpo HTML */
  isInline: boolean;
}

export interface ItemSummary {
  index: number;
  messageClass: string;
  subject: string;
  senderName: string;
  senderEmailAddress: string;
  displayTo: string;
  displayCC: string;
  /** ISO string ou null */
  sentTime: string | null;
  deliveryTime: string | null;
  hasAttachments: boolean;
  isRead: boolean;
  importance: number;
  /** Tamanho da mensagem em bytes (coluna "Tamanho" estilo Outlook) */
  size?: number;
  /** Primeiros ~255 chars do corpo (preview estilo Outlook) */
  preview: string;
  /** Campos específicos de compromisso/tarefa/contato, quando aplicável */
  startTime?: string | null;
  endTime?: string | null;
  location?: string;
  dueDate?: string | null;
  percentComplete?: number;
  isTaskComplete?: boolean;
  contactDisplayName?: string;
  contactEmail?: string;
  contactPhone?: string;
  contactCompany?: string;
  /** Lembrete de compromisso (minutos antes do início) */
  reminderSet?: boolean;
  reminderMinutes?: number;
  /** Compromisso de dia inteiro */
  allDay?: boolean;
  /** Convite/resposta de reunião — método iMIP (REQUEST/CANCEL/REPLY) quando
   *  o item tem parte text/calendar ou é IPM.Schedule.Meeting.*. */
  inviteMethod?: string;
  /** UID iCalendar do VEVENT (compromissos/convites) — correlaciona o convite
   *  original ao item criado no Calendário ao aceitar. */
  apptUid?: string;
  /** Ids de MailCategory aplicados ao item (Categorizar). */
  categories?: string[];
  /** Mensagem já respondida (IMAP \Answered / PST lastVerb) — ícone ↩ estilo Outlook. */
  isAnswered?: boolean;
  /** Mensagem encaminhada (IMAP $Forwarded / PST lastVerb) — ícone → estilo Outlook. */
  isForwarded?: boolean;
  /** Item .eml da parte local da pasta (sombra/base) — itens do PST ficam indefinidos. */
  local?: boolean;
  /** Identidade estável do item — sobrevive à renumeração de `index` quando a
   *  pasta recompacta (exclusão/move). .eml: nome do arquivo; PST: posição
   *  bruta; IMAP: UID. Usado pelo filtro "Não lidas" para não mostrar itens
   *  errados após reindex. */
  stableId?: string;
  /** E-mail da conta que recebeu a mensagem (carimbo X-UM-Account gravado na
   *  entrega) — usado pelo "Executar regras" para o filtro por conta. */
  accountEmail?: string;
}

/** Convite de reunião (iCalendar / IPM.Schedule.Meeting.*) detectado na mensagem. */
export interface InviteInfo {
  /** REQUEST (convite), CANCEL (cancelamento), REPLY (resposta) */
  method: string;
  uid: string;
  summary: string;
  start: string | null;
  end: string | null;
  location: string;
  organizer: string;
  organizerEmail: string;
  /** iCal bruto do VEVENT, para responder/reagendar */
  ics: string;
  /** Resposta já enviada ('accept'|'decline'|'counter') — barra mostra o
   *  estado em vez dos botões, como o Outlook ("Você respondeu em …"). */
  responded?: string;
}

export interface ItemDetail extends ItemSummary {
  body: string;
  bodyHTML: string;
  transportMessageHeaders: string;
  attachments: AttachmentMeta[];
  // Compromisso
  allAttendees?: string;
  requiredAttendees?: string;
  isRecurring?: boolean;
  recurrencePattern?: string;
  organizer?: string;
  // Contato
  contact?: ContactDetail;
  // Tarefa
  task?: TaskDetail;
  /** Convite de reunião detectado (Aceitar/Recusar/Propor novo horário) */
  invite?: InviteInfo;
  /** RFC822/EML do anexo embutido (quando o próprio item é extraído) */
  notImplemented?: never;
}

export interface ContactDetail {
  displayName: string;
  givenName: string;
  middleName: string;
  surname: string;
  company: string;
  jobTitle: string;
  department: string;
  email1: string;
  email1Address: string;
  email2: string;
  email2Address: string;
  email3: string;
  email3Address: string;
  businessPhone: string;
  homePhone: string;
  mobilePhone: string;
  businessAddress: string;
  homeAddress: string;
  birthday: string | null;
  notes: string;
}

export interface TaskDetail {
  status: number;
  percentComplete: number;
  isComplete: boolean;
  startDate: string | null;
  dueDate: string | null;
  dateCompleted: string | null;
  owner: string;
  body: string;
}

/** Dados do diálogo "Novo compromisso / contato / tarefa" (Página Inicial → Novos Itens). */
export interface NewSpecialItem {
  kind: 'appt' | 'contact' | 'task';
  subject: string;
  body?: string;
  // Compromisso
  start?: string;
  end?: string;
  location?: string;
  /** Minutos antes do início; null/undefined = sem lembrete */
  reminderMin?: number | null;
  allDay?: boolean;
  attendees?: string;
  organizer?: string;
  /** UID iCalendar — compartilhado entre o .ics gravado e os convites enviados. */
  uid?: string;
  /** SEQUENCE iCalendar — incrementa a cada edição do compromisso. */
  seq?: number;
  // Contato
  contactName?: string;
  givenName?: string;
  surname?: string;
  company?: string;
  jobTitle?: string;
  dept?: string;
  email?: string;
  email2?: string;
  phone?: string;
  mobile?: string;
  birthday?: string;
  // Tarefa
  dueDate?: string;
  startDate?: string;
  percentComplete?: number;
  isComplete?: boolean;
  owner?: string;
}

export interface AttachmentData {
  filename: string;
  mimeTag: string;
  dataBase64: string;
  /** Se for mensagem embutida, devolve resumo em vez de bytes */
  embeddedMessage?: ItemDetail;
}

export interface SearchResult {
  folderId: string;
  folderName: string;
  item: ItemSummary;
}

export type ModuleId = 'mail' | 'calendar' | 'people' | 'tasks';

// ===================== Múltiplas bases / contas =====================

export type SourceKind = 'pst' | 'imap' | 'local';

/** Informação de uma base/conta aberta, exibida como nó raiz na árvore de pastas */
export interface MailSource {
  /** Identificador único prefixado: "pst:<uuid>", "local:<uuid>" ou "acct:<uuid>" */
  id: string;
  kind: SourceKind;
  /** Nome exibido (nome do arquivo PST, da base local ou endereço de email da conta) */
  label: string;
  /** Caminho do arquivo PST ou diretório da base local */
  filePath?: string;
  /** true enquanto a conexão IMAP está sendo estabelecida */
  connecting?: boolean;
  /** mensagem de erro de conexão, se houver */
  error?: string;
  /**
   * Contas POP3 são apenas transporte de recebimento — igual ao Outlook, não
   * aparecem como nó na árvore de pastas; as mensagens são entregues na base
   * de dados vinculada à conta.
   */
  hidden?: boolean;
}

/** Configuração persistida de uma conta de email (sem senha) */
export interface AccountConfig {
  id: string;
  label: string;
  email: string;
  /** Protocolo de recebimento */
  protocol: 'imap' | 'pop3';
  /** Servidor de entrada (IMAP ou POP3 conforme `protocol`) */
  imapHost: string;
  imapPort: number;
  imapSecure: boolean;
  imapUser: string;
  smtpHost: string;
  smtpPort: number;
  smtpSecure: boolean;
  smtpUser: string;
  /** Assinatura (HTML) embutida — legado/importada; migrada para Signature */
  signature?: string;
  /** Indica que existe uma senha salva no main (a senha nunca sai do processo principal) */
  hasPassword?: boolean;
  /** Indica senha de SMTP separada da de entrada */
  hasSmtpPassword?: boolean;
  /** Assinatura nomeada para novas mensagens (Signature.id) */
  signatureNewId?: string;
  /** Assinatura nomeada para respostas/encaminhamentos (Signature.id) */
  signatureReplyId?: string;
  /** Base onde novas mensagens são entregues por padrão (caminho do PST ou do store local) */
  deliveryPath?: string;
  /** Pasta de entrega dentro da base (padrão "Caixa de Entrada"); para PST deve ser o nome real da pasta para a sombra mesclar corretamente */
  deliveryFolder?: string;
  /** POP3: deixar uma cópia das mensagens no servidor (padrão true — seguro) */
  leaveOnServer?: boolean;
  /** POP3: remover do servidor após N dias (0 = nunca remover) */
  leaveDays?: number;
  /** Aceitar certificado TLS autoassinado/não confiável (servidores próprios
   *  e hospedagens comuns). Vale para entrada (IMAP/POP3) e saída (SMTP). */
  allowSelfSigned?: boolean;
}

/** Assinatura nomeada, editável em "Assinaturas" (estilo Arquivo > Opções > Email > Assinaturas) */
export interface Signature {
  id: string;
  name: string;
  html: string;
}

/** Dados usados para criar/editar/testar uma conta; senha só trafega nesta direção (renderer -> main) */
export interface AccountDraft extends Omit<AccountConfig, 'id'> {
  id?: string;
  password: string;
  /** Senha de SMTP separada (quando difere da senha de entrada). */
  smtpPassword?: string;
}

export interface AccountTestResult {
  ok: boolean;
  message: string;
  /** O teste descobriu a criptografia do servidor de entrada invertida —
   *  o formulário deve aplicar este valor em `imapSecure`. */
  imapSecureFix?: boolean;
  /** O servidor usa certificado autoassinado — só conectou aceitando-o;
   *  o formulário deve marcar `allowSelfSigned`. */
  allowSelfSignedFix?: boolean;
}

/** Palpite extraído do Registro do Windows */
export interface ImportedAccountGuess {
  email: string;
  displayName: string;
  protocol: 'imap' | 'pop3';
  imapHost: string;
  imapPort: number;
  imapSecure: boolean;
  smtpHost: string;
  smtpPort: number;
  smtpSecure: boolean;
  /** Caminho do PST de entrega encontrado no perfil do Outlook */
  pstPath?: string;
  /** Conteúdo HTML da assinatura configurada no Outlook, se encontrada */
  signatureHtml?: string;
  /** Assinatura de resposta/encaminhamento do Outlook, se diferente */
  signatureReplyHtml?: string;
  /** Senha de entrada (POP3/IMAP) recuperada via DPAPI do Registro do Windows — mesmo usuário Windows apenas. */
  password?: string;
  /** Senha de SMTP, se configurada separadamente da de entrada. */
  smtpPassword?: string;
  /** Usuário de entrada (POP3/IMAP) — nem sempre é o endereço de e-mail. */
  imapUser?: string;
  /** Usuário de SMTP, se configurado separadamente. */
  smtpUser?: string;
  /** POP3: manter cópia no servidor (bitmask "Leave on Server" do Registro). */
  leaveOnServer?: boolean;
  /** POP3: remover do servidor após N dias (0 = nunca). */
  leaveDays?: number;
}

/** Pasta de email dentro de um perfil do Thunderbird (arquivo .mbox). */
export interface TbMailFolder {
  name: string;
  /** Caminho do arquivo .mbox no disco. */
  mbox: string;
  children: TbMailFolder[];
}

/** "Conta" de armazenamento do Thunderbird (Local Folders, POP, IMAP). */
export interface TbMailStore {
  label: string;
  folders: TbMailFolder[];
}

/** Perfil do Thunderbird com os armazenamentos de email encontrados. */
export interface TbMailProfile {
  path: string;
  name: string;
  stores: TbMailStore[];
}

export interface ComposeInput {
  accountId: string;
  /** Tipo de mensagem — decide qual assinatura anexar (nova vs resposta) */
  kind?: 'new' | 'reply' | 'forward';
  to: string;
  cc?: string;
  bcc?: string;
  subject: string;
  html: string;
  inReplyTo?: string;
  references?: string;
  /** Caminhos de arquivos locais a anexar */
  attachmentPaths?: string[];
  /** Anexos gerados em memória (ex.: resposta de convite .ics) */
  rawAttachments?: Array<{ filename: string; contentBase64: string; contentType: string }>;
  /** Não anexar assinatura da conta (convites de reunião, msgs automáticas) */
  noSignature?: boolean;
  /** Convite iMIP — envia o evento como alternativa text/calendar (RSVP nativo). */
  icalEvent?: { method: string; contentBase64: string };
  /** Salvar como rascunho — marca X-Unsent: 1 no .eml gravado. */
  draft?: boolean;
  /** Item original de uma resposta/encaminhamento — após enviar, o app
   *  marca \Answered/$Forwarded no servidor (ícone ↩/→ como no Outlook). */
  replyRef?: { sourceId: string; folderId: string; index: number };
}

/** Resultado da criação de compromisso/contato/tarefa + envio de convites. */
export interface CreateSpecialResult {
  tree: FolderNode[];
  /** null quando o item não é compromisso ou não tem participantes. */
  invite: {
    sent: number;
    failed: number;
    /** Motivo do não envio — ex.: nenhuma conta configurada. */
    error?: string;
    noAccount?: boolean;
  } | null;
}

/** Rascunho passado à janela de composição separada (IPC compose:openWindow). */
export interface ComposeDraftPayload {
  accountId: string;
  kind?: 'new' | 'reply' | 'forward';
  to: string;
  cc?: string;
  bcc?: string;
  subject: string;
  bodyText?: string;
  bodyHtml?: string;
  /** Mensagem original citada (resposta/encaminhamento) — abaixo da assinatura */
  quoteHtml?: string;
  inReplyTo?: string;
  references?: string;
  /** Anexos do rascunho — caminhos de arquivo (rascunho reaberto usa cópias temp). */
  attachmentPaths?: string[];
  /** Referência do rascunho salvo — após enviar/regravar, o .eml original sai
   *  de Rascunhos (como o Outlook move o rascunho ao enviar). */
  draftRef?: { sourceId: string; folderId: string; index: number };
}

// ===================== Regras de recebimento =====================

/**
 * Regra estilo Outlook: quando chegar email na conta X que combine as
 * condições, arquiva uma cópia (.eml) na pasta local vinculada ao PST alvo.
 * Todas as condições preenchidas devem combinar (AND). Regra sem condições
 * combina todas as mensagens.
 */
export interface MailRule {
  id: string;
  name: string;
  /** id da conta (AccountConfig.id) ou 'all' para qualquer conta */
  accountId: string;
  enabled: boolean;
  conditions: {
    fromContains?: string;
    toContains?: string;
    subjectContains?: string;
  };
  /** Caminho do arquivo PST que "recebe" os emails (a pasta local fica atrelada a ele) */
  targetPstPath: string;
  /** Nome da pasta de destino dentro do arquivo local do PST */
  targetFolder: string;
  /** "Parar de processar mais regras" — ação padrão do Outlook após mover */
  stopProcessing?: boolean;
}

export interface SyncResult {
  accountId: string;
  scanned: number;
  archived: number;
  /** Mensagens detectadas como infectadas pelo antivírus local e movidas para "Itens infectados". */
  infected?: number;
  /** Mensagens com marcação de spam do servidor movidas para "Lixo eletrônico". */
  junked?: number;
  /** IMAP sem entrega local: mensagens novas vistas no servidor (só aviso/badge). */
  newOnServer?: number;
  error?: string;
}

/** Uma mensagem processada num envio/recebimento (linha da lista de progresso). */
export interface SyncItem {
  key: string;
  subject: string;
  from: string;
  status: 'ok' | 'fail';
  time: string;
}

/** Atualização de progresso de uma tarefa de envio/recebimento. */
export interface SyncProgress {
  /** Id estável da tarefa (ex.: "acct123:recv"). */
  taskId: string;
  /** Rótulo exibido: "email@x.com - Recebendo". */
  label: string;
  accountId: string;
  phase: 'send' | 'recv';
  status: 'pending' | 'processing' | 'done' | 'error' | 'cancelled';
  total?: number;
  done?: number;
  error?: string;
  /** Mensagem processada neste evento (o renderer acumula em `items`). */
  item?: SyncItem;
  /** Todas as mensagens processadas até agora nesta tarefa (montado no renderer). */
  items?: SyncItem[];
}
