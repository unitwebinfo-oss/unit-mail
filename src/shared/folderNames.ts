// Nomes de pastas padrão por idioma + mapeamento semântico.
//
// REGRA DE OURO: nunca renomeamos pastas já criadas no disco. O idioma decide
// o nome das pastas NOVAS e como as pastas conhecidas são EXIBIDAS —
// "Caixa de Entrada" criada em pt continua "Caixa de Entrada" no disco e
// aparece como "Inbox" quando a interface está em inglês.

export type Lang = 'pt' | 'en';

/** Tipos semânticos de pasta padrão (estilo Outlook). */
export type FolderKind =
  | 'inbox'
  | 'sent'
  | 'deleted'
  | 'quarantine'
  | 'drafts'
  | 'outbox'
  | 'junk'
  | 'archive'
  | 'calendar'
  | 'contacts'
  | 'tasks';

interface FolderKindDef {
  pt: string;
  en: string;
  /**
   * Todos os nomes que identificam esta pasta, em qualquer idioma/formato
   * já visto no disco (pt-BR, en-US e variantes comuns de Outlook/Thunderbird).
   * Comparados em minúsculas e sem acentos.
   */
  aliases: string[];
}

const KINDS: Record<FolderKind, FolderKindDef> = {
  inbox: {
    pt: 'Caixa de Entrada',
    en: 'Inbox',
    aliases: ['caixa de entrada', 'inbox', 'caixa postal', 'caixadeentrada'],
  },
  sent: {
    pt: 'Itens Enviados',
    en: 'Sent Items',
    aliases: ['itens enviados', 'sent items', 'sent', 'enviados', 'itens enviado'],
  },
  deleted: {
    pt: 'Itens Excluídos',
    en: 'Deleted Items',
    aliases: ['itens excluidos', 'deleted items', 'deleted', 'trash', 'lixeira', 'itens apagados', 'itens eliminados', 'excluidos', 'eliminados'],
  },
  quarantine: {
    pt: 'Itens infectados',
    en: 'Infected Items',
    aliases: ['itens infectados', 'infected items', 'quarantine', 'quarentena'],
  },
  drafts: {
    pt: 'Rascunhos',
    en: 'Drafts',
    aliases: ['rascunhos', 'drafts', 'draft'],
  },
  outbox: {
    pt: 'Caixa de Saída',
    en: 'Outbox',
    aliases: ['caixa de saida', 'outbox', 'caixa de saída'],
  },
  junk: {
    pt: 'Lixo Eletrônico',
    en: 'Junk Email',
    aliases: ['lixo eletronico', 'lixo eletrônico', 'junk email', 'junk', 'spam', 'junk e-mail', 'lixo eletrnico'],
  },
  archive: {
    pt: 'Arquivo Morto',
    en: 'Archive',
    aliases: ['arquivo morto', 'arquivo', 'archive', 'arquivados', 'archives'],
  },
  calendar: {
    pt: 'Calendário',
    en: 'Calendar',
    aliases: ['calendario', 'calendário', 'calendar', 'agenda', 'calendarios', 'calendários', 'calendars'],
  },
  contacts: {
    pt: 'Contatos',
    en: 'Contacts',
    aliases: ['contatos', 'contacts', 'contactos'],
  },
  tasks: {
    pt: 'Tarefas',
    en: 'Tasks',
    aliases: ['tarefas', 'tasks', 'to-do', 'to do'],
  },
};

/** Minúsculas + sem acentos — comparar nomes de pasta independente de grafia. */
function norm(name: string): string {
  return (name || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .trim();
}

// Índice alias-normalizado → tipo, montado uma vez.
const ALIAS_INDEX = new Map<string, FolderKind>();
for (const [kind, def] of Object.entries(KINDS) as Array<[FolderKind, FolderKindDef]>) {
  for (const a of [def.pt, def.en, ...def.aliases]) {
    const n = norm(a);
    if (n && !ALIAS_INDEX.has(n)) ALIAS_INDEX.set(n, kind);
  }
}

/** Tipo semântico da pasta pelo nome físico (null = pasta comum do usuário). */
export function folderKindOf(name: string): FolderKind | null {
  return ALIAS_INDEX.get(norm(name)) ?? null;
}

/** Nome canônico da pasta no idioma — usado ao CRIAR pastas novas. */
export function folderName(kind: FolderKind, lang: Lang): string {
  return KINDS[kind][lang];
}

/**
 * Nome de EXIBIÇÃO: pastas padrão aparecem no idioma da interface; pastas do
 * usuário (e nomes de PST vindos do arquivo) ficam como estão.
 */
export function localizeFolderName(name: string, lang: Lang): string {
  const kind = folderKindOf(name);
  return kind ? KINDS[kind][lang] : name;
}

/** Nomes das pastas padrão criadas numa base nova, no idioma escolhido. */
export function defaultFolderNames(lang: Lang): string[] {
  const order: FolderKind[] = [
    'inbox',
    'sent',
    'deleted',
    'quarantine',
    'drafts',
    'outbox',
    'junk',
    'archive',
    'calendar',
    'contacts',
    'tasks',
  ];
  return order.map((k) => KINDS[k][lang]);
}

/** Todas as grafias conhecidas do tipo — para achar a pasta já existente. */
export function folderAliases(kind: FolderKind): string[] {
  const def = KINDS[kind];
  return [def.pt, def.en, ...def.aliases];
}

/**
 * Resolve o nome físico a usar para um tipo de pasta numa base:
 * se JÁ existe uma pasta desse tipo (em qualquer idioma), retorna o nome
 * existente — nunca cria "Inbox" ao lado de "Caixa de Entrada". Só quando
 * não existe nenhuma variante usa o nome do idioma atual.
 */
export function resolveFolderName(existingNames: string[], kind: FolderKind, lang: Lang): string {
  const wanted = ALIAS_INDEX;
  for (const n of existingNames) {
    if (wanted.get(norm(n)) === kind) return n;
  }
  return KINDS[kind][lang];
}
