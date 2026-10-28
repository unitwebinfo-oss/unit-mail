import { PSTFile, PSTFolder, PSTMessage, PSTAttachment } from 'pst-extractor';
// subpath interno da lib — PSTUtil não é reexportado pelo index
import { PSTUtil } from 'pst-extractor/dist/PSTUtil.class';
import iconv from 'iconv-lite';
import { simpleParser } from 'mailparser';
import { basename, join } from 'path';
import { createHash, randomUUID } from 'crypto';
import { closeSync, existsSync, mkdirSync, openSync, readdirSync, readFileSync, readSync, renameSync, rmSync, statSync, writeFileSync } from 'fs';
import { isDefaultFolderId } from '../shared/types';
import { normSearch } from '../shared/textNorm';
import { safeFsName } from '../shared/fsName';
import { resolveFolderName, folderKindOf, type FolderKind } from '../shared/folderNames';
import { appLanguage } from './language';
import type {
  AttachmentData,
  AttachmentMeta,
  ContactDetail,
  FolderNode,
  ItemDetail,
  ItemSummary,
  NewSpecialItem,
  PstOpenResult,
  SearchResult,
  TaskDetail,
} from '../shared/types';

// pst-extractor decodifica strings PT_STRING8 com a codepage da mensagem;
// quando a mensagem não declara PidTagMessageCodepage/InternetCodepage cai
// em 'utf8' e bytes ANSI (cp1252) viram U+FFFD ("Ol�") — perda irreversível.
// Interceptamos na fonte: resultado com U+FFFD → re-decodifica os bytes
// originais como windows-1252 (todo byte é válido e cobre o latim/pt-BR).
const origCreateJsString = PSTUtil.createJavascriptString.bind(PSTUtil);
PSTUtil.createJavascriptString = function (
  data: Buffer,
  stringType: number,
  codepage?: string,
): string {
  const s = origCreateJsString(data, stringType, codepage);
  if (s.includes('�') && stringType !== 0x1f && Buffer.isBuffer(data)) {
    try {
      const fixed = iconv.decode(data, 'windows1252');
      return fixed.includes('�') ? s : fixed;
    } catch {
      return s;
    }
  }
  return s;
};

// libmime vem embutido no nodemailer — decodeWords trata =?charset?B/Q?...?=
// eslint-disable-next-line @typescript-eslint/no-var-requires
const libmime = require('libmime') as { decodeWords: (s: string) => string };

/**
 * Header .eml com U+FFFD ("Ol�"): o remetente gravou texto ANSI (cp1252)
 * cru no cabeçalho, sem encoded-word — o mailparser assume utf8 e perde os
 * bytes. Recupera da linha crua (headerLines preserva latin1) → cp1252 →
 * decodeWords para as partes codificadas, se houver.
 */
function fixDecodedHeader(decoded: string, rawLine: string | undefined): string {
  if (!decoded || !decoded.includes('�') || !rawLine) return decoded;
  try {
    const s = libmime.decodeWords(iconv.decode(Buffer.from(rawLine, 'latin1'), 'windows1252'));
    return s.includes('�') || !s.trim() ? decoded : s;
  } catch {
    return decoded;
  }
}

// MAPI named properties para tarefas (PSETID_Task)
const PR_TASK_START_DATE = 0x8104;
const PR_TASK_DUE_DATE = 0x8105;

// MAPI named properties para contatos (PSETID_Address)
const PR_CONTACT_HOME_ADDRESS = 0x801a;
const PR_CONTACT_BUSINESS_ADDRESS = 0x801b;

function toIso(d: Date | null | undefined): string | null {
  return d instanceof Date && !isNaN(d.getTime()) ? d.toISOString() : null;
}

/**
 * Pastas locais não guardam a classe MAPI no disco — inferimos pelo nome,
 * cobrindo as pastas padrão do Outlook em pt-BR e en-US.
 */
export function containerClassForFolderName(name: string): string {
  const n = name.trim().toLowerCase();
  if (/^(calend[aá]rio|calendar)$/.test(n)) return 'IPF.Appointment';
  if (/^(contatos|contacts)$/.test(n)) return 'IPF.Contact';
  if (/^(tarefas|tasks)$/.test(n)) return 'IPF.Task';
  if (/^(notas|notes)$/.test(n)) return 'IPF.StickyNote';
  if (/^(di[aá]rio|journal)$/.test(n)) return 'IPF.Journal';
  return 'IPF.Note';
}

/** Converte data iCalendar (20250101T120000Z ou flutuante) para ISO. */
function icsDate(v: string): string | null {
  const m = v.trim().match(/(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?/);
  if (!m) return null;
  const [, y, mo, d, h, mi, s, z] = m;
  const iso = `${y}-${mo}-${d}T${h || '00'}:${mi || '00'}:${s || '00'}`;
  const date = z ? new Date(`${iso}Z`) : new Date(iso);
  return isNaN(date.getTime()) ? null : date.toISOString();
}

/** Nomes de fuso do Windows (como o Outlook grava em TZID) → IANA. */
const WIN_TZ: Record<string, string> = {
  'E. South America Standard Time': 'America/Sao_Paulo',
  'Central Brazilian Standard Time': 'America/Cuiaba',
  'SA Pacific Standard Time': 'America/Bogota',
  'Pacific Standard Time': 'America/Los_Angeles',
  'Mountain Standard Time': 'America/Denver',
  'Central Standard Time': 'America/Chicago',
  'Eastern Standard Time': 'America/New_York',
  'GMT Standard Time': 'Europe/London',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Romance Standard Time': 'Europe/Paris',
  'UTC': 'UTC',
};

/** Deslocamento (ms) de um fuso IANA em um instante. */
function zoneOffsetMs(zone: string, at: number): number {
  const p: Record<string, number> = {};
  for (const x of new Intl.DateTimeFormat('en-US', {
    timeZone: zone,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(new Date(at))) {
    if (x.type !== 'literal') p[x.type] = Number(x.value);
  }
  return Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second) - Math.floor(at / 1000) * 1000;
}

/** Valor de data iCalendar com TZID → instante ISO (hora de parede no fuso informado). */
function icsDateInZone(v: string, tzid: string): string | null {
  const m = v.trim().match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?)?$/);
  if (!m) return null;
  const [, y, mo, d, h, mi, s] = m;
  const guess = Date.UTC(+y, +mo - 1, +d, +(h || 0), +(mi || 0), +(s || 0));
  try {
    const zone = WIN_TZ[tzid] ?? tzid;
    let off = zoneOffsetMs(zone, guess);
    let t = guess - off;
    const off2 = zoneOffsetMs(zone, t);
    if (off2 !== off) t = guess - off2;
    return new Date(t).toISOString();
  } catch {
    // fuso desconhecido — trata como hora local da máquina
    return icsDate(v);
  }
}

/** Parser mínimo de iCalendar — extrai o primeiro VEVENT + METHOD. */
export function parseIcsInvite(text: string): {
  method: string;
  uid: string;
  summary: string;
  start: string | null;
  end: string | null;
  location: string;
  organizer: string;
  organizerEmail: string;
  reminderMinutes?: number;
} | null {
  if (!/BEGIN:VEVENT/i.test(text)) return null;
  // "Unfold" — linhas continuadas começam com espaço/tab
  const t = text.replace(/\r?\n[ \t]/g, '');
  // Só o VEVENT: o VTIMEZONE também tem DTSTART (1601) e seria lido por engano
  const ev = t.match(/BEGIN:VEVENT([\s\S]*?)END:VEVENT/i)?.[1] ?? t;
  const get = (name: string, src = ev): string =>
    src.match(new RegExp(`^${name}(?:;[^:\\r\\n]*)?:(.*)$`, 'im'))?.[1]?.trim() ?? '';
  const date = (name: string): string | null => {
    const m = ev.match(new RegExp(`^${name}((?:;[^:\\r\\n]*)?):(.*)$`, 'im'));
    if (!m) return null;
    const value = m[2].trim();
    const tzid = m[1].match(/TZID="?([^";:]+)"?/i)?.[1];
    return tzid && !value.endsWith('Z') ? icsDateInZone(value, tzid) : icsDate(value);
  };
  const orgRaw = get('ORGANIZER');
  const orgEmail = orgRaw.match(/mailto:([^\s>]+)/i)?.[1] ?? '';
  const orgName = ev.match(/ORGANIZER[^:\r\n]*CN="?([^";:\r\n]+)"?/i)?.[1]?.trim() ?? '';
  const trig = ev.match(/TRIGGER[^:\r\n]*:-?PT?(\d+)M/i)?.[1];
  return {
    method: (get('METHOD', t) || 'REQUEST').toUpperCase(),
    uid: get('UID'),
    summary: get('SUMMARY'),
    start: date('DTSTART'),
    end: date('DTEND'),
    location: get('LOCATION'),
    organizer: orgName || orgEmail,
    organizerEmail: orgEmail,
    reminderMinutes: trig ? Number(trig) : 15,
  };
}

/** Versão do índice de busca — bump quando o texto indexado muda. */
const SEARCH_IDX_VER = 2;

export class PstService {
  private pst: PSTFile | null = null;
  private localOnly = false;
  private filePath = '';
  private folders = new Map<string, PSTFolder>();
  private rootNodes: FolderNode[] = [];
  /** Nós vindos do PST antes da mesclagem com as pastas-sombra. */
  private pstNodes: FolderNode[] = [];

  /**
   * Diretório de arquivamento local vinculado a este PST. Como o formato PST
   * é somente-leitura nas bibliotecas abertas, emails recebidos por regras
   * são guardados como .eml em subpastas deste diretório e aparecem na
   * árvore como pastas "arc:<nome>" dentro do nó deste PST.
   */
  constructor(private archiveDir?: string) {}

  /**
   * Profundidade de lote — dentro de batch() as gravações de estado em disco
   * (readstate.json/hidden.json) e os recounts da árvore são adiados para uma
   * única gravação no fim. Sem isso, excluir 700 itens regravava o JSON
   * inteiro e re-lera a árvore 700 vezes (custo quadrático).
   */
  private batchDepth = 0;

  /** Executa mutações em lote: persiste estados e reconstrói a árvore uma vez. */
  async batch<T>(fn: () => T | Promise<T>): Promise<T> {
    this.batchDepth++;
    try {
      return await fn();
    } finally {
      if (--this.batchDepth === 0) {
        this.saveReadState();
        this.saveHidden();
        this.saveItemCats();
        this.sumCache.clear();
        this.rebuildTree();
      }
    }
  }

  /**
   * Data da mensagem extraída do cabeçalho Date: do .eml bruto — usada como
   * prefixo do nome do arquivo, porque a listagem ordena por nome (mais novo
   * primeiro). Importações antigas usavam Date.now() (hora da importação),
   * o que embaralhava a ordem cronológica — repairEmlNames corrige.
   */
  private emlDateMs(raw: Buffer): number | null {
    const head = raw.slice(0, 8192).toString('utf8');
    const m = /^Date:\s*(.+?)\r?$/im.exec(head.replace(/\r?\n[ \t]+/g, ' '));
    const t = m ? Date.parse(m[1]) : NaN;
    return isNaN(t) ? null : t;
  }

  /** Lê só os primeiros 8 KB do arquivo — suficiente para o cabeçalho Date:. */
  private emlDateMsFromFile(fp: string): number | null {
    const fd = openSync(fp, 'r');
    try {
      const buf = Buffer.alloc(8192);
      const n = readSync(fd, buf, 0, 8192, 0);
      return this.emlDateMs(buf.slice(0, n));
    } finally {
      closeSync(fd);
    }
  }

  /**
   * Reparo único por base: arquivos .eml gravados com o prefixo "hora da
   * importação" são renomeados para "<data da mensagem>-<rand>.eml".
   * O marker '.dates-fixed' evita rodar de novo. O readstate.json é remapeado
   * porque suas chaves contêm o nome do arquivo.
   */
  private repairEmlNames(): void {
    const root = this.archiveDir;
    if (!root || !existsSync(root)) return;
    if (existsSync(join(root, '.dates-fixed'))) return;
    const st = this.loadReadState();
    let changed = false;
    const walk = (rel: string) => {
      const dir = rel ? join(root, ...rel.split('/')) : root;
      const lst = this.listDir(dir);
      if (!lst) return;
      for (const sub of lst.dirs) walk(rel ? `${rel}/${sub}` : sub);
      for (const name of lst.emls) {
        const r = rel ? `${rel}/${name}` : name;
        const fp = join(dir, name);
        try {
          const ms = this.emlDateMsFromFile(fp);
          if (ms === null) continue; // sem data na mensagem — mantém nome atual
          const cur = Number(name.split('-')[0]);
          // Já está certo se o prefixo bate com a data da mensagem
          if (cur === ms) continue;
          const newName = `${ms}-${randomUUID().slice(0, 8)}.eml`;
          renameSync(fp, join(dir, newName));
          const key = r;
          const was = st.get(key);
          if (was !== undefined) {
            st.delete(key);
            st.set(rel ? `${rel}/${newName}` : newName, was);
          }
          changed = true;
        } catch {
          // arquivo ilegível — pula
        }
      }
    };
    try {
      walk('');
      if (changed) this.saveReadState();
      writeFileSync(join(root, '.dates-fixed'), '1');
      this.sumCache.clear();
      this.fsChanged();
    } catch {
      // reparo best-effort
    }
  }

  get isOpen(): boolean {
    return this.pst !== null || this.localOnly;
  }

  get currentPath(): string {
    return this.filePath;
  }

  open(filePath: string): PstOpenResult {
    this.close();
    this.pst = new PSTFile(filePath);
    this.filePath = filePath;
    this.folders.clear();
    this.repairEmlNames();
    this.rootNodes = this.buildTree();
    this.reindexNodes();
    let storeName = basename(filePath);
    try {
      const ms = this.pst.getMessageStore();
      const dn = (ms as unknown as { displayName?: string }).displayName;
      if (dn) storeName = dn;
    } catch {
      // mantém o nome do arquivo
    }
    return { fileName: basename(filePath), storeName, filePath };
  }

  close(): void {
    if (this.sumIdxTimer) {
      clearTimeout(this.sumIdxTimer);
      this.sumIdxTimer = null;
      this.saveSumIdx();
    }
    this.sumIdx = null;
    if (this.pst) {
      try {
        this.pst.close();
      } catch {
        // ignora
      }
    }
    this.pst = null;
    this.localOnly = false;
    this.folders.clear();
    this.rootNodes = [];
    this.pstNodes = [];
    this.nodeIndex = null;
    this.dirCache.clear();
    this.filePath = '';
    this.sumCache.clear();
  }

  /**
   * Abre uma base de dados local (sem PST): as subpastas de storeDir
   * contêm mensagens .eml e são exibidas como pastas da base.
   */
  openLocal(name: string, storeDir: string): PstOpenResult {
    this.close();
    this.localOnly = true;
    this.archiveDir = storeDir;
    this.filePath = storeDir;
    this.repairEmlNames();
    this.rebuildTree();
    return { fileName: name, storeName: name, filePath: storeDir };
  }

  getFolderTree(): FolderNode[] {
    this.requireOpen();
    // Re-mescla sempre: o Enviar/Receber e movimentações podem ter criado
    // diretórios .eml novos desde a abertura da base. As listagens vêm do
    // dirCache (mtime-keyed) — fora a 1ª vez, o custo é um stat por pasta.
    this.rebuildTree();
    return this.rootNodes;
  }

  /** Reconstrói a árvore mesclada e o índice id→nó. */
  private rebuildTree(): void {
    this.rootNodes = this.pst ? this.mergeArchiveNodes(this.pstNodes) : this.archiveFolderNodes();
    this.reindexNodes();
  }

  /** (Re)monta o índice id→nó da árvore atual. */
  private reindexNodes(): void {
    this.nodeIndex = new Map();
    const idx = (nodes: FolderNode[]) => {
      for (const n of nodes) {
        this.nodeIndex!.set(n.id, n);
        idx(n.children);
      }
    };
    idx(this.rootNodes);
  }

  private buildTree(): FolderNode[] {
    const root = this.pst!.getRootFolder();
    // Pastas internas que o Outlook não exibe: nó raiz de pesquisa e as
    // pooled search folders MS-OLK-* (criadas pelo Outlook p/ buscas).
    const isInternal = (folder: PSTFolder): boolean => {
      const n = (folder.displayName || '').trim();
      const cc = folder.containerClass || '';
      return (
        /^MS-OLK-/i.test(n) ||
        /^(IPM_VIEWS|IPM_COMMON_VIEWS|IPM_SUBTREE|Common Views)$/i.test(n) ||
        /^(pesquisar raiz|raiz de pesquisa|search root|finder)$/i.test(n) ||
        // Pastas de sistema ocultas no Outlook (buscas automáticas, dados de
        // disponibilidade, lembretes, configurações internas)
        /^SPAM Search Folder/i.test(n) ||
        /^(ItemProcSearch|Contact Search|Subscriptions|News Feed)$/i.test(n) ||
        /^Free ?busy( Data)?$/i.test(n) ||
        /^(Lembretes|Reminders|Pesquisa de Itens Pendentes|To-Do Search)$/i.test(n) ||
        /^Processamento de Email/i.test(n) ||
        /^(Conversation Action Settings|Quick Step Settings|PersonMetadata)$/i.test(n) ||
        /^(Recipient Cache|GAL Contacts|externalContacts|Deletions|Recoverable Items|Purges|Versions|Audits|Calendar Logging)$/i.test(n) ||
        cc === 'IPF.Search' ||
        cc === 'IPF.Configuration'
      );
    };
    const build = (folder: PSTFolder, id: string): FolderNode => {
      this.folders.set(id, folder);
      const children: FolderNode[] = [];
      if (folder.hasSubfolders) {
        folder.getSubFolders().forEach((sub, i) => {
          if (isInternal(sub)) return;
          children.push(build(sub, id ? `${id}/${i}` : `${i}`));
        });
      }
      return {
        id,
        name: folder.displayName || '(sem nome)',
        containerClass: folder.containerClass || '',
        contentCount: folder.contentCount || 0,
        unreadCount: (() => {
          const ro = this.pstReadOverrideCounts(id);
          return Math.max(0, (folder.unreadCount || 0) - ro.read) + ro.unread;
        })(),
        hidden: this.isFolderHidden(id) || undefined,
        children,
      };
    };
    const rootNode = build(root, '');
    let top = rootNode.children;
    // Alguns PSTs têm um contêiner intermediário ("Raiz", "Top of Personal
    // Folders") que só envolve as pastas reais — o Outlook o oculta. Se um
    // nó de topo é um contêiner sem conteúdo próprio, promove seus filhos.
    const isWrapper = (n: FolderNode) =>
      n.children.length > 0 &&
      (/^(raiz|root|top of|personal folders|dados do outlook|início do arquivo de dados|inicio do arquivo de dados|top of outlook data file)/i.test(
        n.name,
      ) ||
        (n.containerClass === '' && n.contentCount === 0));
    while (top.length === 1 && isWrapper(top[0])) {
      top = top[0].children;
    }
    this.pstNodes = top;
    return this.mergeArchiveNodes(top);
  }

  /**
   * Funde as pastas-sombra (.eml) nas pastas reais do PST com o mesmo nome —
   * como o Outlook, que entrega POP3 direto nas pastas do arquivo. Pastas
   * locais sem correspondente no PST continuam visíveis como nó próprio.
   */
  private mergeArchiveNodes(pstNodes: FolderNode[]): FolderNode[] {
    const arcNodes = this.archiveFolderNodes();
    if (arcNodes.length === 0) return pstNodes;
    const findByName = (nodes: FolderNode[], name: string): FolderNode | null => {
      for (const n of nodes) {
        if (n.name.toLowerCase() === name.toLowerCase()) return n;
        const hit = findByName(n.children, name);
        if (hit) return hit;
      }
      return null;
    };
    /** Funde um nó-sombra num nó do PST: marca archiveName, soma contagens e
     *  mescla os filhos recursivamente (subpastas da sombra com mesmo nome
     *  entram na subpasta real em vez de virarem pastas soltas). */
    const mergeInto = (match: FolderNode, arc: FolderNode) => {
      let dirName = this.canonShadow(arc.id.slice(4)); // remove "arc:" + remapeia
      if (match.archiveName && match.archiveName !== dirName) {
        // Duas sombras apontam para a mesma pasta do PST (ex.: regra antiga com
        // caminho aninhado "Início.../Pasta"). Sem isso a segunda sobrescrevia
        // archiveName e os arquivos da primeira ficavam invisíveis. Consolida
        // fisicamente: move os .eml para o diretório já vinculado.
        this.consolidateShadow(match.archiveName, dirName);
        this.shadowMap.set(arc.id.slice(4), match.archiveName);
        dirName = match.archiveName;
        this.fsChanged();
      }
      match.archiveName = dirName;
      const own = this.archiveFiles(`arc:${dirName}`);
      const st = this.loadReadState();
      // Baseline PST capturada na 1ª mesclagem — sem ela os contadores
      // acumulariam a cada refresh da árvore
      const mm = match as FolderNode & { _bc?: number; _bu?: number };
      if (mm._bc === undefined) {
        mm._bc = this.folders.get(match.id)?.contentCount ?? match.contentCount;
        mm._bu = this.folders.get(match.id)?.unreadCount ?? match.unreadCount;
      }
      // Desconta os itens do PST ocultos (exclusão lógica)
      match.contentCount = Math.max(0, (mm._bc ?? 0) - this.hiddenSet(match.id).size) + own.length;
      const ro = this.pstReadOverrideCounts(match.id);
      match.unreadCount =
        Math.max(0, (mm._bu ?? 0) - this.hiddenUnreadCount(match.id) - ro.read) +
        ro.unread +
        own.filter((f) => st.get(`${dirName}/${f}`) !== true).length;
      // Filhos do PST (nunca "arc:") + filhos da sombra mesclados por nome
      match.children = match.children.filter((c) => !c.id.startsWith('arc:'));
      for (const sub of arc.children) {
        const subMatch = match.children.find(
          (c) => c.name.toLowerCase() === sub.name.toLowerCase(),
        );
        if (subMatch) mergeInto(subMatch, sub);
        else match.children.push(sub);
      }
    };
    const out = [...pstNodes];
    for (const arc of arcNodes) {
      const match = findByName(out, arc.name);
      if (match) {
        mergeInto(match, arc);
        continue;
      }
      // Sombra sem pasta correspondente: pode ser um contêiner que reproduz a
      // estrutura ("Início do arquivo de dados…/Agencia Unit"). Tenta dissolver:
      // mescla cada filho pelo nome — se todos casam e a pasta está vazia, ela
      // não aparece como nó solto na árvore.
      const own = this.archiveFiles(arc.id);
      if (arc.children.length > 0 && own.length === 0) {
        const left: FolderNode[] = [];
        for (const sub of arc.children) {
          const subMatch = findByName(out, sub.name);
          if (subMatch) mergeInto(subMatch, sub);
          else left.push(sub);
        }
        if (left.length === 0) {
          // Contêiner dissolvido por completo — remove a estrutura-sombra se
          // não restou nenhum .eml dentro (senão viraria pasta fantasma).
          const dir = this.archiveDirFor(arc.id);
          const stillHasEml = (d: string): boolean => {
            const lst = this.listDir(d);
            return !!lst && (lst.emls.length > 0 || lst.dirs.some((s) => stillHasEml(join(d, s))));
          };
          try {
            if (!stillHasEml(dir)) {
              rmSync(dir, { recursive: true, force: true });
              this.fsChanged();
            }
          } catch {
            // limpeza best-effort
          }
          continue;
        }
        arc.children = left;
      }
      out.push(arc);
    }
    // Nós PST não mesclados: recalcula contagens a cada refresh — os badges
    // de não-lidos/ocultos mudam conforme o usuário lê e exclui itens, e a
    // contagem original de buildTree ficaria congelada desde a abertura.
    const fixPstCounts = (nodes: FolderNode[]) => {
      for (const n of nodes) {
        if (!n.id.startsWith('arc:') && !n.archiveName) {
          const f = this.folders.get(n.id);
          if (f) {
            const hidden = this.hiddenSet(n.id).size;
            n.contentCount = Math.max(0, (f.contentCount || 0) - hidden);
            const ro = this.pstReadOverrideCounts(n.id);
            n.unreadCount =
              Math.max(0, (f.unreadCount || 0) - this.hiddenUnreadCount(n.id) - ro.read) + ro.unread;
          }
        }
        fixPstCounts(n.children);
      }
    };
    fixPstCounts(out);
    return out;
  }

  /** Índice id→nó — nodeFor era uma varredura recursiva da árvore por chamada. */
  private nodeIndex: Map<string, FolderNode> | null = null;

  /** Localiza um nó da árvore por id (necessário para pastas mescladas). */
  private nodeFor(folderId: string): FolderNode | null {
    if (!this.nodeIndex) this.rebuildTree();
    const hit = this.nodeIndex!.get(folderId);
    if (hit) return hit;
    // Nó recém-criado fora do rebuild (ex.: getSubFolders tardio) — fallback
    const walk = (nodes: FolderNode[]): FolderNode | null => {
      for (const n of nodes) {
        if (n.id === folderId) return n;
        const f = walk(n.children);
        if (f) return f;
      }
      return null;
    };
    return walk(this.rootNodes);
  }

  /**
   * Diretório gravável de uma pasta: "arc:" direto, ou pasta do PST que já
   * tem (ou passa a ter) diretório-sombra. null se a base não tem arquivamento.
   */
  writableDirFor(folderId: string): string | null {
    if (!this.archiveDir) return null;
    if (folderId.startsWith('arc:')) {
      // Garante o diretório em disco — destinos como "arc:Itens Excluídos"
      // podem não existir ainda, e gravar neles falhava com ENOENT.
      const dir = this.archiveDirFor(folderId);
      mkdirSync(dir, { recursive: true });
      this.fsChanged();
      return dir;
    }
    const node = this.nodeFor(folderId);
    if (node && node.containerClass === 'IPF.Note') {
      return this.ensureArchiveDirFor(node);
    }
    return null;
  }

  /** Garante o diretório-sombra de uma pasta do PST (cria e marca o nó). */
  private ensureArchiveDirFor(node: FolderNode): string {
    if (!this.archiveDir) throw new Error('Base sem diretório de arquivamento');
    if (!node.archiveName) {
      node.archiveName = safeFsName(node.name);
    }
    const dir = join(this.archiveDir, ...node.archiveName.split('/'));
    mkdirSync(dir, { recursive: true });
    this.fsChanged();
    return dir;
  }

  /** Contagem de .eml na parte local de uma pasta (mesclada ou arc:). */
  private arcCountFor(node: FolderNode): number {
    return node.archiveName ? this.archiveFiles(`arc:${node.archiveName}`).length : 0;
  }

  /**
   * Mapa de consolidação de sombras: caminho antigo → caminho canônico.
   * Preenchido quando dois diretórios-sombra mapeiam para a mesma pasta do PST.
   */
  private shadowMap = new Map<string, string>();

  /** Resolve um caminho-sombra para o canônico (se o dir ou um ancestral foi consolidado). */
  private canonShadow(name: string): string {
    if (this.shadowMap.has(name)) return this.shadowMap.get(name)!;
    const parts = name.split('/');
    for (let i = parts.length - 1; i > 0; i--) {
      const anc = this.shadowMap.get(parts.slice(0, i).join('/'));
      if (anc) return `${anc}/${parts.slice(i).join('/')}`;
    }
    return name;
  }

  /**
   * Consolida o diretório-sombra `dup` dentro de `primary`: move os .eml (e
   * subpastas), remapeia o estado de leitura e remove a estrutura vazia.
   * Garante que mensagens entregues por uma regra com caminho alternativo
   * (ex.: "Início do arquivo.../Agencia Unit" vs "Agencia Unit") apareçam.
   */
  private consolidateShadow(primary: string, dup: string): void {
    if (!this.archiveDir || primary === dup) return;
    const src = join(this.archiveDir, ...dup.split('/'));
    const dst = join(this.archiveDir, ...primary.split('/'));
    if (!existsSync(src)) return;
    const st = this.loadReadState();
    const move = (s: string, d: string, rel: string) => {
      mkdirSync(d, { recursive: true });
      for (const ent of readdirSync(s, { withFileTypes: true })) {
        const relName = rel ? `${rel}/${ent.name}` : ent.name;
        if (ent.isDirectory()) {
          move(join(s, ent.name), join(d, ent.name), relName);
          continue;
        }
        if (!ent.name.endsWith('.eml')) continue;
        let target = join(d, ent.name);
        if (existsSync(target)) target = join(d, `${Date.now()}-m-${ent.name}`);
        try {
          renameSync(join(s, ent.name), target);
        } catch {
          continue; // falha ao mover — não perde o original
        }
        const oldKey = `${dup}/${relName}`;
        if (st.has(oldKey)) {
          st.set(`${primary}/${relName}`, st.get(oldKey)!);
          st.delete(oldKey);
        }
      }
    };
    try {
      move(src, dst, '');
      rmSync(src, { recursive: true, force: true });
      this.saveReadState();
    } catch {
      // consolidação é best-effort — na pior hipótese a pasta duplicada
      // continua visível como antes
    }
  }

  // ---------- Pastas de arquivamento local (.eml) ----------

  /** Nó da árvore: id "arc:" + caminho relativo (a/b) — suporta subpastas aninhadas. */
  /**
   * Listagem de diretório cacheada por mtime — criar/remover/renomear .eml
   * dentro de um diretório altera o mtime dele, então o cache invalida
   * sozinho; mutações feitas pelo app também limpam via fsChanged().
   * Sem isso, cada refresh da árvore relia milhares de nomes de arquivo.
   */
  private dirCache = new Map<string, { mtimeMs: number; dirs: string[]; emls: string[] }>();

  /** Algo mudou no disco dentro da base — descarta listagens cacheadas. */
  private fsChanged(): void {
    this.dirCache.clear();
  }

  private listDir(dir: string): { mtimeMs: number; dirs: string[]; emls: string[] } | null {
    if (!existsSync(dir)) return null;
    const mt = statSync(dir).mtimeMs;
    const c = this.dirCache.get(dir);
    if (c && c.mtimeMs === mt) return c;
    const dirs: string[] = [];
    const emls: string[] = [];
    for (const e of readdirSync(dir, { withFileTypes: true })) {
      if (e.isDirectory()) dirs.push(e.name);
      else if (e.name.endsWith('.eml')) emls.push(e.name);
    }
    // Nomes começam com o timestamp da mensagem — reversa = mais novo primeiro
    emls.sort().reverse();
    const out = { mtimeMs: mt, dirs, emls };
    this.dirCache.set(dir, out);
    return out;
  }

  private archiveFolderNodes(rel = ''): FolderNode[] {
    if (!this.archiveDir) return [];
    const dir = rel ? join(this.archiveDir, ...rel.split('/')) : this.archiveDir;
    const lst = this.listDir(dir);
    if (!lst) return [];
    return lst.dirs
      .map((d) => {
        const relPath = rel ? `${rel}/${d}` : d;
        const files = this.listDir(join(dir, d))?.emls ?? [];
        const st = this.loadReadState();
        const unread = files.filter((f) => st.get(`${relPath}/${f}`) !== true).length;
        return {
          id: `arc:${relPath}`,
          name: d,
          containerClass: containerClassForFolderName(d),
          contentCount: files.length,
          unreadCount: unread,
          hidden: this.isFolderHidden(`arc:${relPath}`) || undefined,
          children: this.archiveFolderNodes(relPath),
        };
      });
  }

  private archiveRelPath(folderId: string): string {
    return folderId.slice(4);
  }

  private archiveFiles(folderId: string): string[] {
    if (!this.archiveDir) return [];
    // Já sai ordenada do cache (mais novo primeiro)
    return this.listDir(this.archiveDirFor(folderId))?.emls ?? [];
  }

  private archiveDirFor(folderId: string): string {
    // canonShadow resolve ids de sombras já consolidadas (ex.: caminho aninhado
    // antigo) para o diretório canônico — listagem/leitura nunca perde itens.
    return join(this.archiveDir!, ...this.canonShadow(this.archiveRelPath(folderId)).split('/'));
  }

  /** Cache de resumos .eml — evita re-parsear a pasta inteira a cada clique. */
  private sumCache = new Map<string, { mtimeMs: number; s: Omit<ItemSummary, 'index' | 'isRead'> }>();

  /** Estado de leitura dos .eml — persistido em readstate.json na raiz do arquivo local. */
  private readState: Map<string, boolean> | null = null;

  /**
   * Itens do PST ocultos por "exclusão" — o arquivo é somente leitura, então
   * apagar = marcar como oculto (hidden.json, por pasta → índices brutos do
   * PST). O item some de toda a interface, mas continua dentro do .pst
   * físico (abrir o arquivo no Outlook ainda o mostra).
   */
  private hidden: Map<string, Set<number>> | null = null;

  private loadHidden(): Map<string, Set<number>> {
    if (!this.hidden) {
      this.hidden = new Map();
      try {
        const raw = readFileSync(join(this.archiveDir!, 'hidden.json'), 'utf8');
        const obj = JSON.parse(raw) as Record<string, number[]>;
        for (const k of Object.keys(obj)) this.hidden.set(k, new Set(obj[k]));
      } catch {
        // primeira vez — sem ocultos
      }
    }
    return this.hidden;
  }

  private saveHidden(): void {
    if (this.batchDepth > 0) return; // grava uma vez no fim do lote
    if (!this.hidden || !this.archiveDir) return;
    try {
      const obj: Record<string, number[]> = {};
      for (const [k, v] of this.hidden) obj[k] = [...v];
      writeFileSync(join(this.archiveDir, 'hidden.json'), JSON.stringify(obj), 'utf8');
    } catch {
      // diretório pode não existir ainda
    }
  }

  /**
   * Quantos itens ocultos têm FLAG de não-lido no PST. O badge de não-lidos
   * subtrai SÓ esses do baseline — um oculto já lido não pode descontar.
   * Sem isso, excluir/mover um item já lido (ex.: convite aceito, que vai
   * para Itens Excluídos após a resposta) derrubava o contador em 1 a mais
   * a cada operação e o número ficava errado para sempre.
   */
  private hiddenUnreadCount(folderId: string): number {
    const hidden = this.hiddenSet(folderId);
    if (hidden.size === 0) return 0;
    const folder = this.folders.get(folderId);
    if (!folder) return 0;
    let n = 0;
    for (const raw of hidden) {
      try {
        folder.moveChildCursorTo(raw);
        const msg = folder.getNextChild();
        if (msg && !msg.isRead) n++;
      } catch {
        // índice fora de faixa — ignora
      }
    }
    return n;
  }

  private hiddenSet(folderId: string): Set<number> {
    return this.loadHidden().get(folderId) ?? new Set();
  }

  // ---------- Categorias por item (itemcats.json no diretório da base) ----------

  /** Categorias aplicadas a itens — chave "a|<pasta>/<arquivo>" para .eml
   *  locais ou "p|<folderId>|<posição bruta>" para itens gravados no PST. */
  private itemCats: Map<string, string[]> | null = null;

  private loadItemCats(): Map<string, string[]> {
    if (!this.itemCats) {
      this.itemCats = new Map();
      if (this.archiveDir) {
        try {
          const raw = readFileSync(join(this.archiveDir, 'itemcats.json'), 'utf8');
          const obj = JSON.parse(raw) as Record<string, string[]>;
          for (const [k, v] of Object.entries(obj)) {
            if (Array.isArray(v)) this.itemCats.set(k, v.map(String));
          }
        } catch {
          // sem arquivo ainda
        }
      }
    }
    return this.itemCats;
  }

  private saveItemCats(): void {
    if (this.batchDepth > 0) return; // grava uma vez no fim do lote
    if (!this.itemCats || !this.archiveDir) return;
    try {
      const obj: Record<string, string[]> = {};
      for (const [k, v] of this.itemCats) if (v.length) obj[k] = v;
      writeFileSync(join(this.archiveDir, 'itemcats.json'), JSON.stringify(obj), 'utf8');
    } catch {
      // diretório pode não existir ainda
    }
  }

  /** Chave estável do item para categorias. */
  private itemCatKey(folderId: string, index: number, stableId?: string): string | null {
    index = this.idxFor(folderId, index, stableId);
    if (folderId.startsWith('arc:')) {
      const f = this.archiveFiles(folderId)[index];
      return f ? `a|${folderId.slice(4)}/${f}` : null;
    }
    const arcCount = this.mergedArcCount(folderId);
    if (index < arcCount) {
      const arcId = `arc:${this.nodeFor(folderId)!.archiveName}`;
      const f = this.archiveFiles(arcId)[index];
      return f ? `a|${arcId.slice(4)}/${f}` : null;
    }
    return `p|${folderId}|${index - arcCount}`;
  }

  private catsOfKey(key: string | null): string[] | undefined {
    if (!key) return undefined;
    const c = this.loadItemCats().get(key);
    return c && c.length ? [...c] : undefined;
  }

  /** Categorias do item (ids da lista-mestre). */
  getItemCategories(folderId: string, index: number, stableId?: string): string[] {
    return this.catsOfKey(this.itemCatKey(folderId, index, stableId)) ?? [];
  }

  /** Substitui as categorias do item — lista vazia remove a marcação. */
  setItemCategories(folderId: string, index: number, cats: string[], stableId?: string): void {
    const key = this.itemCatKey(folderId, index, stableId);
    if (!key) return;
    const m = this.loadItemCats();
    if (cats.length) m.set(key, [...new Set(cats)]);
    else m.delete(key);
    this.saveItemCats();
  }

  /** Vincula categorias a um arquivo recém-gravado (usado por moveItem). */
  setItemCategoriesByFile(folderId: string, fileName: string, cats: string[]): void {
    if (!cats.length) return;
    const rel = folderId.startsWith('arc:')
      ? folderId.slice(4)
      : this.nodeFor(folderId)?.archiveName;
    if (!rel) return;
    this.loadItemCats().set(`a|${rel}/${fileName}`, [...new Set(cats)]);
    this.saveItemCats();
  }

  /**
   * Pastas ocultadas pelo usuário ("não uso esta pasta") — persiste em
   * hiddenfolders.json. Ocultar não apaga nada: a pasta só some da árvore.
   * Funciona para pastas do PST (id "0/2/…") e locais ("arc:…").
   */
  private hiddenFolders: Set<string> | null = null;

  private loadHiddenFolders(): Set<string> {
    if (this.hiddenFolders) return this.hiddenFolders;
    this.hiddenFolders = new Set();
    try {
      if (this.archiveDir) {
        const raw = readFileSync(join(this.archiveDir, 'hiddenfolders.json'), 'utf8');
        const arr = JSON.parse(raw);
        if (Array.isArray(arr)) this.hiddenFolders = new Set(arr.map(String));
      }
    } catch {
      // sem arquivo ainda
    }
    return this.hiddenFolders;
  }

  private saveHiddenFolders(): void {
    if (!this.hiddenFolders || !this.archiveDir) return;
    try {
      writeFileSync(
        join(this.archiveDir, 'hiddenfolders.json'),
        JSON.stringify([...this.hiddenFolders]),
        'utf8',
      );
    } catch {
      // ignora
    }
  }

  isFolderHidden(folderId: string): boolean {
    return this.loadHiddenFolders().has(folderId);
  }

  /** Oculta/reexibe uma pasta e retorna a árvore atualizada. */
  setFolderHidden(folderId: string, hidden: boolean): FolderNode[] {
    const set = this.loadHiddenFolders();
    if (hidden) set.add(folderId);
    else set.delete(folderId);
    this.saveHiddenFolders();
    this.rebuildTree();
    return this.rootNodes;
  }

  /** Oculta um item gravado no PST (exclusão lógica — PST é somente leitura). */
  hideItem(folderId: string, index: number, stableId?: string): void {
    if (folderId.startsWith('arc:')) return; // itens locais são apagados de verdade
    index = this.idxFor(folderId, index, stableId);
    const raw = index - this.mergedArcCount(folderId);
    if (raw < 0) return;
    const m = this.loadHidden();
    if (!m.has(folderId)) m.set(folderId, new Set());
    m.get(folderId)!.add(raw);
    this.saveHidden();
    this.sumCache.clear();
    const node = this.nodeFor(folderId);
    if (node) {
      node.contentCount = Math.max(0, (node.contentCount || 0) - 1);
      // Ocultou um item efetivamente NÃO-LIDO → o badge desce já aqui,
      // sem esperar o próximo rebuild da árvore.
      try {
        const folder = this.folders.get(folderId);
        if (folder) {
          folder.moveChildCursorTo(raw);
          const msg = folder.getNextChild();
          const eff = msg ? (this.pstReadOverride(folderId, raw) ?? !!msg.isRead) : true;
          if (!eff) node.unreadCount = Math.max(0, (node.unreadCount || 0) - 1);
        }
      } catch {
        // falha na leitura do flag — o recount da árvore corrige
      }
    }
  }

  private loadReadState(): Map<string, boolean> {
    if (!this.readState) {
      this.readState = new Map();
      const dir = this.archiveDir;
      if (!dir) return this.readState;
      try {
        const raw = readFileSync(join(dir, 'readstate.json'), 'utf8');
        const obj = JSON.parse(raw) as Record<string, boolean>;
        for (const k of Object.keys(obj)) this.readState.set(k, obj[k]);
      } catch {
        // Primeira vez sem readstate.json: os .eml que já existiam (entregues
        // antes deste recurso) entram como LIDOS — só novas chegadas marcam
        // como não-lidas.
        const seed = (rel: string) => {
          const d = rel ? join(dir, ...rel.split('/')) : dir;
          const lst = this.listDir(d);
          if (!lst) return;
          for (const sub of lst.dirs) seed(rel ? `${rel}/${sub}` : sub);
          for (const name of lst.emls) {
            this.readState!.set(rel ? `${rel}/${name}` : name, true);
          }
        };
        seed('');
      }
      // Mensagens que o próprio usuário enviou nunca são "não lidas" — saneia
      // cópias gravadas antes desta regra (ficavam sem entrada = não lidas).
      // Reconhece a pasta Enviados/Saída em qualquer idioma suportado.
      for (const d of readdirSync(dir, { withFileTypes: true })) {
        if (!d.isDirectory()) continue;
        const kind = folderKindOf(d.name);
        if (kind !== 'sent' && kind !== 'outbox') continue;
        for (const f of readdirSync(join(dir, d.name))) {
          if (f.endsWith('.eml') && !this.readState.has(`${d.name}/${f}`)) {
            this.readState.set(`${d.name}/${f}`, true);
          }
        }
      }
      this.readStateVer++;
    }
    return this.readState;
  }

  /** Versão do readState — invalida o cache de contagens por pasta. */
  private readStateVer = 0;

  private saveReadState(): void {
    if (this.batchDepth > 0) return; // grava uma vez no fim do lote
    if (!this.readState) return;
    this.readStateVer++;
    try {
      const obj: Record<string, boolean> = {};
      for (const [k, v] of this.readState) obj[k] = v;
      writeFileSync(join(this.archiveDir!, 'readstate.json'), JSON.stringify(obj), 'utf8');
    } catch {
      // diretório pode não existir ainda
    }
  }

  /** Chave do estado de leitura: "<arcId>/<arquivo>" — sempre com o caminho
   *  de sombra CANÔNICO (canonShadow), senão gravar por um id não-canônico
   *  (ex.: pasta renomeada/consolidada) criaria estado órfão e a árvore —
   *  que conta pelo caminho real do diretório — mostraria não-lidos
   *  fantasma que nenhum "marcar tudo como lido" apaga. */
  private readKey(arcId: string, file: string): string {
    return `${this.canonShadow(arcId.slice(4))}/${file}`;
  }

  private isReadFile(arcId: string, file: string): boolean {
    return this.loadReadState().get(this.readKey(arcId, file)) === true;
  }

  /**
   * Override de leitura para itens gravados no PST — o arquivo é somente
   * leitura, então "marcar como lida" é persistido no readstate.json com a
   * chave "pst|<folderId>|<índice bruto>". O valor guarda o isRead desejado;
   * itens sem override usam o flag gravado no PST.
   */
  private pstReadKey(folderId: string, rawIndex: number): string {
    return `pst|${folderId}|${rawIndex}`;
  }

  private pstReadOverride(folderId: string, rawIndex: number): boolean | undefined {
    return this.loadReadState().get(this.pstReadKey(folderId, rawIndex));
  }

  private setPstRead(folderId: string, rawIndex: number, read: boolean): void {
    const st = this.loadReadState();
    st.set(this.pstReadKey(folderId, rawIndex), read);
    this.saveReadState();
  }

  /**
   * Overrides de leitura do PST desta pasta: `read` = marcados lidos sobre
   * flag não-lido do PST (subtrai do badge), `unread` = marcados não-lidos
   * sobre flag lido (soma no badge). Itens ocultos ("excluídos") não contam.
   * O mapa por pasta é pré-computado por versão do readState — antes cada
   * chamada varria o readstate inteiro (O(pastas × overrides) por refresh).
   */
  private roCounts: {
    ver: number;
    read: Map<string, number>;
    unread: Map<string, number>;
  } | null = null;
  private pstReadOverrideCounts(folderId: string): { read: number; unread: number } {
    if (!this.roCounts || this.roCounts.ver !== this.readStateVer) {
      const read = new Map<string, number>();
      const unread = new Map<string, number>();
      for (const [k, v] of this.loadReadState()) {
        if (!k.startsWith('pst|')) continue;
        const lastBar = k.lastIndexOf('|');
        const fid = k.slice(4, lastBar);
        const raw = parseInt(k.slice(lastBar + 1), 10);
        if (this.hiddenSet(fid).has(raw)) continue;
        const m = v === true ? read : unread;
        m.set(fid, (m.get(fid) ?? 0) + 1);
      }
      this.roCounts = { ver: this.readStateVer, read, unread };
    }
    return {
      read: this.roCounts.read.get(folderId) ?? 0,
      unread: this.roCounts.unread.get(folderId) ?? 0,
    };
  }

  /** Marca um item como lido (usado ao abrir o detalhe). */
  markItemRead(folderId: string, index: number, stableId?: string): void {
    index = this.idxFor(folderId, index, stableId);
    if (!folderId.startsWith('arc:')) {
      const arcCount = this.mergedArcCount(folderId);
      if (index >= arcCount) {
        // Item do PST — override lógico
        const raw = index - arcCount;
        const folder = this.folders.get(folderId);
        if (!folder) return;
        folder.moveChildCursorTo(raw);
        const msg = folder.getNextChild();
        // Considera o estado EFETIVO (flag do PST + override) — se o usuário
        // marcou "não lido" num item que o PST grava como lido, abrir a
        // mensagem deve gravar o override "lido" mesmo assim.
        if (msg) {
          const eff = this.pstReadOverride(folderId, raw) ?? !!msg.isRead;
          if (!eff) this.setPstRead(folderId, raw, true);
        }
        return;
      }
    }
    const arcId = folderId.startsWith('arc:')
      ? folderId
      : `arc:${this.nodeFor(folderId)?.archiveName ?? ''}`;
    if (arcId === 'arc:') return;
    const file = this.archiveFiles(arcId)[index];
    if (!file) return;
    this.loadReadState().set(this.readKey(arcId, file), true);
    this.saveReadState();
  }

  /** Define lido/não-lido — local altera readstate do .eml, item do PST grava override. */
  setItemRead(folderId: string, index: number, read: boolean, stableId?: string): void {
    index = this.idxFor(folderId, index, stableId);
    if (!folderId.startsWith('arc:')) {
      const arcCount = this.mergedArcCount(folderId);
      if (index >= arcCount) {
        const raw = index - arcCount;
        const folder = this.folders.get(folderId);
        if (!folder) return;
        folder.moveChildCursorTo(raw);
        const msg = folder.getNextChild();
        if (!msg) return;
        const orig = !!msg.isRead;
        const key = this.pstReadKey(folderId, raw);
        const st = this.loadReadState();
        // Se o override coincide com o flag do PST, não precisa persistir
        if (orig === read) st.delete(key);
        else st.set(key, read);
        this.saveReadState();
        return;
      }
    }
    const arcId = folderId.startsWith('arc:')
      ? folderId
      : `arc:${this.nodeFor(folderId)?.archiveName ?? ''}`;
    if (arcId === 'arc:') return;
    const file = this.archiveFiles(arcId)[index];
    if (!file) return;
    this.loadReadState().set(this.readKey(arcId, file), read);
    this.saveReadState();
  }

  /** Marca todos os itens de uma pasta como lidos (locais e do PST). */
  markAllRead(folderId: string): void {
    const st = this.loadReadState();
    const mark = (arcId: string) => {
      for (const f of this.archiveFiles(arcId)) st.set(this.readKey(arcId, f), true);
    };
    if (folderId.startsWith('arc:')) {
      mark(folderId);
    } else {
      const node = this.nodeFor(folderId);
      if (node?.archiveName) mark(`arc:${node.archiveName}`);
      // Itens não-lidos do PST ganham override "lido"
      const folder = this.folders.get(folderId);
      if (folder) {
        const hidden = this.hiddenSet(folderId);
        const total = folder.contentCount || 0;
        folder.moveChildCursorTo(0);
        for (let i = 0; i < total; i++) {
          const msg = folder.getNextChild();
          if (!msg) break;
          // Estado EFETIVO — um override "não lido" anterior também precisa
          // ser virado para lido (msg.isRead do PST seria true e pulava).
          const eff = this.pstReadOverride(folderId, i) ?? !!msg.isRead;
          if (!hidden.has(i) && !eff) st.set(this.pstReadKey(folderId, i), true);
        }
      }
    }
    this.saveReadState();
  }

  private async archiveItems(folderId: string, offset: number, limit: number): Promise<ItemSummary[]> {
    const files = this.archiveFiles(folderId).slice(offset, offset + limit);
    const dir = this.archiveDirFor(folderId);
    const idx = this.loadSumIdx();
    const rel = folderId.slice(4);
    const out: ItemSummary[] = [];
    for (let i = 0; i < files.length; i++) {
      const fp = join(dir, files[i]);
      try {
        const mtimeMs = statSync(fp).mtimeMs;
        let c = this.sumCache.get(fp);
        if (c && c.mtimeMs === mtimeMs) {
          // cache quente em memória
        } else {
          // Índice persistido em disco — evita re-parsear o .eml a cada
          // sessão. Só o arquivo novo/alterado é parseado de fato.
          const hit = idx.items[`a|${rel}|${files[i]}`];
          if (hit && hit.m === mtimeMs) {
            c = { mtimeMs, s: hit.s };
            this.sumCache.set(fp, c);
          } else {
            const parsed = await simpleParser(readFileSync(fp));
            const { index: _i, isRead: _r, ...rest } = this.summarizeParsed(
              parsed,
              0,
              statSync(fp).size,
            );
            c = { mtimeMs, s: rest };
            this.sumCache.set(fp, c);
            idx.items[`a|${rel}|${files[i]}`] = { m: mtimeMs, s: rest };
            this.markSumIdxDirty();
          }
        }
        out.push({
          ...c.s,
          index: offset + i,
          isRead: this.isReadFile(folderId, files[i]),
          categories: this.catsOfKey(`a|${rel}/${files[i]}`),
          local: true,
          stableId: `a|${this.canonShadow(rel)}/${files[i]}`,
        });
      } catch {
        // arquivo corrompido — pula
      }
    }
    return out;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private summarizeParsed(parsed: any, index: number, size = 0): ItemSummary {
    const from = parsed.from?.value?.[0] || {};
    const to = (parsed.to as { value?: Array<{ address?: string; name?: string }> } | undefined)?.value || [];
    // Cabeçalhos X-UM-* — itens criados pelo Unit Mail (compromisso/contato/tarefa)
    // guardam os campos estruturados como headers do .eml.
    const hdr = (name: string): string => {
      const v = parsed.headers?.get?.(name);
      return Array.isArray(v) ? String(v[0] ?? '') : v != null ? String(v) : '';
    };
    const umClass = hdr('x-um-class');
    const s: ItemSummary = {
      index,
      messageClass: umClass || 'IPM.Note',
      subject: fixDecodedHeader(
        parsed.subject || '(sem assunto)',
        (parsed.headerLines || []).find((h: { key: string; line: string }) => h.key === 'subject')?.line,
      ),
      senderName: from.name || from.address || '',
      senderEmailAddress: from.address || '',
      displayTo: to.map((a) => (a.name ? `${a.name} <${a.address}>` : a.address || '')).join('; '),
      displayCC: '',
      sentTime: parsed.date ? parsed.date.toISOString() : null,
      deliveryTime: parsed.date ? parsed.date.toISOString() : null,
      // Mesmo filtro do painel de leitura: inline-com-cid (logo de
      // assinatura) não é anexo — sem isso o clipe aparecia à toa.
      hasAttachments: (parsed.attachments || []).some(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (a: any) => !(a.contentDisposition === 'inline' && a.contentId),
      ),
      isRead: true,
      importance: 0,
      size,
      preview: (parsed.text || '').slice(0, 200),
      accountEmail: hdr('x-um-account') || undefined,
    };
    // Mensagem com parte iCalendar (convite salvo/compromisso aceito) —
    // extrai data/local para a pasta de Calendário e para lembretes.
    const icsAtt = (parsed.attachments || []).find(
      (a: { contentType?: string; filename?: string }) =>
        (a.contentType || '').includes('text/calendar') ||
        (a.filename || '').toLowerCase().endsWith('.ics'),
    );
    if (icsAtt?.content) {
      const inv = parseIcsInvite(icsAtt.content.toString('utf8'));
      if (inv) {
        s.startTime = inv.start;
        s.endTime = inv.end;
        s.location = inv.location;
        s.apptUid = inv.uid || s.apptUid;
        s.reminderSet = true;
        s.reminderMinutes = inv.reminderMinutes ?? 15;
        // Convite/resposta recebido (REQUEST/CANCEL/REPLY) → ícone de reunião.
        // PUBLISH = compromisso criado pelo próprio usuário — não é convite.
        if (inv.method !== 'PUBLISH') s.inviteMethod = inv.method;
      }
    }

    // Compromisso criado pelo Unit Mail — headers X-UM-Appt-* (o .ics anexo
    // já preencheu start/end/location; os headers servem de fallback).
    if (umClass === 'IPM.Appointment') {
      s.apptUid = s.apptUid || hdr('x-um-appt-uid') || undefined;
      s.startTime = s.startTime || hdr('x-um-appt-start') || null;
      s.endTime = s.endTime || hdr('x-um-appt-end') || null;
      s.location = s.location || hdr('x-um-appt-location');
      s.allDay = hdr('x-um-appt-allday') === '1';
      if (!s.reminderSet && hdr('x-um-appt-reminder')) {
        s.reminderSet = true;
        s.reminderMinutes = Number(hdr('x-um-appt-reminder')) || 15;
      }
    }
    if (umClass === 'IPM.Contact') {
      s.contactDisplayName = hdr('x-um-contact-name') || s.subject;
      s.contactEmail = hdr('x-um-contact-email');
      s.contactPhone = hdr('x-um-contact-phone') || hdr('x-um-contact-mobile');
      s.contactCompany = hdr('x-um-contact-company');
    }
    if (umClass === 'IPM.Task') {
      s.dueDate = hdr('x-um-task-due') || null;
      s.percentComplete = Number(hdr('x-um-task-pct')) || 0;
      s.isTaskComplete = hdr('x-um-task-complete') === '1' || s.percentComplete >= 100;
    }
    return s;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async detailFromParsed(parsed: any, index: number): Promise<ItemDetail> {
    const attachments: AttachmentMeta[] = (parsed.attachments || []).map(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (a: any, i: number) => ({
        index: i,
        filename: a.filename || `anexo-${i}`,
        size: a.size || a.content?.length || 0,
        mimeTag: a.contentType || '',
        contentId: a.contentId || '',
        isEmbeddedMessage: a.contentType === 'message/rfc822',
        isInline: a.contentDisposition === 'inline' && !!a.contentId,
      }),
    );
    const detail: ItemDetail = {
      ...this.summarizeParsed(parsed, index),
      body: parsed.text || '',
      bodyHTML: typeof parsed.html === 'string' ? parsed.html : '',
      transportMessageHeaders: parsed.headerLines?.map((h: { key: string; line: string }) => `${h.key}: ${h.line}`).join('\n') || '',
      attachments: attachments.filter((a) => !a.isInline),
    };
    // Convite iCalendar anexado (.ics / text/calendar) → Aceitar/Recusar.
    // METHOD:PUBLISH = compromisso criado pelo próprio usuário — não mostra
    // os botões de resposta (não é um convite recebido).
    const icsAtt = (parsed.attachments || []).find(
      (a: { contentType?: string; filename?: string; content?: Buffer }) =>
        ((a.contentType || '').includes('text/calendar') ||
          (a.filename || '').toLowerCase().endsWith('.ics')) &&
        a.content,
    );
    if (icsAtt) {
      const icsText = icsAtt.content.toString('utf8');
      const inv = parseIcsInvite(icsText);
      if (inv && inv.method !== 'PUBLISH') {
        detail.invite = {
          ...inv,
          ics: icsText,
          organizer: inv.organizer || detail.senderName,
          organizerEmail: inv.organizerEmail || detail.senderEmailAddress,
          summary: inv.summary || detail.subject,
        };
      }
    }

    // Itens criados pelo Unit Mail — campos estruturados nos headers X-UM-*.
    const hdr = (name: string): string => {
      const v = parsed.headers?.get?.(name);
      return Array.isArray(v) ? String(v[0] ?? '') : v != null ? String(v) : '';
    };
    if (detail.messageClass === 'IPM.Appointment') {
      detail.organizer = hdr('x-um-appt-organizer') || detail.organizer;
      detail.requiredAttendees = hdr('x-um-appt-attendees') || detail.requiredAttendees;
      detail.isRecurring = false;
    }
    if (detail.messageClass === 'IPM.Contact') {
      const contact: ContactDetail = {
        displayName: hdr('x-um-contact-name') || detail.subject,
        givenName: hdr('x-um-contact-given'),
        middleName: '',
        surname: hdr('x-um-contact-surname'),
        company: hdr('x-um-contact-company'),
        jobTitle: hdr('x-um-contact-jobtitle'),
        department: hdr('x-um-contact-dept'),
        email1: hdr('x-um-contact-email'),
        email1Address: hdr('x-um-contact-email'),
        email2: hdr('x-um-contact-email2'),
        email2Address: hdr('x-um-contact-email2'),
        email3: '',
        email3Address: '',
        businessPhone: hdr('x-um-contact-phone'),
        homePhone: '',
        mobilePhone: hdr('x-um-contact-mobile'),
        businessAddress: '',
        homeAddress: '',
        birthday: hdr('x-um-contact-birthday') || null,
        notes: detail.body,
      };
      detail.contact = contact;
    }
    if (detail.messageClass === 'IPM.Task') {
      const task: TaskDetail = {
        status: Number(hdr('x-um-task-status')) || 0,
        percentComplete: detail.percentComplete ?? (Number(hdr('x-um-task-pct')) || 0),
        isComplete: detail.isTaskComplete ?? (hdr('x-um-task-complete') === '1'),
        startDate: hdr('x-um-task-start') || null,
        dueDate: detail.dueDate || hdr('x-um-task-due') || null,
        dateCompleted: hdr('x-um-task-completed') || null,
        owner: hdr('x-um-task-owner'),
        body: detail.body,
      };
      detail.task = task;
    }
    return detail;
  }

  private getFolder(folderId: string): PSTFolder {
    const folder = this.folders.get(folderId);
    if (!folder) throw new Error(`Pasta não encontrada: ${folderId}`);
    return folder;
  }

  /** Conta .eml da parte local da pasta (se mesclada a uma pasta do PST). */
  private mergedArcCount(folderId: string): number {
    if (folderId.startsWith('arc:')) return 0;
    const node = this.nodeFor(folderId);
    return node?.archiveName ? this.archiveFiles(`arc:${node.archiveName}`).length : 0;
  }

  /**
   * Resolve um `stableId` para o índice posicional ATUAL na pasta.
   *
   * O `index` dos itens é posicional e instável: um .eml entregue pelo
   * Enviar/Receber entra no topo e desloca todos os índices da pasta (nas
   * mescladas, arcCount cresce e todos os itens do PST deslocam junto).
   * Qualquer chamada com índice velho abriria/apagaria a mensagem errada.
   * O stableId identifica o item de verdade:
   *   a|<rel-canon>/<arquivo> → posição do arquivo na ordenação atual
   *   p|<folderId>|<pos-bruta> → posição imutável no PST + arcCount atual
   * Retorna null quando o item não está mais nesta pasta — falhar é melhor
   * do que agir sobre a mensagem errada.
   */
  private indexByStableId(folderId: string, stableId: string): number | null {
    if (stableId.startsWith('a|')) {
      const rest = stableId.slice(2);
      const slash = rest.lastIndexOf('/');
      if (slash < 0) return null;
      const rel = rest.slice(0, slash);
      const file = rest.slice(slash + 1);
      // Garante que o .eml pertence a ESTA pasta — um item movido tem o rel
      // da pasta antiga no stableId e não pode resolver aqui.
      const expected = folderId.startsWith('arc:')
        ? this.canonShadow(folderId.slice(4))
        : (this.nodeFor(folderId)?.archiveName
            ? this.canonShadow(this.nodeFor(folderId)!.archiveName!)
            : null);
      if (!expected || rel !== expected) return null;
      const i = this.archiveFiles(`arc:${rel}`).indexOf(file);
      return i < 0 ? null : i;
    }
    if (stableId.startsWith('p|')) {
      const cut = stableId.lastIndexOf('|');
      const raw = Number(stableId.slice(cut + 1));
      // A pasta gravada no stableId precisa ser a pasta consultada — itens
      // do PST só existem em pastas mescladas (nunca em arc:).
      if (folderId.startsWith('arc:') || stableId.slice(2, cut) !== folderId) return null;
      if (!Number.isFinite(raw) || raw < 0) return null;
      return this.mergedArcCount(folderId) + raw;
    }
    return null;
  }

  /**
   * Índice efetivo de uma chamada: quando o emissor informou o stableId do
   * item (sempre que veio de um ItemSummary), resolve para a posição atual
   * — imune a entregas/exclusões que deslocaram índices depois da listagem.
   */
  private idxFor(folderId: string, index: number, stableId?: string): number {
    if (!stableId) return index;
    const resolved = this.indexByStableId(folderId, stableId);
    if (resolved === null) {
      throw new Error('A mensagem não está mais nesta pasta — a lista foi atualizada, selecione-a novamente.');
    }
    return resolved;
  }

  /** Lista itens de uma pasta. Retorna em ordem reversa (mais recentes primeiro). */
  async getItems(folderId: string, offset = 0, limit = 500): Promise<ItemSummary[]> {
    this.requireOpen();
    if (folderId.startsWith('arc:')) return this.archiveItems(folderId, offset, limit);
    const folder = this.getFolder(folderId);
    const pstTotal = folder.contentCount || 0;
    const arcCount = this.mergedArcCount(folderId);
    const total = pstTotal + arcCount;
    if (total <= 0 || offset >= total) return [];

    // Espaço de índices mesclado: [0..arcCount) = .eml locais (mais novos
    // primeiro), [arcCount..) = itens do PST (mais novos primeiro).
    const items: ItemSummary[] = [];
    if (offset < arcCount) {
      const arcItems = await this.archiveItems(
        `arc:${this.nodeFor(folderId)!.archiveName}`,
        offset,
        Math.min(limit, arcCount - offset),
      );
      items.push(...arcItems); // index já sai como offset+i — casa com o espaço mesclado
    }
    const hidden = this.hiddenSet(folderId);
    const remaining = limit - items.length;
    const pOffset = Math.max(0, offset - arcCount);
    if (remaining > 0 && pOffset < pstTotal) {
      // O cursor do PST itera do mais antigo ao mais novo — pegamos do fim.
      const end = pstTotal - pOffset;
      const start = Math.max(0, end - remaining - hidden.size);
      folder.moveChildCursorTo(start);
      const idx = this.loadSumIdx();
      const pstItems: ItemSummary[] = [];
      for (let i = start; i < end; i++) {
        const child = folder.getNextChild();
        if (!child) break;
        if (hidden.has(i)) continue; // excluído logicamente — não exibir
        // Resumo persistido: o PST é imutável, então 'p|folder|i' é estável
        // entre sessões — pula o summarize (o cursor avança igual).
        const hit = idx.items[`p|${folderId}|${i}`];
        let s: ItemSummary;
        if (hit) {
          s = { ...hit.s, index: i, isRead: !!child.isRead };
        } else {
          s = this.summarize(child, i);
          const { index: _i, isRead: _r, ...rest } = s;
          idx.items[`p|${folderId}|${i}`] = { s: rest };
          this.markSumIdxDirty();
        }
        const ov = this.pstReadOverride(folderId, i);
        if (ov !== undefined) s.isRead = ov;
        s.categories = this.catsOfKey(`p|${folderId}|${i}`);
        s.stableId = `p|${folderId}|${i}`; // posição bruta — estável entre reindex
        s.index = arcCount + i; // desloca para o espaço mesclado
        pstItems.push(s);
      }
      items.push(...pstItems.reverse());
    }
    return items;
  }

  getItemCount(folderId: string): number {
    if (folderId.startsWith('arc:')) return this.archiveFiles(folderId).length;
    const folder = this.getFolder(folderId);
    return (
      Math.max(0, (folder.contentCount || 0) - this.hiddenSet(folderId).size) +
      this.mergedArcCount(folderId)
    );
  }

  async getItem(folderId: string, index: number, stableId?: string): Promise<ItemDetail> {
    this.requireOpen();
    index = this.idxFor(folderId, index, stableId);
    if (folderId.startsWith('arc:')) {
      const file = this.archiveFiles(folderId)[index];
      if (!file) throw new Error(`Item ${index} não encontrado na pasta ${folderId}`);
      this.markItemRead(folderId, index);
      const parsed = await simpleParser(readFileSync(join(this.archiveDirFor(folderId), file)));
      const d = await this.detailFromParsed(parsed, index);
      d.local = true;
      d.stableId = `a|${this.canonShadow(folderId.slice(4))}/${file}`;
      d.categories = this.catsOfKey(`a|${folderId.slice(4)}/${file}`);
      return d;
    }
    const arcCount = this.mergedArcCount(folderId);
    if (index < arcCount) {
      const arcId = `arc:${this.nodeFor(folderId)!.archiveName}`;
      const file = this.archiveFiles(arcId)[index];
      if (!file) throw new Error(`Item ${index} não encontrado na pasta ${folderId}`);
      this.markItemRead(folderId, index);
      const parsed = await simpleParser(readFileSync(join(this.archiveDirFor(arcId), file)));
      const d = await this.detailFromParsed(parsed, index);
      d.local = true;
      d.stableId = `a|${this.canonShadow(arcId.slice(4))}/${file}`;
      d.categories = this.catsOfKey(`a|${arcId.slice(4)}/${file}`);
      return d;
    }
    const folder = this.getFolder(folderId);
    folder.moveChildCursorTo(index - arcCount);
    const child = folder.getNextChild();
    if (!child) throw new Error(`Item ${index} não encontrado na pasta ${folderId}`);
    // Marcar como lido persiste via override (PST é somente leitura)
    this.markItemRead(folderId, index);
    const d = this.detail(child, index);
    d.isRead = this.pstReadOverride(folderId, index - arcCount) ?? d.isRead;
    d.stableId = `p|${folderId}|${index - arcCount}`;
    d.categories = this.catsOfKey(`p|${folderId}|${index - arcCount}`);
    return d;
  }

  private async arcAttachment(folderId: string, index: number, attachIndex: number): Promise<AttachmentData> {
    const file = this.archiveFiles(folderId)[index];
    if (!file) throw new Error('Item não encontrado');
    const parsed = await simpleParser(readFileSync(join(this.archiveDirFor(folderId), file)));
    // attachIndex é a posição no array COMPLETO de anexos do parse — o
    // detalhe filtra os inline para exibir mas preserva o `index` original.
    // Filtrar aqui deslocava as posições: mensagem com imagem inline antes
    // do anexo real dava "Anexo não encontrado" (ou baixava o errado).
    const att = (parsed.attachments || [])[attachIndex];
    if (!att) throw new Error('Anexo não encontrado');
    return {
      filename: att.filename || `anexo-${attachIndex}`,
      mimeTag: att.contentType || '',
      dataBase64: att.content ? att.content.toString('base64') : '',
    };
  }

  async getAttachmentData(
    folderId: string,
    index: number,
    attachIndex: number,
    stableId?: string,
  ): Promise<AttachmentData> {
    index = this.idxFor(folderId, index, stableId);
    if (folderId.startsWith('arc:')) {
      return this.arcAttachment(folderId, index, attachIndex);
    }
    const arcCount = this.mergedArcCount(folderId);
    if (index < arcCount) {
      return this.arcAttachment(`arc:${this.nodeFor(folderId)!.archiveName}`, index, attachIndex);
    }
    const folder = this.getFolder(folderId);
    folder.moveChildCursorTo(index - arcCount);
    const child = folder.getNextChild() as PSTMessage;
    if (!child) throw new Error('Item não encontrado');
    const att = child.getAttachment(attachIndex) as PSTAttachment | null;
    if (!att) throw new Error('Anexo não encontrado');
    const filename = att.longFilename || att.filename || `anexo-${attachIndex}`;

    if (att.attachMethod === 5 || att.embeddedPSTMessage) {
      const embedded = att.embeddedPSTMessage;
      return {
        filename: filename.endsWith('.msg') ? filename : `${filename}.msg`,
        mimeTag: 'message/rfc822',
        dataBase64: '',
        embeddedMessage: embedded ? this.detail(embedded, -1) : undefined,
      };
    }

    const stream = att.fileInputStream;
    if (!stream) {
      return { filename, mimeTag: att.mimeTag || '', dataBase64: '' };
    }
    const size = att.filesize || stream.length.toNumber();
    const buf = Buffer.alloc(size);
    stream.readCompletely(buf);
    return { filename, mimeTag: att.mimeTag || '', dataBase64: buf.toString('base64') };
  }

  // ---------- Mover itens (drag & drop entre pastas) ----------

  /** Exporta um item como .eml — para pasta local lê o arquivo; para PST recompõe o MIME. */
  async exportItemEml(folderId: string, index: number, stableId?: string): Promise<Buffer> {
    this.requireOpen();
    index = this.idxFor(folderId, index, stableId);
    if (folderId.startsWith('arc:')) {
      const file = this.archiveFiles(folderId)[index];
      if (!file) throw new Error('Item não encontrado');
      return readFileSync(join(this.archiveDirFor(folderId), file));
    }
    // Item da parte local de uma pasta mesclada → lê o .eml direto
    const arcCount = this.mergedArcCount(folderId);
    if (index < arcCount) {
      const arcId = `arc:${this.nodeFor(folderId)!.archiveName}`;
      const file = this.archiveFiles(arcId)[index];
      if (!file) throw new Error('Item não encontrado');
      return readFileSync(join(this.archiveDirFor(arcId), file));
    }
    // Item do PST: recomponhe o MIME a partir do detalhe + anexos
    const d = await this.getItem(folderId, index, stableId);
    const atts: Array<{ filename: string; content: Buffer; contentType?: string }> = [];
    for (const a of d.attachments) {
      try {
        const ad = await this.getAttachmentData(folderId, index, a.index, stableId);
        if (ad.dataBase64) {
          atts.push({
            filename: ad.filename,
            content: Buffer.from(ad.dataBase64, 'base64'),
            contentType: ad.mimeTag || undefined,
          });
        }
      } catch {
        // anexo ilegível — pula
      }
    }
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const MailComposer = require('nodemailer/lib/mail-composer');
    // Message-ID estável: o real (dos cabeçalhos de transporte do PST) quando
    // existe; senão um id determinístico dos campos do item. Sem isso o
    // MailComposer gerava um id aleatório a cada export — o mesmo item do PST
    // virava um .eml "diferente" e a deduplicação nunca casava.
    const midReal = /^message-id:[ \t]*(<?[^>\r\n]+>?)/im.exec(d.transportMessageHeaders || '');
    const messageId = midReal
      ? midReal[1].startsWith('<')
        ? midReal[1]
        : `<${midReal[1].trim()}>`
      : `<pst-${createHash('sha1')
          .update(
            `${d.senderEmailAddress}|${d.displayTo}|${d.displayCC}|${d.subject}|${d.deliveryTime || d.sentTime}`,
          )
          .digest('hex')}@unit-mail.pst>`;
    return new Promise<Buffer>((resolve, reject) => {
      new MailComposer({
        messageId,
        from: d.senderEmailAddress
          ? `"${(d.senderName || '').replace(/"/g, '')}" <${d.senderEmailAddress}>`
          : d.senderName,
        to: d.displayTo || undefined,
        cc: d.displayCC || undefined,
        subject: d.subject,
        date: d.deliveryTime || d.sentTime ? new Date((d.deliveryTime || d.sentTime) as string) : undefined,
        text: d.body || undefined,
        html: d.bodyHTML || undefined,
        attachments: atts,
      })
        .compile()
        .build((e: Error | null, msg: Buffer) => (e ? reject(e) : resolve(msg)));
    });
  }

  /**
   * Grava um .eml numa pasta gravável — local (arc:) ou pasta do PST com
   * sombra. `read` preserva o estado de leitura do original (ex.: deletar um
   * item lido não deve criar um "não lido" em Itens Excluídos).
   */
  async importItemEml(folderId: string, raw: Buffer, read?: boolean): Promise<string> {
    this.requireOpen();
    const dir = this.writableDirFor(folderId);
    if (!dir) {
      throw new Error('A pasta de destino não aceita itens');
    }
    const name = `${this.emlDateMs(raw) ?? Date.now()}-${randomUUID().slice(0, 8)}.eml`;
    writeFileSync(join(dir, name), raw);
    this.fsChanged();
    if (read !== undefined) {
      const arcId = folderId.startsWith('arc:')
        ? folderId
        : `arc:${this.nodeFor(folderId)?.archiveName ?? ''}`;
      if (arcId !== 'arc:') {
        this.loadReadState().set(this.readKey(arcId, name), read);
        this.saveReadState();
      }
    }
    if (this.batchDepth > 0) return name; // a árvore é reconstruída no fim do lote
    // Atualiza o contador exibido na árvore
    const bump = (nodes: FolderNode[]) => {
      for (const n of nodes) {
        if (n.id === folderId) {
          const arcPart = folderId.startsWith('arc:')
            ? this.archiveFiles(folderId).length
            : this.arcCountFor(n);
          n.contentCount = folderId.startsWith('arc:')
            ? arcPart
            : (this.folders.get(folderId)?.contentCount || 0) + arcPart;
        }
        bump(n.children);
      }
    };
    bump(this.rootNodes);
    return name;
  }

  /**
   * Cria um compromisso/contato/tarefa na base — grava um .eml com headers
   * X-UM-* (campos estruturados) e, para compromissos, um anexo .ics real
   * (METHOD:PUBLISH) — o mesmo formato que os lembretes e a listagem já leem.
   */
  async createSpecialItem(item: NewSpecialItem): Promise<string> {
    this.requireOpen();
    const kind: FolderKind =
      item.kind === 'appt' ? 'calendar' : item.kind === 'contact' ? 'contacts' : 'tasks';
    // Pasta já existente de qualquer idioma, senão o nome do idioma atual —
    // nunca cria "Calendar" ao lado de "Calendário" nem renomeia o que existe.
    let existing: string[] = [];
    try {
      if (this.archiveDir && existsSync(this.archiveDir)) {
        existing = readdirSync(this.archiveDir, { withFileTypes: true })
          .filter((d) => d.isDirectory())
          .map((d) => d.name);
      }
    } catch {
      // usa o nome localizado
    }
    const folderName = resolveFolderName(existing, kind, appLanguage());
    const raw = await this.buildSpecialEml(item);
    this.importNamedEml(folderName, raw, true);
    return folderName;
  }

  /** Monta o .eml de um item especial (compromisso/contato/tarefa). */
  private async buildSpecialEml(item: NewSpecialItem): Promise<Buffer> {
    const headers: Record<string, string> = {};
    const set = (k: string, v: string | number | boolean | null | undefined) => {
      if (v !== undefined && v !== null && v !== '') headers[k] = String(v);
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const atts: any[] = [];
    if (item.kind === 'appt') {
      set('X-UM-Class', 'IPM.Appointment');
      set('X-UM-Appt-Start', item.start);
      set('X-UM-Appt-End', item.end);
      set('X-UM-Appt-Location', item.location);
      set('X-UM-Appt-Organizer', item.organizer);
      set('X-UM-Appt-Attendees', item.attendees);
      set('X-UM-Appt-Uid', item.uid);
      set('X-UM-Appt-Seq', item.seq ?? 0);
      set('X-UM-Appt-AllDay', item.allDay ? '1' : '');
      if (item.reminderMin != null) set('X-UM-Appt-Reminder', item.reminderMin);
      const icsLocal = (iso: string): string => {
        const d = new Date(iso);
        if (isNaN(d.getTime())) return '';
        const p = (n: number) => String(n).padStart(2, '0');
        return item.allDay
          ? `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`
          : `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}T${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
      };
      // UID vem do chamador quando há convites a enviar — o mesmo UID vai
      // nos REQUESTs, para que respostas/atualizações correlacionem.
      const uid = item.uid || `${randomUUID()}@unitmail`;
      const dt = (v?: string) => (v ? icsLocal(v) : '');
      const ics = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//Unit Mail//Compromisso//PT-BR',
        'METHOD:PUBLISH',
        'BEGIN:VEVENT',
        `UID:${uid}`,
        `DTSTAMP:${icsLocal(new Date().toISOString())}`,
        item.allDay ? `DTSTART;VALUE=DATE:${dt(item.start)}` : `DTSTART:${dt(item.start)}`,
        item.allDay ? `DTEND;VALUE=DATE:${dt(item.end)}` : `DTEND:${dt(item.end)}`,
        `SUMMARY:${(item.subject || '').replace(/\r?\n/g, ' ')}`,
        item.location ? `LOCATION:${item.location.replace(/\r?\n/g, ' ')}` : '',
        (() => {
          // "Nome <e@x>" → ORGANIZER;CN=Nome:mailto:e@x (formato iCalendar)
          const m = /^(.*?)\s*<([^>\s]+)>\s*$/.exec(item.organizer || '');
          return m
            ? `ORGANIZER;CN=${(m[1].trim() || m[2]).replace(/["\\;:,]/g, ' ')}:mailto:${m[2]}`
            : item.organizer
              ? `ORGANIZER:mailto:${item.organizer}`
              : '';
        })(),
        item.body ? `DESCRIPTION:${item.body.replace(/\r?\n/g, '\\n')}` : '',
        item.reminderMin != null
          ? `BEGIN:VALARM\nTRIGGER:-PT${item.reminderMin}M\nACTION:DISPLAY\nEND:VALARM`
          : '',
        'END:VEVENT',
        'END:VCALENDAR',
      ]
        .filter(Boolean)
        .join('\r\n');
      atts.push({
        filename: 'compromisso.ics',
        contentType: 'text/calendar; method=PUBLISH; charset=utf-8',
        content: Buffer.from(ics, 'utf8'),
      });
    } else if (item.kind === 'contact') {
      set('X-UM-Class', 'IPM.Contact');
      set('X-UM-Contact-Name', item.contactName || item.subject);
      set('X-UM-Contact-Given', item.givenName);
      set('X-UM-Contact-Surname', item.surname);
      set('X-UM-Contact-Company', item.company);
      set('X-UM-Contact-JobTitle', item.jobTitle);
      set('X-UM-Contact-Dept', item.dept);
      set('X-UM-Contact-Email', item.email);
      set('X-UM-Contact-Email2', item.email2);
      set('X-UM-Contact-Phone', item.phone);
      set('X-UM-Contact-Mobile', item.mobile);
      set('X-UM-Contact-Birthday', item.birthday);
    } else {
      set('X-UM-Class', 'IPM.Task');
      set('X-UM-Task-Start', item.startDate);
      set('X-UM-Task-Due', item.dueDate);
      set('X-UM-Task-Pct', item.percentComplete ?? 0);
      set('X-UM-Task-Complete', item.isComplete ? '1' : '0');
      set('X-UM-Task-Completed', item.isComplete ? new Date().toISOString() : '');
      set('X-UM-Task-Owner', item.owner);
      set('X-UM-Task-Status', item.isComplete ? 2 : item.percentComplete && item.percentComplete > 0 ? 1 : 0);
    }

    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const MailComposer = require('nodemailer/lib/mail-composer');
    const raw = await new Promise<Buffer>((resolve, reject) => {
      new MailComposer({
        // Compromisso: organizador como remetente ("De" da lista/do inspetor)
        // e participantes no "Para" — antes saía sem From: e a linha ficava vazia.
        from: item.kind === 'appt' && item.organizer ? item.organizer : undefined,
        to: item.kind === 'appt' && item.attendees ? item.attendees : undefined,
        subject: item.subject || '(sem título)',
        date: new Date(),
        text: item.body || '',
        headers,
        attachments: atts,
      })
        .compile()
        .build((e: Error | null, msg: Buffer) => (e ? reject(e) : resolve(msg)));
    });
    return raw;
  }

  /**
   * Edita um item especial existente — regrava o MESMO arquivo .eml (posição
   * na lista e estado de leitura preservados; o mtime novo invalida os
   * caches de resumo/pesquisa automaticamente). Compromissos mantêm o UID
   * iCalendar original e incrementam a SEQUENCE — o reenvio do convite
   * atualiza a reunião nos clientes dos participantes em vez de criar outra.
   * Retorna o item efetivamente gravado (com uid/seq resolvidos).
   */
  async updateSpecialItem(
    folderId: string,
    index: number,
    item: NewSpecialItem,
    stableId?: string,
  ): Promise<NewSpecialItem> {
    this.requireOpen();
    index = this.idxFor(folderId, index, stableId);
    let arcId: string;
    if (folderId.startsWith('arc:')) {
      arcId = folderId;
    } else {
      const arcCount = this.mergedArcCount(folderId);
      if (index >= arcCount) {
        throw new Error('Só é possível editar itens criados nesta base');
      }
      arcId = `arc:${this.nodeFor(folderId)!.archiveName}`;
    }
    const file = this.archiveFiles(arcId)[index];
    if (!file) throw new Error('Item não encontrado');
    const fp = join(this.archiveDirFor(arcId), file);
    const merged = { ...item };
    try {
      const parsed = await simpleParser(readFileSync(fp));
      const hdr = (n: string): string => {
        const v = parsed.headers?.get?.(n);
        return Array.isArray(v) ? String(v[0] ?? '') : v != null ? String(v) : '';
      };
      if (item.kind === 'appt') {
        merged.uid = hdr('x-um-appt-uid') || item.uid;
        merged.seq = (Number(hdr('x-um-appt-seq')) || 0) + 1;
        merged.organizer = hdr('x-um-appt-organizer') || item.organizer;
      }
    } catch {
      // arquivo ilegível — segue com os valores recebidos
    }
    const raw = await this.buildSpecialEml(merged);
    writeFileSync(fp, raw);
    this.fsChanged();
    this.recountNamedFolder(arcId.slice(4));
    return merged;
  }

  /**
   * Grava um .eml numa pasta nomeada (arc:<nome>) já com estado de leitura
   * definido. Usado para "Caixa de Saída"/"Itens Enviados" — mensagens que o
   * próprio usuário enviou nunca aparecem como não lidas.
   * Retorna o nome do arquivo gravado.
   */
  importNamedEml(folderName: string, raw: Buffer, read: boolean): string {
    this.requireOpen();
    const arcId = `arc:${safeFsName(folderName)}`;
    const dir = this.writableDirFor(arcId);
    if (!dir) throw new Error('A pasta de destino não aceita itens');
    const name = `${this.emlDateMs(raw) ?? Date.now()}-${randomUUID().slice(0, 8)}.eml`;
    writeFileSync(join(dir, name), raw);
    this.fsChanged();
    const st = this.loadReadState();
    st.set(this.readKey(arcId, name), read);
    this.saveReadState();
    this.recountNamedFolder(arcId.slice(4));
    return name;
  }

  /**
   * Move um .eml entre pastas nomeadas do arquivo local — o fluxo do Outlook
   * ao concluir o envio: Caixa de Saída → Itens Enviados. Preserva o estado
   * de leitura.
   */
  moveNamedFile(fromFolder: string, toFolder: string, fileName: string): void {
    const fromArc = `arc:${safeFsName(fromFolder)}`;
    const toArc = `arc:${safeFsName(toFolder)}`;
    const fromDir = this.writableDirFor(fromArc);
    const toDir = this.writableDirFor(toArc);
    if (!fromDir || !toDir) return;
    const from = join(fromDir, fileName);
    if (!existsSync(from)) return;
    renameSync(from, join(toDir, fileName));
    this.fsChanged();
    const st = this.loadReadState();
    const was = st.get(this.readKey(fromArc, fileName));
    st.delete(this.readKey(fromArc, fileName));
    st.set(this.readKey(toArc, fileName), was ?? true);
    this.saveReadState();
    this.moveItemCatKey(fromArc.slice(4), toArc.slice(4), fileName);
    this.recountNamedFolder(fromArc.slice(4));
    this.recountNamedFolder(toArc.slice(4));
  }

  /** Recalcula o contador de um nó da árvore cuja pasta local é <folderName>. */
  private recountNamedFolder(folderName: string): void {
    if (this.batchDepth > 0) return; // a árvore é reconstruída no fim do lote
    const walk = (nodes: FolderNode[]) => {
      for (const n of nodes) {
        const arcId = n.id.startsWith('arc:')
          ? n.id
          : n.archiveName
            ? `arc:${n.archiveName}`
            : null;
        if (arcId === `arc:${folderName}`) {
          const arcPart = this.archiveFiles(arcId).length;
          n.contentCount = n.id.startsWith('arc:')
            ? arcPart
            : (this.folders.get(n.id)?.contentCount || 0) + arcPart;
        }
        walk(n.children);
      }
    };
    walk(this.rootNodes);
  }

  /** Transfere a marcação de categorias quando um .eml muda de pasta. */
  private moveItemCatKey(fromRel: string, toRel: string, fileName: string): void {
    const m = this.loadItemCats();
    const cats = m.get(`a|${fromRel}/${fileName}`);
    if (!cats) return;
    m.delete(`a|${fromRel}/${fileName}`);
    if (cats.length) m.set(`a|${toRel}/${fileName}`, cats);
    this.saveItemCats();
  }

  /** Estado de leitura efetivo de um item — local (readstate) ou PST (+override). */
  isItemRead(folderId: string, index: number, stableId?: string): boolean {
    index = this.idxFor(folderId, index, stableId);
    let arcId = folderId;
    if (folderId.startsWith('arc:')) {
      const f = this.archiveFiles(folderId)[index];
      return f ? this.isReadFile(folderId, f) : true;
    }
    const arcCount = this.mergedArcCount(folderId);
    if (index < arcCount) {
      arcId = `arc:${this.nodeFor(folderId)!.archiveName}`;
      const f = this.archiveFiles(arcId)[index];
      return f ? this.isReadFile(arcId, f) : true;
    }
    const raw = index - arcCount;
    const ov = this.pstReadOverride(folderId, raw);
    if (ov !== undefined) return ov;
    const folder = this.folders.get(folderId);
    if (!folder) return true;
    folder.moveChildCursorTo(raw);
    const msg = folder.getNextChild();
    return !!msg?.isRead;
  }

  /** O item é um .eml local (pasta arc: ou parte mesclada de pasta do PST). */
  isLocalItem(folderId: string, index: number, stableId?: string): boolean {
    if (folderId.startsWith('arc:')) return true;
    // O prefixo do stableId já diz a origem — sem nem consultar o disco.
    if (stableId) {
      if (stableId.startsWith('a|')) return true;
      if (stableId.startsWith('p|')) return false;
    }
    index = this.idxFor(folderId, index, stableId);
    return index < this.mergedArcCount(folderId);
  }

  /** Remove um item. Item local é apagado do disco; item gravado no PST é
   *  ocultado (exclusão lógica — o arquivo é somente leitura). */
  deleteItem(folderId: string, index: number, stableId?: string): void {
    index = this.idxFor(folderId, index, stableId);
    let arcId = folderId;
    if (!folderId.startsWith('arc:')) {
      if (!this.isLocalItem(folderId, index)) {
        this.hideItem(folderId, index);
        return;
      }
      arcId = `arc:${this.nodeFor(folderId)!.archiveName}`;
    }
    const file = this.archiveFiles(arcId)[index];
    if (!file) throw new Error('Item não encontrado');
    const { unlinkSync } = require('fs') as typeof import('fs');
    unlinkSync(join(this.archiveDirFor(arcId), file));
    this.fsChanged();
    if (this.loadItemCats().delete(`a|${arcId.slice(4)}/${file}`)) this.saveItemCats();
    if (this.batchDepth > 0) {
      this.sumCache.clear();
      return; // a árvore é reconstruída no fim do lote
    }
    const dec = (nodes: FolderNode[]) => {
      for (const n of nodes) {
        if (n.id === folderId) {
          n.contentCount = folderId.startsWith('arc:')
            ? this.archiveFiles(arcId).length
            : (this.folders.get(folderId)?.contentCount || 0) + this.arcCountFor(n);
        }
        dec(n.children);
      }
    };
    dec(this.rootNodes);
    this.sumCache.clear();
  }

  /** Cria uma pasta gravável (local/arquivo), opcionalmente dentro de outra. */
  createArchiveFolder(name: string, parentFolderId?: string | null): void {
    if (!this.archiveDir) throw new Error('Base sem diretório de arquivamento');
    const clean = safeFsName(name, '');
    if (!clean) throw new Error('Nome de pasta inválido');
    let parentDir = this.archiveDir;
    if (parentFolderId) {
      // Pasta-mãe do PST ganha diretório-sombra se ainda não tiver
      parentDir = this.writableDirFor(parentFolderId) ?? this.archiveDir;
    }
    mkdirSync(join(parentDir, clean), { recursive: true });
    this.fsChanged();
    this.rebuildTree();
  }

  /**
   * Move uma pasta local para dentro de outra pasta (ou para a raiz quando
   * newParentId é null). Só funciona para pastas "arc:" desta base.
   */
  moveArchiveFolder(folderId: string, newParentFolderId: string | null): void {
    if (!folderId.startsWith('arc:')) throw new Error('Pasta do PST é somente leitura');
    if (isDefaultFolderId(folderId)) {
      throw new Error(`"${folderId.slice(4)}" é uma pasta padrão e não pode ser movida`);
    }
    if (!this.archiveDir) throw new Error('Base sem diretório de arquivamento');
    const srcDir = this.archiveDirFor(folderId);
    let parentDir = this.archiveDir;
    if (newParentFolderId) {
      const resolved = this.writableDirFor(newParentFolderId);
      if (!resolved) throw new Error('Pasta de destino não aceita subpastas');
      parentDir = resolved;
    }
    const destDir = join(parentDir, this.archiveRelPath(folderId).split('/').pop()!);
    if (srcDir === destDir) return;
    // Impede mover para dentro de si mesma ou de uma descendente
    if (destDir.startsWith(srcDir + (srcDir.includes('\\') ? '\\' : '/')) || destDir === srcDir) {
      throw new Error('Não é possível mover uma pasta para dentro dela mesma');
    }
    if (!existsSync(srcDir)) throw new Error('Pasta não encontrada');
    if (existsSync(destDir)) throw new Error('Já existe uma pasta com esse nome no destino');
    const { renameSync } = require('fs') as typeof import('fs');
    const oldRel = this.archiveRelPath(folderId);
    renameSync(srcDir, destDir);
    this.fsChanged();
    const { relative } = require('path') as typeof import('path');
    this.remapReadKeys(oldRel, relative(this.archiveDir, destDir).replace(/\\/g, '/'));
    this.rebuildTree();
    this.sumCache.clear();
  }

  /** Remapeia as chaves de estado de leitura quando uma pasta muda de lugar/nome. */
  private remapReadKeys(oldRel: string, newRel: string): void {
    if (!this.readState || oldRel === newRel) return;
    const prefix = `${oldRel}/`;
    const moves: Array<[string, string]> = [];
    for (const k of this.readState.keys()) {
      if (k.startsWith(prefix)) moves.push([k, `${newRel}/${k.slice(prefix.length)}`]);
    }
    for (const [o, n] of moves) {
      this.readState.set(n, this.readState.get(o)!);
      this.readState.delete(o);
    }
    if (moves.length) this.saveReadState();
  }

  /** Renomeia uma pasta local (arc:). */
  renameArchiveFolder(folderId: string, newName: string): void {
    if (!folderId.startsWith('arc:')) throw new Error('Pasta do PST é somente leitura');
    if (isDefaultFolderId(folderId)) {
      throw new Error(`"${folderId.slice(4)}" é uma pasta padrão e não pode ser renomeada`);
    }
    const name = newName.trim();
    if (!name) throw new Error('Nome inválido');
    if (/[\\/:*?"<>|]/.test(name) || /[. ]$/.test(name)) {
      throw new Error('O nome contém caracteres inválidos ou termina em ponto/espaço');
    }
    const srcDir = this.archiveDirFor(folderId);
    if (!existsSync(srcDir)) throw new Error('Pasta não encontrada');
    const destDir = join(srcDir, '..', name);
    if (existsSync(destDir)) throw new Error('Já existe uma pasta com esse nome');
    const oldRel = this.archiveRelPath(folderId);
    const { renameSync } = require('fs') as typeof import('fs');
    renameSync(srcDir, destDir);
    this.fsChanged();
    const { relative } = require('path') as typeof import('path');
    this.remapReadKeys(oldRel, relative(this.archiveDir!, destDir).replace(/\\/g, '/'));
    this.rebuildTree();
    this.sumCache.clear();
  }

  /** Exclui uma pasta local (arc:) com todo o conteúdo. */
  deleteArchiveFolder(folderId: string): void {
    if (!folderId.startsWith('arc:')) throw new Error('Pasta do PST é somente leitura');
    if (isDefaultFolderId(folderId)) {
      throw new Error(`"${folderId.slice(4)}" é uma pasta padrão e não pode ser excluída`);
    }
    const dir = this.archiveDirFor(folderId);
    if (!existsSync(dir)) throw new Error('Pasta não encontrada');
    const { rmSync } = require('fs') as typeof import('fs');
    rmSync(dir, { recursive: true, force: true });
    this.fsChanged();
    const prefix = `${this.archiveRelPath(folderId)}/`;
    if (this.readState) {
      for (const k of [...this.readState.keys()]) {
        if (k.startsWith(prefix)) this.readState.delete(k);
      }
      this.saveReadState();
    }
    this.rebuildTree();
    this.sumCache.clear();
  }

  /** Copia uma pasta local (arc:) para "<nome> - Cópia" no mesmo nível. */
  copyArchiveFolder(folderId: string): void {
    if (!folderId.startsWith('arc:')) throw new Error('Pasta do PST é somente leitura');
    const srcDir = this.archiveDirFor(folderId);
    if (!existsSync(srcDir)) throw new Error('Pasta não encontrada');
    const rel = this.archiveRelPath(folderId);
    const name = rel.split('/').pop()!;
    const parentRel = rel.split('/').slice(0, -1).join('/');
    let copyName = `${name} - Cópia`;
    let i = 2;
    while (existsSync(join(this.archiveDir!, parentRel, copyName))) {
      copyName = `${name} - Cópia ${i++}`;
    }
    const { cpSync } = require('fs') as typeof import('fs');
    cpSync(srcDir, join(this.archiveDir!, parentRel, copyName), { recursive: true });
    this.fsChanged();
    this.rebuildTree();
    this.sumCache.clear();
  }

  /** Esvazia uma pasta: apaga os .eml locais e oculta os itens gravados no
   *  PST (exclusão lógica — somente leitura). Retorna quantos sumiram. */
  emptyFolder(folderId: string): number {
    const arcId = folderId.startsWith('arc:')
      ? folderId
      : this.nodeFor(folderId)?.archiveName
        ? `arc:${this.nodeFor(folderId)!.archiveName}`
        : null;
    let removed = 0;
    if (arcId) {
      const dir = this.archiveDirFor(arcId);
      const { unlinkSync } = require('fs') as typeof import('fs');
      for (const f of this.archiveFiles(arcId)) {
        try {
          unlinkSync(join(dir, f));
          this.readState?.delete(this.readKey(arcId, f));
          removed++;
        } catch {
          // arquivo travado — ignora
        }
      }
      if (removed) this.fsChanged();
      this.saveReadState();
    }
    // Itens gravados no PST: oculta todos — a pasta fica vazia na interface
    if (!folderId.startsWith('arc:')) {
      const folder = this.folders.get(folderId);
      const total = folder?.contentCount || 0;
      const m = this.loadHidden();
      if (!m.has(folderId)) m.set(folderId, new Set());
      const set = m.get(folderId)!;
      for (let i = 0; i < total; i++) {
        if (!set.has(i)) {
          set.add(i);
          removed++;
        }
      }
      if (removed > 0) this.saveHidden();
    }
    this.sumCache.clear();
    const node = this.nodeFor(folderId);
    if (node) node.contentCount = 0;
    return removed;
  }

  /** Busca simples: varre todas as pastas do tipo indicado (PST + arquivo local). */
  /**
   * Índice de pesquisa persistido — para cada item guarda o texto pesquisável
   * (assunto + remetente + destinatários). Itens do PST são estáveis (o arquivo
   * é somente leitura) então o índice só é reconstruído se o .pst mudar de
   * mtime; itens .eml levam o mtime do arquivo na entrada.
   * Segunda busca em diante é praticamente instantânea.
   */
  /**
   * Versão do índice de busca — muda quando o "texto pesquisável" muda
   * (ex.: entrou preview/CC, normalização de acentos), forçando reindexação.
   */
  private searchIdx: {
    fileMtime: number;
    v?: number;
    items: Record<string, { t: string; m?: number }>;
  } | null = null;
  /** Geração da busca em curso — uma busca nova aborta a anterior. */
  private searchGen = 0;

  /**
   * Índice persistente de RESUMOS (o ".msf" do Thunderbird): mapeia cada
   * item para seu ItemSummary já calculado, sem index/isRead (dinâmicos).
   * Sem isso toda sessão re-parseia todos os .eml e re-itera o PST.
   *   'a|<rel>|<file>' → .eml local, invalidado por mtime
   *   'p|<folderId>|<i>' → item do PST, invalidado pelo mtime do arquivo
   */
  // Versão do formato do índice de resumos — bumpar quando campos novos
  // entram no ItemSummary (ex.: inviteMethod), pois entradas antigas não os têm.
  private static readonly SUMIDX_VERSION = 4;

  private sumIdx: {
    v?: number;
    fileMtime: number;
    items: Record<string, { m?: number; s: Omit<ItemSummary, 'index' | 'isRead'> }>;
  } | null = null;
  private sumIdxTimer: ReturnType<typeof setTimeout> | null = null;

  private loadSumIdx(): NonNullable<PstService['sumIdx']> {
    if (this.sumIdx) return this.sumIdx;
    const fm =
      !this.localOnly && this.filePath && existsSync(this.filePath)
        ? statSync(this.filePath).mtimeMs
        : 0;
    this.sumIdx = { v: PstService.SUMIDX_VERSION, fileMtime: fm, items: {} };
    try {
      if (this.archiveDir) {
        const p = join(this.archiveDir, 'sumindex.json');
        const raw = JSON.parse(readFileSync(p, 'utf8')) as typeof this.sumIdx;
        if (raw && raw.v === PstService.SUMIDX_VERSION && raw.fileMtime === fm && raw.items)
          this.sumIdx.items = raw.items;
      }
    } catch {
      // sem índice ainda — constrói do zero
    }
    return this.sumIdx;
  }

  private saveSumIdx(): void {
    if (!this.sumIdx || !this.archiveDir) return;
    try {
      mkdirSync(this.archiveDir, { recursive: true });
      writeFileSync(join(this.archiveDir, 'sumindex.json'), JSON.stringify(this.sumIdx));
    } catch {
      // índice é cache — falha não é crítica
    }
  }

  /** Gravação adiada — uma pasta com 20 mil itens escreve o índice uma vez. */
  private markSumIdxDirty(): void {
    if (this.sumIdxTimer) clearTimeout(this.sumIdxTimer);
    this.sumIdxTimer = setTimeout(() => {
      this.sumIdxTimer = null;
      this.saveSumIdx();
    }, 2000);
  }

  private loadSearchIdx(): NonNullable<PstService['searchIdx']> {
    if (this.searchIdx) return this.searchIdx;
    const fm =
      !this.localOnly && this.filePath && existsSync(this.filePath)
        ? statSync(this.filePath).mtimeMs
        : 0;
    this.searchIdx = { fileMtime: fm, v: SEARCH_IDX_VER, items: {} };
    try {
      if (this.archiveDir) {
        const p = join(this.archiveDir, 'searchindex.json');
        const raw = JSON.parse(readFileSync(p, 'utf8')) as typeof this.searchIdx;
        // v2: texto normalizado (sem acentos) + preview/CC — índice antigo é
        // reindexado para a pesquisa encontrar os mesmos itens dos critérios novos
        if (raw && raw.v === SEARCH_IDX_VER && raw.fileMtime === fm && raw.items)
          this.searchIdx.items = raw.items;
      }
    } catch {
      // sem índice ainda — constrói do zero
    }
    return this.searchIdx;
  }

  private saveSearchIdx(): void {
    if (!this.searchIdx || !this.archiveDir) return;
    try {
      mkdirSync(this.archiveDir, { recursive: true });
      writeFileSync(
        join(this.archiveDir, 'searchindex.json'),
        JSON.stringify(this.searchIdx),
      );
    } catch {
      // índice é otimização — falha não é crítica
    }
  }

  /** Texto pesquisável de um resumo — é o que o índice persiste.
   *  Normalizado (minúsculas, sem acentos) para "cafe" achar "café". */
  private searchText(s: {
    subject?: string;
    senderName?: string;
    senderEmailAddress?: string;
    displayTo?: string;
    displayCC?: string;
    preview?: string;
  }): string {
    return normSearch(
      `${s.subject || ''}\n${s.senderName || ''}\n${s.senderEmailAddress || ''}\n${s.displayTo || ''}\n${s.displayCC || ''}\n${s.preview || ''}`,
    );
  }

  async search(
    query: string,
    containerClass = 'IPF.Note',
    maxResults = 2000,
    onBatch?: (hits: SearchResult[]) => void,
    folderId?: string,
  ): Promise<SearchResult[]> {
    this.requireOpen();
    const gen = ++this.searchGen; // busca nova aborta a anterior
    // Multi-token estilo Outlook: "roger nota fiscal" exige todos os termos
    // (sem exigir ordem/adjacência), com acentos normalizados.
    const tokens = normSearch(query).split(/\s+/).filter(Boolean);
    const match = (text: string) => tokens.every((t) => text.includes(t));
    const idx = this.loadSearchIdx();
    const sumIdx = this.loadSumIdx();
    let dirty = false;
    const results: SearchResult[] = [];
    // Emite o que foi achado na pasta que acabou de ser varrida — live search
    let emitted = 0;
    const flushBatch = () => {
      if (onBatch && results.length > emitted) {
        onBatch(results.slice(emitted));
        emitted = results.length;
      }
    };
    const walk = async (nodes: FolderNode[]): Promise<void> => {
      for (const node of nodes) {
        if (results.length >= maxResults || gen !== this.searchGen) return;
        if (
          // Escopo "pasta atual" — outras pastas nem são varridas
          (!folderId || node.id === folderId) &&
          node.containerClass === containerClass &&
          node.contentCount > 0
        ) {
          // Pastas com sombra .eml: a parte local é varrida junto (mesclada ou não)
          const arcIds: string[] = [
            ...(node.id.startsWith('arc:') ? [node.id] : []),
            ...(node.archiveName ? [`arc:${node.archiveName}`] : []),
          ];
          for (const arcId of arcIds) {
            const files = this.archiveFiles(arcId);
            const dir = this.archiveDirFor(arcId);
            for (let i = 0; i < files.length && results.length < maxResults; i++) {
              if (gen !== this.searchGen) return;
              try {
                const fp = join(dir, files[i]);
                const mtime = statSync(fp).mtimeMs;
                const key = `a|${arcId.slice(4)}|${files[i]}`;
                const hit = idx.items[key];
                let text: string;
                let s: ItemSummary | null = null;
                if (hit && hit.m === mtime) {
                  text = hit.t; // índice válido — não re-parseia o arquivo
                } else {
                  // Reusa o cache de resumos quando já foi parseado na listagem
                  let c = this.sumCache.get(fp);
                  if (!c || c.mtimeMs !== mtime) {
                    // sumindex.json já persistido pela listagem — evita o
                    // re-parse do .eml na primeira pesquisa da sessão
                    const pre = sumIdx.items[key];
                    if (pre && pre.m === mtime) {
                      c = { mtimeMs: mtime, s: pre.s };
                      this.sumCache.set(fp, c);
                    } else {
                      const parsed = await simpleParser(readFileSync(fp));
                      const { index: _i, isRead: _r, ...rest } = this.summarizeParsed(
                        parsed,
                        0,
                        statSync(fp).size,
                      );
                      c = { mtimeMs: mtime, s: rest };
                      this.sumCache.set(fp, c);
                    }
                  }
                  s = { ...c.s, index: i, isRead: this.isReadFile(arcId, files[i]) };
                  idx.items[key] = { t: this.searchText(c.s), m: mtime };
                  dirty = true;
                  text = idx.items[key].t;
                }
                if (match(text)) {
                  if (!s) {
                    // Índice de resumos evita re-parsear só para exibir o hit
                    const pre = sumIdx.items[key];
                    if (pre && pre.m === mtime) {
                      s = { ...pre.s, index: i, isRead: this.isReadFile(arcId, files[i]) };
                    } else {
                      const parsed = await simpleParser(readFileSync(fp));
                      s = {
                        ...this.summarizeParsed(parsed, i, mtime ? statSync(fp).size : 0),
                        isRead: this.isReadFile(arcId, files[i]),
                      };
                    }
                  }
                  s.categories = this.catsOfKey(`a|${arcId.slice(4)}/${files[i]}`);
                  s.local = true;
                  s.stableId = `a|${this.canonShadow(arcId.slice(4))}/${files[i]}`;
                  results.push({ folderId: node.id, folderName: node.name, item: s });
                }
              } catch {
                // arquivo corrompido — pula
              }
            }
          }
          if (!node.id.startsWith('arc:')) {
            const folder = this.folders.get(node.id)!;
            const total = folder.contentCount;
            const hidden = this.hiddenSet(node.id);
            // Pasta inteira já indexada → a pesquisa consulta só o texto do
            // índice, sem percorrer o cursor do PST (a parte lenta). O resumo
            // para exibir o hit vem do sumindex ou do cursor posicionado.
            let allIdx = total > 0;
            if (allIdx) {
              for (let i = 0; i < total; i++) {
                if (!idx.items[`p|${node.id}|${i}`]) {
                  allIdx = false;
                  break;
                }
              }
            }
            if (allIdx) {
              for (let i = 0; i < total && results.length < maxResults; i++) {
                if (gen !== this.searchGen) return;
                if (hidden.has(i)) continue; // excluído p/ lixeira — some da busca
                const key = `p|${node.id}|${i}`;
                if (!match(idx.items[key].t)) continue;
                // Resumo real só para os hits — posiciona o cursor no item
                // (o sumindex omite isRead, então não dá para exibir por ele)
                folder.moveChildCursorTo(i);
                const child = folder.getNextChild();
                if (!child) continue;
                const s = this.summarize(child, i);
                const ov = this.pstReadOverride(node.id, i);
                if (ov !== undefined) s.isRead = ov;
                s.index = arcCountIdx(node, i);
                s.stableId = `p|${node.id}|${i}`;
                s.categories = this.catsOfKey(key);
                results.push({ folderId: node.id, folderName: node.name, item: s });
              }
            } else {
              const start = 0; // varredura completa — pesquisa precisa ver tudo
              folder.moveChildCursorTo(start);
              for (let i = start; i < total && results.length < maxResults; i++) {
                if (gen !== this.searchGen) return;
                const child = folder.getNextChild();
                if (!child) break;
                if (hidden.has(i)) continue;
                const key = `p|${node.id}|${i}`;
                const hit = idx.items[key];
                let text: string;
                let s: ItemSummary | null = null;
                if (hit) {
                  text = hit.t;
                } else {
                  // sumindex do último list — dispensa o summarize; o isRead
                  // fica para o caminho de exibição (que tem o child real)
                  const pre = sumIdx.items[key];
                  idx.items[key] = {
                    t: pre ? this.searchText(pre.s) : this.searchText((s = this.summarize(child, i))),
                  };
                  dirty = true;
                  text = idx.items[key].t;
                }
                if (match(text)) {
                  if (!s) s = this.summarize(child, i);
                  const ov = this.pstReadOverride(node.id, i);
                  if (ov !== undefined) s.isRead = ov;
                  s.index = arcCountIdx(node, i);
                  s.stableId = `p|${node.id}|${i}`;
                  s.categories = this.catsOfKey(`p|${node.id}|${i}`);
                  results.push({ folderId: node.id, folderName: node.name, item: s });
                }
              }
            }
          }
        }
        flushBatch();
        await walk(node.children);
      }
    };
    const arcCountIdx = (node: FolderNode, rawIdx: number) =>
      (node.archiveName ? this.archiveFiles(`arc:${node.archiveName}`).length : 0) + rawIdx;
    await walk(this.rootNodes);
    if (dirty) this.saveSearchIdx();
    return results;
  }

  /** Monta "Nome <email>; ..." a partir da tabela de recipients do item.
   *  recipientType MAPI: 1 = To, 2 = Cc, 3 = Bcc. */
  private recipientList(msg: any, type: number): string {
    const out: string[] = [];
    const n: number = msg.numberOfRecipients || 0;
    for (let i = 0; i < n; i++) {
      try {
        const r = msg.getRecipient(i);
        if (!r || r.recipientType !== type) continue;
        const addr = (r.smtpAddress || r.emailAddress || '').trim();
        const name = ((r as { displayName?: string }).displayName || '')
          .replace(/[<>]/g, '')
          .trim();
        if (name && addr) out.push(`${name} <${addr}>`);
        else if (addr) out.push(addr);
        else if (name) out.push(name);
      } catch {
        // recipient ilegível — ignora
      }
    }
    return out.join('; ');
  }

  private requireOpen(): void {
    if (!this.pst && !this.localOnly) throw new Error('Nenhum arquivo PST aberto');
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private summarize(msg: any, index: number): ItemSummary {
    const mc: string = msg.messageClass || 'IPM.Note';
    const s: ItemSummary = {
      index,
      messageClass: mc,
      subject: msg.subject || msg.displayName || '(sem assunto)',
      senderName: msg.senderName || '',
      senderEmailAddress: msg.senderEmailAddress || '',
      // PR_DISPLAY_TO/CC nem sempre existe no PST — alguns itens só têm a
      // tabela de recipients (MAPI). Sem o fallback o .eml exportado saía sem
      // To:/Cc: e a pasta de enviados mostrava "(sem destinatário)".
      displayTo: msg.displayTo || this.recipientList(msg, 1),
      displayCC: msg.displayCC || this.recipientList(msg, 2),
      sentTime: toIso(msg.clientSubmitTime),
      deliveryTime: toIso(msg.messageDeliveryTime),
      hasAttachments: !!msg.hasAttachments,
      isRead: !!msg.isRead,
      // PR_LAST_VERB_EXECUTED do Outlook: 102/103 = respondeu, 104-106 =
      // encaminhou — ícones ↩/→ na lista como no Outlook.
      isAnswered: msg.lastVerbExecuted === 102 || msg.lastVerbExecuted === 103,
      isForwarded: msg.lastVerbExecuted === 104 || msg.lastVerbExecuted === 105 || msg.lastVerbExecuted === 106,
      importance: msg.importance || 0,
      size: msg.messageSize || 0,
      preview: (msg.bodyPrefix || '').trim(),
    };
    if (mc.startsWith('IPM.Schedule.Meeting')) {
      s.inviteMethod = /Canc/i.test(mc) ? 'CANCEL' : /Resp/i.test(mc) ? 'REPLY' : 'REQUEST';
    }
    if (mc.startsWith('IPM.Appointment') || mc.startsWith('IPM.Schedule.Meeting')) {
      s.startTime = toIso(msg.startTime);
      s.endTime = toIso(msg.endTime);
      s.location = msg.location || '';
      s.reminderSet = !!msg.reminderSet;
      s.reminderMinutes =
        msg.reminderMinutesBeforeStart != null && msg.reminderMinutesBeforeStart !== 0x5ae980e1
          ? msg.reminderMinutesBeforeStart
          : 15;
    } else if (mc === 'IPM.Contact') {
      s.contactDisplayName =
        msg.displayName || `${msg.givenName || ''} ${msg.surname || ''}`.trim();
      s.contactEmail = msg.email1EmailAddress || '';
      s.contactPhone = msg.businessTelephoneNumber || msg.mobileTelephoneNumber || '';
      s.contactCompany = msg.companyName || '';
      s.subject = s.contactDisplayName || '(contato)';
    } else if (mc === 'IPM.Task') {
      s.dueDate = toIso(msg.getDateItem?.(PR_TASK_DUE_DATE));
      s.percentComplete = Math.round((msg.percentComplete || 0) * 100);
      s.isTaskComplete = !!msg.isTaskComplete;
    } else if (mc === 'IPM.StickyNote') {
      s.subject = (msg.body || '').split('\n')[0] || '(nota)';
    }
    return s;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private detail(msg: any, index: number): ItemDetail {
    const summary = this.summarize(msg, index);
    const attachments: AttachmentMeta[] = [];
    const n: number = msg.numberOfAttachments || 0;
    for (let i = 0; i < n; i++) {
      try {
        const att = msg.getAttachment(i) as PSTAttachment;
        const embedded = att.attachMethod === 5 || !!att.embeddedPSTMessage;
        attachments.push({
          index: i,
          filename: att.longFilename || att.filename || `anexo-${i}`,
          size: att.filesize || att.size || 0,
          mimeTag: att.mimeTag || '',
          contentId: att.contentId || '',
          isEmbeddedMessage: embedded,
          isInline: !!att.contentId,
        });
      } catch {
        // anexo ilegível — ignora
      }
    }

    // Processa imagens inline (cid:) embutindo como data URI
    let bodyHTML: string = msg.bodyHTML || '';
    if (bodyHTML && attachments.some((a) => a.isInline && a.contentId)) {
      for (const a of attachments) {
        if (!a.contentId) continue;
        try {
          const att = msg.getAttachment(a.index) as PSTAttachment;
          const stream = att.fileInputStream;
          if (!stream) continue;
          const size = att.filesize || stream.length.toNumber();
          if (size > 8 * 1024 * 1024) continue;
          const buf = Buffer.alloc(size);
          stream.readCompletely(buf);
          const dataUri = `data:${a.mimeTag || 'application/octet-stream'};base64,${buf.toString('base64')}`;
          const cid = a.contentId.replace(/[<>]/g, '');
          bodyHTML = bodyHTML.split(`cid:${a.contentId}`).join(dataUri);
          bodyHTML = bodyHTML.split(`cid:${cid}`).join(dataUri);
        } catch {
          // falha em imagem inline — mantém cid
        }
      }
    }

    const mc = summary.messageClass;
    const detail: ItemDetail = {
      ...summary,
      body: msg.body || '',
      bodyHTML,
      transportMessageHeaders: msg.transportMessageHeaders || '',
      attachments: attachments.filter((a) => !a.isInline),
    };

    if (mc.startsWith('IPM.Appointment') || mc.startsWith('IPM.Schedule.Meeting')) {
      detail.allAttendees = msg.allAttendees || '';
      detail.requiredAttendees = msg.requiredAttendees || '';
      detail.isRecurring = !!msg.isRecurring;
      detail.recurrencePattern = msg.recurrencePattern || '';
      detail.organizer = msg.senderName || '';
    }

    // Convite de reunião (PST): IPM.Schedule.Meeting.Request/Canc/Resp
    if (mc.startsWith('IPM.Schedule.Meeting')) {
      detail.invite = {
        method: /Canc/i.test(mc)
          ? 'CANCEL'
          : /Resp/i.test(mc)
            ? 'REPLY'
            : 'REQUEST',
        uid: String(msg.internetMessageId || msg.descriptorNodeId || index),
        summary: summary.subject,
        start: summary.startTime ?? null,
        end: summary.endTime ?? null,
        location: summary.location || '',
        organizer: msg.senderName || '',
        organizerEmail: msg.senderEmailAddress || '',
        ics: '',
      };
    }

    if (mc === 'IPM.Contact') {
      const str = (id: number): string => {
        try {
          return msg.getStringItem?.(id) || '';
        } catch {
          return '';
        }
      };
      const contact: ContactDetail = {
        displayName: summary.contactDisplayName || '',
        givenName: msg.givenName || '',
        middleName: msg.middleName || '',
        surname: msg.surname || '',
        company: msg.companyName || '',
        jobTitle: msg.title || '',
        department: msg.departmentName || '',
        email1: msg.email1OriginalDisplayName || msg.email1EmailAddress || '',
        email1Address: msg.email1EmailAddress || '',
        email2: msg.email2OriginalDisplayName || msg.email2EmailAddress || '',
        email2Address: msg.email2EmailAddress || '',
        email3: msg.email3OriginalDisplayName || msg.email3EmailAddress || '',
        email3Address: msg.email3EmailAddress || '',
        businessPhone: msg.businessTelephoneNumber || '',
        homePhone: msg.homeTelephoneNumber || '',
        mobilePhone: msg.mobileTelephoneNumber || '',
        businessAddress: str(PR_CONTACT_BUSINESS_ADDRESS) || msg.businessPostalAddress || '',
        homeAddress: str(PR_CONTACT_HOME_ADDRESS) || msg.homePostalAddress || '',
        birthday: toIso(msg.birthday),
        notes: msg.body || '',
      };
      detail.contact = contact;
    }

    if (mc === 'IPM.Task') {
      const task: TaskDetail = {
        status: msg.taskStatus || 0,
        percentComplete: Math.round((msg.percentComplete || 0) * 100),
        isComplete: !!msg.isTaskComplete,
        startDate: toIso(msg.getDateItem?.(PR_TASK_START_DATE)),
        dueDate: toIso(msg.getDateItem?.(PR_TASK_DUE_DATE)),
        dateCompleted: toIso(msg.taskDateCompleted),
        owner: msg.taskOwner || '',
        body: msg.body || '',
      };
      detail.task = task;
    }

    return detail;
  }
}
