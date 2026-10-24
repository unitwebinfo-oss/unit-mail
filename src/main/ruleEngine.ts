import { createHash, randomUUID } from 'crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'fs';
import { mkdir, open, readdir, readFile, writeFile } from 'fs/promises';
import { basename, dirname, join } from 'path';
import { sourceManager } from './sourceManager';
import { accountStore } from './accountStore';
import { archiveRootForPst } from './paths';
import { isLocalStorePath } from './localStore';
import { dataDir } from './dataDir';
import { log } from './logger';
import { recordRecipients } from './contactHistory';
import { scanEmailBuffer } from './antivirus';
import { blockedSenders } from './blockedSenders';
import { getHeader } from './mailHeaders';
import { appLanguage } from './language';
import { resolveFolderName, type FolderKind } from '../shared/folderNames';
import type { AccountConfig, MailRule, SyncProgress, SyncResult } from '../shared/types';
import { safeFsName } from '../shared/fsName';

function rulesPath(): string {
  return join(dataDir(), 'rules.json');
}

function statePath(): string {
  return join(dataDir(), 'imap-state.json');
}

function loadRules(): MailRule[] {
  try {
    if (existsSync(rulesPath())) {
      const rules: MailRule[] = JSON.parse(readFileSync(rulesPath(), 'utf-8'));
      // As regras "Entrega padrão — email" (criadas por versões antigas a
      // cada conta salva) são redundantes: account.deliveryPath/deliveryFolder
      // já é o destino padrão usado por ensureDeliveryStore/imapInboxDir.
      // Filtra para fora — some da aba Regras e não é mais persistida.
      const seenId = new Set<string>();
      return rules.filter((r) => {
        if (r.name?.startsWith(DELIVERY_RULE_PREFIX)) return false;
        if (r.id && seenId.has(r.id)) return false;
        if (r.id) seenId.add(r.id);
        return true;
      });
    }
  } catch {
    // corrompido — começa do zero
  }
  return [];
}

function saveRules(rules: MailRule[]): void {
  mkdirSync(dataDir(), { recursive: true });
  writeFileSync(rulesPath(), JSON.stringify(rules, null, 2));
}

/** Último UID processado por conta+pasta, para só avaliar mensagens novas. */
function loadImapState(): Record<string, Record<string, number>> {
  try {
    if (existsSync(statePath())) return JSON.parse(readFileSync(statePath(), 'utf-8'));
  } catch {
    // ignora
  }
  return {};
}

function saveImapState(state: Record<string, Record<string, number>>): void {
  mkdirSync(dataDir(), { recursive: true });
  writeFileSync(statePath(), JSON.stringify(state));
}

/**
 * UIDLs POP3 já processados para a conta — chaveado por E-MAIL (estável),
 * com merge dos mapas legados por id de conta. Usado pelo sync e pelo
 * "Testar" da tela de contas (para distinguir novas × já baixadas).
 */
export function pop3SeenUids(account: { id?: string; email?: string }): Set<string> {
  const state = loadImapState();
  const emailKey = (account.email || account.id || '').toLowerCase();
  const aliases = new Set<string>([
    ...(account.id ? [account.id] : []),
    ...(((state[`pop3alias:${emailKey}`] as unknown as { ids?: string[] })?.ids) || []),
  ]);
  const seen = new Set<string>();
  const mergeUids = (raw: unknown) => {
    if (Array.isArray(raw)) for (const u of raw) seen.add(u as string);
    else if (raw && typeof raw === 'object') for (const u of Object.keys(raw)) seen.add(u);
  };
  mergeUids(state[`pop3seen:${emailKey}`]?.uids);
  for (const id of aliases) mergeUids(state[`pop3seen:${id}`]?.uids);
  return seen;
}

export { matchRule } from './ruleMatch';
import { matchRule } from './ruleMatch';

/** Diretório físico de destino de uma regra (PST → arquivo local; base local → a própria pasta). */
function ruleDestDir(rule: MailRule): string {
  if (isLocalStorePath(rule.targetPstPath)) {
    return join(rule.targetPstPath, rule.targetFolder);
  }
  return join(archiveRootForPst(rule.targetPstPath), rule.targetFolder);
}

const DELIVERY_RULE_PREFIX = 'Entrega padrão — ';

/**
 * Diretório físico de uma pasta padrão pelo TIPO dentro de uma raiz de
 * pastas (base local ou arquivo sombra do PST): usa a pasta já existente
 * em QUALQUER idioma — nunca cria "Junk Email" ao lado de "Lixo Eletrônico"
 * nem renomeia o que já está no disco. Só cria com o nome do idioma atual
 * quando nenhuma variante existe.
 */
function kindDirInRoot(rootDir: string, kind: FolderKind): string {
  let names: string[] = [];
  try {
    names = readdirSync(rootDir, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => d.name);
  } catch {
    // raiz ainda não existe — cria com o nome localizado
  }
  return join(rootDir, safeFsName(resolveFolderName(names, kind, appLanguage())));
}

/** Raiz de pastas de uma base (loja local = a própria; PST = arquivo sombra). */
function storeFolderRoot(storePath: string): string {
  return isLocalStorePath(storePath) ? storePath : archiveRootForPst(storePath);
}

/** Diretório resolvido de uma pasta padrão da base (lixo, saída, enviados…). */
export function kindDirFor(storePath: string, kind: FolderKind): string {
  return kindDirInRoot(storeFolderRoot(storePath), kind);
}

/** Nome físico resolvido da pasta padrão — para APIs que recebem o nome. */
export function resolveStoreFolderName(storePath: string, kind: FolderKind): string {
  let names: string[] = [];
  try {
    names = readdirSync(storeFolderRoot(storePath), { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => d.name);
  } catch {
    // raiz ainda não existe
  }
  return resolveFolderName(names, kind, appLanguage());
}

/** Configurações → Geral → "Verificar vírus" (padrão: ligado; desligado no Linux). */
function virusScanEnabled(): boolean {
  // No Linux não há inspeção de antivírus — o usuário optou por não rodar
  // ClamAV/motor local; a filtragem de spam do servidor continua ativa.
  if (process.platform === 'linux') return false;
  try {
    const p = join(dataDir(), 'settings.json');
    if (existsSync(p)) {
      const s = JSON.parse(readFileSync(p, 'utf-8')) as { virusScan?: boolean };
      return s.virusScan !== false;
    }
  } catch {
    // ignora
  }
  return true;
}

/**
 * Dedupe de entrega por diretório: hash SHA-1 do .eml cru. Impede que uma
 * mensagem já presente na pasta seja gravada de novo — protege contra
 * re-download total do POP3 quando o estado de "já vistos" se perde
 * (conta recriada, dados movidos de pasta), e contra regras duplicadas.
 * O índice é montado sob demanda por pasta e mantido em memória.
 */
const DEDUP_FILE = '_dedup.json';

interface DedupIdx {
  /** nome do .eml → chaves (mid/fp/sha1 — mais de uma por arquivo) */
  files: Map<string, string[]>;
  /** chave → nome do .eml (para validar arquivo que sumiu) */
  byKey: Map<string, string>;
  dirty: boolean;
  /** Indexação em background de arquivos ainda não mapeados — importações aguardam. */
  pending?: Promise<void>;
}

const dedupCache = new Map<string, Promise<DedupIdx>>();
const dedupTimers = new Map<string, ReturnType<typeof setTimeout>>();

async function persistDedup(dir: string, idx: DedupIdx): Promise<void> {
  try {
    await writeFile(
      join(dir, DEDUP_FILE),
      JSON.stringify({ files: Object.fromEntries(idx.files) }),
    );
    idx.dirty = false;
  } catch {
    // sem permissão — dedupe continua em memória nesta sessão
  }
}

function scheduleDedupSave(dir: string, idx: DedupIdx): void {
  idx.dirty = true;
  const old = dedupTimers.get(dir);
  if (old) clearTimeout(old);
  dedupTimers.set(
    dir,
    setTimeout(() => {
      dedupTimers.delete(dir);
      void persistDedup(dir, idx);
    }, 3000),
  );
}

/** Grava os índices pendentes — chamado no fim de cada sincronização. */
async function flushDedup(): Promise<void> {
  for (const t of dedupTimers.values()) clearTimeout(t);
  dedupTimers.clear();
  for (const [dir, p] of dedupCache) {
    const idx = await p.catch(() => null);
    if (idx?.dirty) await persistDedup(dir, idx);
  }
}

function hashOf(buf: Buffer): string {
  return createHash('sha1').update(buf).digest('hex');
}

/**
 * Chaves-candidatas de um .eml para dedupe — a mensagem é considerada
 * duplicada quando QUALQUER chave casa:
 *  - `mid:<id>`    — Message-ID (estável em mensagens baixadas e no export
 *                    de PST, que passa a carregar o id original);
 *  - `fp:<sha1>`   — fingerprint de cabeçalhos estáveis (De/Para/Cc/Assunto/
 *                    Data, com e-mails e encoded-words normalizados). Cobre
 *                    .eml reconstruídos cujo Message-ID varia a cada export —
 *                    ex.: bases já importadas por versões antigas — e cópias
 *                    do mesmo e-mail vindas de protocolos/fontes diferentes;
 *  - `sha1:<hash>` — último recurso, quando não há cabeçalho aproveitável.
 */
function keysOfHead(head: string, rawForHash?: Buffer): string[] {
  const keys: string[] = [];
  const mid = getHeader(head, 'Message-ID').replace(/[<>]/g, '').trim().toLowerCase();
  if (mid) keys.push(`mid:${mid}`);
  const canonText = (s: string) => s.toLowerCase().replace(/\s+/g, ' ').trim();
  // Endereços: só os e-mails contam — "Nome <a@b>", "<a@b>" e "a@b" são iguais.
  const canonAddr = (s: string) => {
    const addrs = s.toLowerCase().match(/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g);
    return addrs ? addrs.sort().join(',') : canonText(s);
  };
  const dateRaw = getHeader(head, 'Date');
  const t = Date.parse(dateRaw);
  const date = isNaN(t) ? canonText(dateRaw) : new Date(t).toISOString();
  const fields = [
    canonAddr(getHeader(head, 'From')),
    canonAddr(getHeader(head, 'To')),
    canonAddr(getHeader(head, 'Cc')),
    canonText(getHeader(head, 'Subject')),
    date,
  ];
  if (fields.some((f) => f)) {
    keys.push(`fp:${hashOf(Buffer.from(fields.join('|'), 'utf8'))}`);
  }
  if (!keys.length && rawForHash) keys.push(`sha1:${hashOf(rawForHash)}`);
  return keys;
}

function keysOf(buf: Buffer): string[] {
  return keysOfHead(buf.subarray(0, Math.min(buf.length, 32768)).toString('latin1'), buf);
}

/** Extrai cabeçalhos básicos do .eml baixado — substitui o TOP do POP3 quando
 *  nenhuma regra precisa de envelope (economia de um roundtrip por mensagem). */
function headInfo(buf: Buffer): { from: string; to: string; subject: string; messageId: string } {
  const head = buf.subarray(0, Math.min(buf.length, 32768)).toString('latin1');
  const get = (re: RegExp) => {
    const m = re.exec(head);
    return m ? m[1].trim() : '';
  };
  return {
    from: get(/^from:[ \t]*([^\r\n]+)/im),
    to: get(/^to:[ \t]*([^\r\n]+)/im),
    subject: get(/^subject:[ \t]*([^\r\n]+)/im),
    messageId: get(/^message-id:[ \t]*<?([^>\r\n]+)>?/im),
  };
}

/**
 * Marcação de spam feita pelo servidor/provedor antes da entrega —
 * SpamAssassin/cPanel (X-Spam-Flag/Status/Level), gateways com
 * X-Spam-Detected e assunto reescrito ("*** SPAM *** …", "[SPAM] …").
 * A mensagem vai para "Lixo eletrônico" da base, como o filtro de lixo do
 * Outlook — antes das regras do usuário.
 */
function isSpamFlagged(raw: Buffer): boolean {
  const head = raw.subarray(0, Math.min(raw.length, 16384)).toString('latin1');
  const yes = (v: string) => /^yes\b/i.test(v.trim());
  if (yes(getHeader(head, 'X-Spam-Flag'))) return true;
  if (yes(getHeader(head, 'X-Spam-Status'))) return true;
  if (yes(getHeader(head, 'X-Spam-Detected'))) return true;
  const stars = (getHeader(head, 'X-Spam-Level').match(/\*/g) || []).length;
  if (stars >= 5) return true;
  const subj = getHeader(head, 'Subject').trim().toLowerCase();
  return (
    subj.startsWith('***spam***') ||
    subj.startsWith('**spam**') ||
    subj.startsWith('[spam]')
  );
}

/** Chaves de um .eml em disco: cabeçalho basta para mid/fp; sha1 do arquivo inteiro só sem headers. */
async function fileKeys(fp: string): Promise<string[] | null> {
  try {
    const fh = await open(fp, 'r');
    try {
      const head = Buffer.alloc(32768);
      const { bytesRead } = await fh.read(head, 0, 32768, 0);
      const keys = keysOfHead(head.subarray(0, bytesRead).toString('latin1'));
      if (keys.length) return keys;
    } finally {
      await fh.close();
    }
    return [`sha1:${hashOf(await readFile(fp))}`];
  } catch {
    return null; // arquivo sumiu entre o readdir e a leitura
  }
}

/**
 * Chaves da pasta, calculadas UMA vez por sessão. A leitura é assíncrona e em
 * lotes: pastas com milhares de .eml não podem travar o processo principal
 * (a janela ficava "não respondendo" na primeira entrega).
 */
/**
 * Índice anti-duplicata da pasta, PERSISTENTE (_dedup.json dentro do diretório
 * — viaja junto no backup/cópia da base). Carregar = ler um JSON + um readdir;
 * só os arquivos ainda não indexados são parseados (incremental). Antes, cada
 * sessão re-lera o cabeçalho de TODOS os .eml da pasta — numa base de ~30 mil
 * mensagens isso travava o primeiro Enviar/Receber por minutos.
 */
function dirKeys(dir: string): Promise<DedupIdx> {
  let p = dedupCache.get(dir);
  if (!p) {
    p = (async () => {
      const files = new Map<string, string[]>();
      let stored: Record<string, string | string[]> = {};
      try {
        stored =
          (JSON.parse(await readFile(join(dir, DEDUP_FILE), 'utf8')) as {
            files?: Record<string, string | string[]>;
          }).files || {};
      } catch {
        // sem índice ainda — será montado agora e persistido
      }
      let names: string[] = [];
      try {
        names = (await readdir(dir)).filter((f) => f.endsWith('.eml'));
      } catch {
        // diretório ainda não existe — vazio
      }
      // Reconcilia: índice só vale para arquivo que ainda existe no disco.
      // Formato antigo (uma chave por arquivo) é promovido para lista.
      for (const f of names) {
        const v = stored[f];
        if (v) files.set(f, Array.isArray(v) ? v : [v]);
      }
      // Arquivos sem entrada no índice (importados, movidos manualmente, ou
      // pasta de uma instalação antiga sem _dedup.json) → indexados em
      // BACKGROUND. A entrega não pode esperar a varredura de dezenas de
      // milhares de arquivos — era o "travado na primeira mensagem" numa
      // base restaurada. Durante a janela de montagem, o mapa de vistos +
      // Message-IDs em memória cobre a deduplicação.
      const missing = names.filter((f) => !files.has(f));
      const idx: DedupIdx = {
        files,
        byKey: new Map(
          [...files].flatMap(([f, ks]) => ks.map((k) => [k, f] as [string, string])),
        ),
        dirty: missing.length > 0 || Object.keys(stored).length !== files.size,
      };
      if (missing.length > 0) {
        const build = async () => {
          for (let i = 0; i < missing.length; i += 64) {
            const slice = missing.slice(i, i + 64);
            const batch = await Promise.all(slice.map((f) => fileKeys(join(dir, f))));
            slice.forEach((f, j) => {
              const ks = batch[j];
              if (ks?.length) {
                files.set(f, ks);
                for (const k of ks) idx.byKey.set(k, f);
              }
            });
          }
          scheduleDedupSave(dir, idx);
        };
        idx.pending = build();
      }
      return idx;
    })();
    dedupCache.set(dir, p);
  }
  return p;
}

/**
 * Aguarda a indexação em background terminar antes de consultar o índice.
 * Sem isso, num re-download total (estado de "já vistos" zerado/restaurado
 * sem _dedup.json), as primeiras mensagens entregues passam batido pela
 * dedupe — gravadas como duplicatas não lidas.
 */
async function dedupReady(idx: DedupIdx): Promise<void> {
  if (idx.pending) {
    await idx.pending.catch(() => {});
    idx.pending = undefined;
  }
}

/** A chave já existe na pasta E o arquivo correspondente ainda está no disco. */
function dedupHas(idx: DedupIdx, dir: string, key: string): boolean {
  const f = idx.byKey.get(key);
  if (!f) return false;
  if (existsSync(join(dir, f))) return true;
  // Entrada órfã (arquivo apagado por fora) — limpa e trata como "não existe"
  idx.byKey.delete(key);
  idx.files.delete(f);
  return false;
}

/**
 * A mensagem já existe na pasta `dir`? Checagem de dedupe usada pelas
 * importações (PST/Thunderbird) antes de gravar — reimportar a mesma fonte
 * (ou importar sobre uma base que já tem parte das mensagens) não duplica.
 */
export async function dedupContains(dir: string, raw: Buffer): Promise<boolean> {
  try {
    const idx = await dirKeys(dir);
    // Importação não pode correr à frente da indexação: numa pasta sem
    // _dedup.json (base criada por versão antiga ou montada manualmente) o
    // índice ainda está sendo montado — responder agora daria falso negativo
    // e a mensagem seria gravada duplicada. Ao contrário da entrega (que tem
    // o conjunto de vistos em memória), aqui não há outra rede de segurança.
    await dedupReady(idx);
    return keysOf(raw).some((k) => dedupHas(idx, dir, k));
  } catch {
    return false;
  }
}

/**
 * Registra no índice de dedupe da pasta um .eml gravado por fora do
 * writeEmlDedup (ex.: importItemEml das importações) — a mensagem passa a
 * ser reconhecida em reimportações e por entregas futuras.
 */
export async function dedupNote(dir: string, fileName: string, raw: Buffer): Promise<void> {
  try {
    const idx = await dirKeys(dir);
    await dedupReady(idx);
    const ks = keysOf(raw);
    idx.files.set(fileName, ks);
    for (const k of ks) idx.byKey.set(k, fileName);
    scheduleDedupSave(dir, idx);
  } catch {
    // índice é best-effort — não pode atrapalhar a importação
  }
}

/**
 * Grava o .eml em `dir` só se ainda não existir uma cópia idêntica lá.
 * Retorna true quando a mensagem está na pasta (já estava ou gravou agora),
 * false quando a gravação falhou.
 * `accountEmail` carimba X-UM-Account no arquivo — prova qual conta recebeu
 * a mensagem (o cabeçalho To não prova: BCC e listas ficam invisíveis). O
 * "Executar regras" usa esse carimbo para o filtro por conta.
 * A chave de dedupe é calculada no buffer ORIGINAL — o carimbo não muda a
 * identidade da mensagem para fins de deduplicação.
 */
async function writeEmlDedup(
  dir: string,
  fileName: string,
  buf: Buffer,
  accountEmail?: string,
): Promise<boolean> {
  try {
    const idx = await dirKeys(dir);
    // Entrega também não pode correr à frente da indexação — o "mapa de
    // vistos" não a cobre (ele só marca DEPOIS de entregue), então sem o
    // await um re-download total duplica tudo que ainda não foi indexado.
    await dedupReady(idx);
    const ks = keysOf(buf);
    // já entregue — conta como sucesso
    if (ks.some((k) => dedupHas(idx, dir, k))) return true;
    await mkdir(dir, { recursive: true });
    const out =
      accountEmail && /^[^\r\n]+$/.test(accountEmail)
        ? Buffer.concat([Buffer.from(`X-UM-Account: ${accountEmail}\r\n`, 'latin1'), buf])
        : buf;
    await writeFile(join(dir, fileName), out);
    idx.files.set(fileName, ks);
    for (const k of ks) idx.byKey.set(k, fileName);
    scheduleDedupSave(dir, idx);
    return true;
  } catch {
    return false;
  }
}

/** Arquivo de dados marcado como "principal" nas Configurações (fallback de entrega). */
export function defaultStorePath(): string {
  try {
    const p = join(dataDir(), 'settings.json');
    if (existsSync(p)) {
      const s = JSON.parse(readFileSync(p, 'utf-8')) as { defaultStorePath?: string };
      if (s.defaultStorePath && existsSync(s.defaultStorePath)) return s.defaultStorePath;
    }
  } catch {
    // ignora
  }
  return '';
}

/** Caminho com sintaxe de OUTRO sistema operacional ("D:\…" no Linux,
 *  "/home/…" no Windows) — nunca é destino válido, mesmo que exista um
 *  diretório fantasma criado por engano com esse nome literal. */
export function foreignPath(p: string): boolean {
  return process.platform === 'win32' ? p.startsWith('/') : /^[a-zA-Z]:[\\/]/.test(p);
}

/**
 * Base de entrega de uma conta POP3: a escolhida pelo usuário ou, se nenhuma,
 * uma base local criada automaticamente (equivalente ao PST padrão do Outlook).
 * Retorna o diretório físico onde gravar a pasta.
 */
export function ensureDeliveryStore(account: AccountConfig): string {
  // IMAP já é o armazenamento da conta. Não crie nem escolha uma base local
  // silenciosamente para ela; a mensagem permanece no servidor IMAP.
  if (account.protocol === 'imap') return '';
  // deliveryPath só vale se a pasta existe NESTA máquina — um backup
  // restaurado de outro SO traz caminho inválido ("D:\…" no Linux), e
  // gravar nele criava um diretório fantasma fora da base ('\' é nome
  // válido no Linux!). Sem base válida, a entrega falha com mensagem
  // clara em vez de sumir num caminho errado.
  if (account.deliveryPath && existsSync(account.deliveryPath) && !foreignPath(account.deliveryPath))
    return account.deliveryPath;
  // Sem base vinculada à conta → usa o arquivo de dados marcado como principal.
  // PST nunca é destino de entrega: é somente leitura — serve só para importar.
  const principal = defaultStorePath();
  if (principal && !principal.toLowerCase().endsWith('.pst')) {
    try {
      // Abre o tipo certo — openLocalStore num .pst registraria o arquivo como
      // base local e a entrega depois tentaria criar pastas DENTRO do arquivo
      if (principal.toLowerCase().endsWith('.pst')) sourceManager.openPst(principal);
      else sourceManager.openLocalStore(principal);
    } catch {
      // já aberta ou falha não fatal
    }
    return principal;
  }
  // Sem base vinculada nem principal: NÃO cria base nativa automática — o
  // usuário escolhe o destino em Configurações → Email / Arquivos de Dados.
  return '';
}

/** Diretório físico de uma pasta dentro de uma base (local → subpasta; PST → arquivo vinculado). */
export function folderDirFor(storePath: string, folderName: string): string {
  const dir = safeFsName(folderName);
  return isLocalStorePath(storePath)
    ? join(storePath, dir)
    : join(archiveRootForPst(storePath), dir);
}

export class RuleEngine {
  /** Callback de progresso (registrado pelo index.ts para repassar ao renderer). */
  onProgress: ((ev: SyncProgress) => void) | null = null;
  private cancelFlag = false;
  /** taskIds cancelados individualmente (botão "Cancelar tarefa" do diálogo). */
  private cancelledTasks = new Set<string>();
  /**
   * Tarefa em execução + como abortar a operação de rede em voo. Sem isso o
   * cancelamento só era lido ENTRE mensagens — um RETR/download travado
   * ignorava o botão até o timeout de 60s (o "cancelar não funciona").
   */
  private currentTask: { taskId: string; abort: () => void } | null = null;

  /** Cancela o envio/recebimento em andamento ("Cancelar tudo" do diálogo). */
  cancelSync(): void {
    this.cancelFlag = true;
    this.currentTask?.abort();
  }

  /** Cancela uma tarefa específica pelo taskId (ex.: "acct-123:recv"). */
  cancelTask(taskId: string): void {
    this.cancelledTasks.add(taskId);
    if (this.currentTask?.taskId === taskId) this.currentTask.abort();
  }

  /**
   * Esquece o estado de "já processadas" da conta — o próximo Enviar/Receber
   * re-baixa TUDO que ainda estiver no servidor (POP3: mapa de UIDLs +
   * Message-IDs + aliases; IMAP: último UID + pendentes). A dedupe por pasta
   * (_dedup.json) impede duplicar o que já está na base — o re-download só
   * grava o que de fato falta. É a recuperação para quando um bug antigo
   * marcou mensagens como entregues sem gravar (ex.: diretório "D:\…"
   * fantasma num backup restaurado no Linux).
   */
  resetSyncState(accountId: string): void {
    const account = accountStore.getFull(accountId);
    const emailKey = (account?.email || accountId).toLowerCase();
    const state = loadImapState();
    // Aliases ANTES de apagar a chave-mãe — os ids legados por accountId
    // também guardam mapas de vistos.
    const aliases = new Set<string>([
      accountId,
      ...(((state[`pop3alias:${emailKey}`] as unknown as { ids?: string[] })?.ids) || []),
    ]);
    for (const id of aliases) delete state[`pop3seen:${id}`];
    delete state[`pop3seen:${emailKey}`];
    delete state[`pop3mid:${emailKey}`];
    delete state[`pop3alias:${emailKey}`];
    delete state[accountId]; // IMAP: lastUid por pasta + fila de retry
    saveImapState(state);
    log.info('sync', `estado de sincronização da conta ${emailKey} reiniciado — próximo envio/recebimento re-baixa o servidor`);
  }

  private isCancelled(taskId: string): boolean {
    return this.cancelFlag || this.cancelledTasks.has(taskId);
  }

  private emit(ev: SyncProgress): void {
    try {
      this.onProgress?.(ev);
    } catch {
      // renderer pode estar fechado
    }
  }

  list(): MailRule[] {
    return loadRules();
  }

  upsert(rule: Omit<MailRule, 'id'> & { id?: string }): MailRule {
    const rules = loadRules();
    const full: MailRule = { ...rule, id: rule.id || `rule-${randomUUID()}` };
    const idx = rules.findIndex((r) => r.id === full.id);
    if (idx >= 0) rules[idx] = full;
    else rules.push(full);
    saveRules(rules);
    return full;
  }

  remove(id: string): void {
    saveRules(loadRules().filter((r) => r.id !== id));
  }

  /**
   * Sincroniza uma conta IMAP: lê mensagens novas na caixa de entrada e
   * arquiva cópias .eml nas pastas locais vinculadas aos PSTs de destino,
   * conforme as regras. Não remove nem marca nada no servidor.
   */
  /** Uma sincronização por vez: timer automático e botão manual não se sobrepõem. */
  private queue: Promise<unknown> = Promise.resolve();

  private serialize<T>(fn: () => Promise<T>): Promise<T> {
    const run = this.queue.then(fn, fn);
    this.queue = run.catch(() => {});
    return run;
  }

  syncAccount(accountId: string, resetCancel = true): Promise<SyncResult> {
    return this.serialize(() => this.syncAccountNow(accountId, resetCancel));
  }

  private async syncAccountNow(accountId: string, resetCancel = true): Promise<SyncResult> {
    if (resetCancel) {
      this.cancelFlag = false;
      this.cancelledTasks.clear();
    }
    const sourceId = `acct:${accountId}`;
    const backend = sourceManager.backend(sourceId);
    const result: SyncResult = { accountId, scanned: 0, archived: 0 };

    // Cast para acesso aos métodos extras de sync (fora do contrato MailBackend)
    const imap = backend as unknown as {
      fetchNewEnvelopes(
        folderId: string,
        sinceUid: number,
      ): Promise<Array<{ uid: number; from: string; to: string; subject: string }>>;
      downloadRaw(folderId: string, uid: number): Promise<Buffer>;
      getFolderTree(): Promise<Array<{ id: string; name: string }>>;
      abortNow?(): void;
    };
    // O cancelamento aborta a operação de rede em voo (RETR/FETCH travado) —
    // cada caminho atribui o abort do seu transporte. Sem isso o botão só
    // fazia efeito entre uma mensagem e outra.
    let abortInFlight: (() => void) | null = null;
    const recvTaskId = `${accountId}:recv`;
    this.currentTask = { taskId: recvTaskId, abort: () => abortInFlight?.() };

    const account = accountStore.getFull(accountId);
    const recvTask = `${accountId}:recv`;
    const email = account?.email || accountId;
    this.emit({ taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing' });
    const all = loadRules().filter((r) => r.enabled && (r.accountId === accountId || r.accountId === 'all'));
    // Separa a regra automática "entregar mensagens em" das regras do usuário:
    // ela é o destino PADRÃO e só roda quando nenhuma regra do usuário casou —
    // senão a mensagem caía na pasta da regra E na Caixa de Entrada (duplicada).
    const userRules = all.filter((r) => !r.name.startsWith(DELIVERY_RULE_PREFIX));
    const deliveryRules = all.filter((r) => r.name.startsWith(DELIVERY_RULE_PREFIX));
    const virusScan = virusScanEnabled();
    // IMAP: só vinculação EXPLÍCITA conta como destino — a base principal do
    // app NUNCA recebe cópia de mensagens IMAP (elas moram no servidor).
    const hasDeliveryTarget = deliveryRules.length > 0 || !!account?.deliveryPath;
    // IMAP sem nenhum destino/regra: não entrega nada, mas ainda DETECTA o
    // que chegou novo no servidor (envelopes apenas — nenhum corpo baixado).
    // É o que faz o badge de não-lidos e a notificação funcionarem como num
    // cliente normal. POP3 sempre baixa (entrega padrão cria base se preciso).
    if (userRules.length === 0 && !hasDeliveryTarget && account?.protocol !== 'pop3') {
      if (account?.protocol === 'imap') {
        const st = loadImapState();
        st[accountId] = st[accountId] || {};
        const lastUid = st[accountId]['INBOX'] || 0;
        const fresh = await imap.fetchNewEnvelopes('INBOX', lastUid).catch(() => []);
        let maxUid = lastUid;
        for (const e of fresh) if (e.uid > maxUid) maxUid = e.uid;
        st[accountId]['INBOX'] = maxUid;
        saveImapState(st);
        result.scanned = fresh.length;
        result.newOnServer = fresh.length;
        this.emit({
          taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv',
          status: 'done', total: fresh.length, done: fresh.length,
        });
        this.currentTask = null;
        return result;
      }
      this.emit({ taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'done', total: 0, done: 0 });
      this.currentTask = null;
      return result;
    }

    type DeliverResult = 'delivered' | 'failed' | 'noMatch';
    /**
     * Avalia regras na ordem; cada regra que casa entrega uma cópia e,
     * se stopProcessing, encerra a avaliação. Retorna:
     *  - 'delivered': ao menos uma regra casou E gravou a mensagem
     *  - 'failed': regra casou mas a gravação falhou (tenta de novo depois)
     *  - 'noMatch': nenhuma regra casou — cai no destino padrão
     * Importante: uma regra que casa mas FALHA ao gravar não pode marcar a
     * mensagem como processada — senão ela se perde para sempre.
     */
    const deliverByRules = async (
      raw: () => Promise<Buffer>,
      env: { from: string; to: string; subject: string },
      rules: MailRule[],
    ): Promise<DeliverResult> => {
      let matched = false;
      for (const rule of rules) {
        if (!matchRule(rule, env)) continue;
        matched = true;
        // Nunca grava num destino que o usuário não escolheu mais: o alvo
        // precisa ser um PST (.pst) ou uma base local REGISTRADA. Uma base
        // fechada/removida deixa a regra falhar → a mensagem cai no destino
        // padrão da conta (ou fica no servidor se não houver).
        const t = (rule.targetPstPath || '').trim();
        // existsSync + foreignPath fecham o buraco do diretório fantasma:
        // uma base restaurada de outro SO pode deixar "D:\…" registrado em
        // localstores.json — isLocalStorePath devolvia true e a regra gravava
        // a mensagem num diretório literal "D:\…" fora da base.
        const targetOk =
          t && existsSync(t) && !foreignPath(t) &&
          (t.toLowerCase().endsWith('.pst') || isLocalStorePath(t));
        try {
          if (!targetOk) throw new Error('destino de regra inválido');
          const buf = await raw();
          const dir = ruleDestDir(rule);
          // Dedupe: cópia (mesmo Message-ID/fingerprint) já presente → entregue
          const ridx = await dirKeys(dir);
          await dedupReady(ridx);
          if (keysOf(buf).some((k) => dedupHas(ridx, dir, k))) return 'delivered';
          if (await writeEmlDedup(dir, `${Date.now()}-${randomUUID().slice(0, 8)}.eml`, buf, account?.email)) {
            result.archived++;
            return 'delivered';
          }
          throw new Error('falha ao gravar');
        } catch (e) {
          log.error(`regra "${rule.name}"`, e);
          // gravação falhou — tenta a próxima regra / destino padrão
        }
        if (rule.stopProcessing) break;
      }
      return matched ? 'failed' : 'noMatch';
    };

    /**
     * Entrega efetiva de uma mensagem: regras do usuário → regra de entrega
     * padrão da conta → Caixa de Entrada da base principal/vinculada.
     * Retorna true somente quando a mensagem foi gravada em algum lugar.
     */
    const deliverMessage = async (
      fetchRaw: () => Promise<Buffer>,
      env: { from: string; to: string; subject: string },
      uidTag: string | number,
      /** Resolve o diretório da Caixa de Entrada de destino (null = entrega opcional). */
      inboxDir: () => string | null,
      /**
       * POP3 exige destino (o servidor é só transporte — sem base configurada a
       * mensagem não pode ser considerada entregue). IMAP tolera null: a
       * mensagem fica no servidor sem penalidade.
       */
      inboxRequired = false,
    ): Promise<boolean> => {
      // O raw é baixado uma vez — a verificação de vírus e a entrega dividem
      // o mesmo buffer (sem dois RETR/FETCH por mensagem).
      let rawCache: Buffer | null = null;
      const getRaw = async () => {
        const r = (rawCache ??= await fetchRaw());
        // Envelope vazio (TOP pulado porque nenhuma regra tem condição):
        // preenche do próprio .eml — alimenta autocompletar e o dedupe por
        // Message-ID sem custo extra de rede.
        if (!env.from && !env.subject) Object.assign(env, headInfo(r));
        return r;
      };
      const dir = inboxDir();
      // Antivírus local (AMSI no Windows — cobre Defender e AVs de terceiros
      // registrados; ClamAV no Linux): mensagem infectada vai direto para
      // "Itens infectados", sem passar por regras nem Caixa de Entrada.
      if (virusScan && dir) {
        try {
          const raw = await getRaw();
          const av = await scanEmailBuffer(raw, `${uidTag}.eml`);
          if (av.status === 'infected') {
            const qdir = kindDirInRoot(dirname(dir), 'quarantine');
            const ok = await writeEmlDedup(
              qdir,
              `${Date.now()}-${uidTag}-virus-${randomUUID().slice(0, 8)}.eml`,
              raw,
              account?.email,
            );
            if (ok) {
              result.archived++;
              result.infected = (result.infected ?? 0) + 1;
              log.info(
                'antivirus',
                `mensagem de "${env.from}" movida para "${basename(qdir)}"` +
                  (av.threat ? ` — ${av.threat}` : ''),
              );
            }
            return ok;
          }
        } catch (e) {
          // Scanner falhou — a mensagem é entregue normalmente (não bloqueia email)
          log.error('antivirus: verificação falhou', e);
        }
      }
      // Marcação de spam do servidor/provedor (SpamAssassin, cPanel, etc.)
      // ou remetente na lista de bloqueados: vai direto para "Lixo
      // eletrônico" da base — como o filtro de lixo do Outlook, sem passar
      // por regras nem Caixa de Entrada.
      if (dir) {
        try {
          const raw = await getRaw();
          if (isSpamFlagged(raw) || blockedSenders.isBlocked(env.from)) {
            const jdir = kindDirInRoot(dirname(dir), 'junk');
            const ok = await writeEmlDedup(
              jdir,
              `${Date.now()}-${uidTag}-spam-${randomUUID().slice(0, 8)}.eml`,
              raw,
              account?.email,
            );
            if (ok) {
              result.archived++;
              result.junked = (result.junked ?? 0) + 1;
              const why = blockedSenders.isBlocked(env.from) ? 'remetente bloqueado' : 'spam do servidor';
              log.info('spam', `mensagem de "${env.from}" movida para "${basename(jdir)}" (${why})`);
            }
            return ok;
          }
        } catch (e) {
          // Falha na leitura do flag — segue o fluxo normal (não perde a mensagem)
          log.error('spam: verificação falhou', e);
        }
      }
      let st = await deliverByRules(getRaw, env, userRules);
      if (st === 'noMatch') st = await deliverByRules(getRaw, env, deliveryRules);
      if (st === 'delivered') {
        // Remetente entra no cache do autocompletar (Para/Cc/Cco).
        recordRecipients(env.from);
        return true;
      }
      // Sem destino padrão (IMAP sem base vinculada): a mensagem fica no
      // servidor — não é falha, é "não precisava baixar". No POP3 é falha:
      // sem base de entrega configurada a mensagem ficaria sem destino.
      if (!dir) return !inboxRequired;
      try {
        const raw = await getRaw();
        // Dedupe também no destino padrão — cópia idêntica já na pasta conta
        // como entregue (evita duplicar se o sync re-baixar mensagens antigas).
        const ok = await writeEmlDedup(
          dir,
          `${Date.now()}-${uidTag}-${randomUUID().slice(0, 8)}.eml`,
          raw,
          account?.email,
        );
        if (ok) {
          result.archived++;
          recordRecipients(env.from);
        }
        return ok;
      } catch (e) {
        log.error(`entrega (${uidTag})`, e);
        // Não entregou — NÃO marca como vista nem remove do servidor;
        // a mensagem será tentada de novo na próxima sincronização.
        return false;
      }
    };

    // ---- Caminho POP3: sessão única + controle de cópia no servidor ----
    if (account?.protocol === 'pop3') {
      const pop3 = backend as unknown as {
        openSyncSession(): Promise<{
          uidl(): Promise<Array<{ num: number; uidl: string }>>;
          envelope(num: number): Promise<{ uid: number; from: string; to: string; subject: string; messageId: string }>;
          retr(num: number): Promise<Buffer>;
          dele(num: number): Promise<void>;
          abort(): void;
          end(): Promise<void>;
        }>;
      };
      // uidl → timestamp da primeira vez que a mensagem foi processada.
      // A chave é o E-MAIL da conta (estável), não o id — o id muda quando a
      // conta é recriada e o estado antigo se perdia, fazendo o POP3 re-baixar
      // a caixa inteira como "nova" (milhares de duplicatas não lidas).
      const state = loadImapState();
      const seenKey = `pop3seen:${(account.email || accountId).toLowerCase()}`;
      const aliasKey = `pop3alias:${(account.email || accountId).toLowerCase()}`;
      // Message-IDs já entregues: segunda trava de dedupe — servidores que
      // geram UIDL novo a cada sessão (ou estado perdido) não re-baixam a
      // mensagem inteira: o TOP traz o Message-ID e ela é reconhecida aqui.
      const midKey = `pop3mid:${(account.email || accountId).toLowerCase()}`;
      const seenMids = new Set<string>(
        (((state[midKey] as unknown as { mids?: string[] })?.mids) || []).map((m) =>
          m.toLowerCase(),
        ),
      );
      const aliases = new Set<string>([
        accountId,
        ...(((state[aliasKey] as unknown as { ids?: string[] })?.ids) || []),
      ]);
      const seen: Record<string, number> = {};
      // Migração: junta o mapa novo (por e-mail) com os legados (por id de conta)
      const mergeUids = (raw: unknown) => {
        if (Array.isArray(raw)) for (const u of raw) seen[u as string] = Date.now();
        else if (raw && typeof raw === 'object') Object.assign(seen, raw);
      };
      mergeUids(state[seenKey]?.uids);
      for (const id of aliases) mergeUids(state[`pop3seen:${id}`]?.uids);

      // Grava o estado "já vistos" no disco — chamado em lotes durante o
      // sync (um reinício no meio não perde o progresso) e no final.
      const flushState = () => {
        // Limita o histórico de Message-IDs — é rede de segurança contra
        // UIDL novo por sessão; 50 mil mais recentes bastam. Sem teto o
        // imap-state.json crescia sem fim e cada flush regravava tudo.
        while (seenMids.size > 50000) {
          const first = seenMids.values().next().value;
          if (first === undefined) break;
          seenMids.delete(first);
        }
        state[seenKey] = { uids: seen } as unknown as Record<string, number>;
        state[aliasKey] = { ids: [...aliases] } as unknown as Record<string, number>;
        state[midKey] = { mids: [...seenMids] } as unknown as Record<string, number>;
        saveImapState(state);
      };
      let sinceFlush = 0;

      let session = await pop3.openSyncSession();
      // Cancelamento derruba a sessão → o RETR em voo rejeita na hora
      abortInFlight = () => {
        try {
          session.abort();
        } catch {
          // sessão já morta
        }
      };
      let undelivered = 0;
      // Destino padrão: base vinculada à conta (ou a principal) — nunca uma
      // base criada automaticamente. Sem destino, a mensagem fica no servidor.
      const popInboxDir = () => {
        const p = ensureDeliveryStore(account);
        return p
          ? account.deliveryFolder
            ? folderDirFor(p, account.deliveryFolder)
            : kindDirFor(p, 'inbox')
          : null;
      };
      // TOP só é necessário quando existe regra COM condições (avalia
      // assunto/remetente antes de baixar) ou para a trava por Message-ID.
      // Sem regras condicionais, pular o TOP corta um roundtrip por mensagem —
      // numa caixa grande é a diferença entre minutos e segundos. O envelope
      // é preenchido do .eml baixado (headInfo) para dedupe/autocompletar.
      const needEnvelope =
        userRules.some((r) => {
          const c = r.conditions || {};
          return !!(c.fromContains?.trim() || c.toContains?.trim() || c.subjectContains?.trim());
        }) ||
        deliveryRules.some((r) => {
          const c = r.conditions || {};
          return !!(c.fromContains?.trim() || c.toContains?.trim() || c.subjectContains?.trim());
        });
      // Se o servidor derrubar a sessão no meio do ciclo (rate-limit/timeout),
      // o worker reabre a PRÓPRIA sessão e remapeia os números via UIDL —
      // DELEs commitados por outra sessão podem mudar a numeração.
      let numByUidl = new Map<string, number>();
      try {
        const list = await session.uidl();
        for (const e of list) numByUidl.set(e.uidl, e.num);
        // Esquece UIDLs que já não existem no servidor (apagados por outro
        // cliente/webmail) — mantém o mapa de "vistos" enxuto em vez de
        // acumular para sempre.
        {
          const onServer = new Set(list.map((e) => e.uidl));
          for (const u of Object.keys(seen)) if (!onServer.has(u)) delete seen[u];
        }
        // O progresso conta só o que é NOVO — antes o "total" era a caixa
        // inteira do servidor, e o app parecia re-baixar tudo a cada sync.
        const pendingList = list.filter((e) => !seen[e.uidl]);
        log.info(
          'pop3',
          `${email}: ${list.length} mensagem(ns) no servidor — ${pendingList.length} nova(s) a baixar ` +
            `(${Object.keys(seen).length} já marcadas como vistas)`,
        );
        let done = 0;
        let lastEmit = 0;
        this.emit({ taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing', total: pendingList.length, done });
        // Pool de sessões paralelas: POP3 é estritamente sequencial por
        // conexão (comando → resposta), então numa caixa grande o tempo é
        // dominado pelo roundtrip de cada RETR, não pela banda. Várias
        // sessões dividem a fila e cortam o tempo proporcionalmente.
        // Cada worker tem sessão e mapa num↔uidl próprios — DELEs
        // commitados numa sessão podem mudar a numeração das outras.
        const liveSessions = new Set<{ abort(): void }>([session]);
        abortInFlight = () => {
          for (const s of liveSessions) {
            try {
              s.abort();
            } catch {
              // já morta
            }
          }
        };
        const emitProgress = (
          uidl: string,
          env: { from: string; subject: string },
          delivered: boolean,
        ) => {
          // Progresso com throttle (~150 ms): um evento por mensagem mandava
          // milhares de IPCs ao renderer — principal fonte da lentidão antiga.
          const now = Date.now();
          if (now - lastEmit >= 150 || done === pendingList.length) {
            lastEmit = now;
            this.emit({
              taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing',
              total: pendingList.length, done,
              item: { key: uidl, subject: env.subject, from: env.from, status: delivered ? 'ok' : 'fail', time: new Date().toISOString() },
            });
          }
        };
        // Fila compartilhada — cursor atômico (JS é single-threaded, basta ++).
        let cursor = 0;
        // Até 4 sessões extras além da principal. Servidor que recusa login
        // concorrente faz o worker extra sair na conexão — os demais cobrem.
        const workerCount = Math.min(5, pendingList.length || 1);
        const worker = async (primary: boolean) => {
          let sess: typeof session | null = null;
          if (primary) {
            sess = session;
          } else {
            try {
              sess = await pop3.openSyncSession();
              liveSessions.add(sess);
            } catch {
              // servidor recusou a conexão extra — segue com menos paralelismo
              return;
            }
          }
          let myNums = primary
            ? numByUidl
            : new Map((await sess.uidl()).map((e) => [e.uidl, e.num]));
          let consecFail = 0;
          let myReconnects = 0;
          try {
            for (;;) {
              if (this.isCancelled(recvTask)) return;
              const item = pendingList[cursor++];
              if (!item) return;
              const { uidl } = item;
              const num = myNums.get(uidl);
              if (!num) { done++; emitProgress(uidl, env0, true); continue; }
              let failed = false;
              let delivered = false;
              let env = { uid: num, from: '', to: '', subject: '', messageId: '' };
              try {
                // TOP é opcional no POP3 — se o servidor recusar numa mensagem,
                // baixa mesmo assim: regras só não casam e cai no destino padrão
                if (needEnvelope) {
                  try {
                    env = await sess.envelope(num);
                  } catch {
                    // segue sem envelope
                  }
                }
                // UIDL desconhecido mas Message-ID já entregue (servidor gera
                // UIDL novo por sessão, ou o mapa "vistos" se perdeu): conta
                // como vista sem baixar o corpo — impede a re-descarga total.
                if (env.messageId && seenMids.has(env.messageId.toLowerCase())) {
                  seen[uidl] = Date.now();
                  if (++sinceFlush >= 100) { flushState(); sinceFlush = 0; }
                  done++;
                  emitProgress(uidl, env, true);
                  continue;
                }
                result.scanned++;
                delivered = await deliverMessage(
                  () => sess!.retr(num),
                  env,
                  num,
                  popInboxDir,
                  true, // POP3: destino obrigatório — falha = tenta de novo no próximo sync
                );
                // Só marca como vista e só remove do servidor quando a
                // mensagem foi REALMENTE gravada — senão se perde pra sempre.
                if (delivered) {
                  seen[uidl] = Date.now();
                  if (env.messageId) seenMids.add(env.messageId.toLowerCase());
                  if (++sinceFlush >= 100) { flushState(); sinceFlush = 0; }
                  // "Deixar uma cópia no servidor" desmarcado → DELE após baixar
                  if (account.leaveOnServer === false) {
                    await sess.dele(num).catch(() => {});
                  }
                } else {
                  undelivered++;
                  failed = true;
                }
              } catch (e) {
                log.error(`POP3 mensagem ${num}`, e);
                // mensagem ilegível — pula (não marca: tenta de novo no próximo sync)
                failed = true;
              }
              if (failed) {
                consecFail++;
                // Várias falhas seguidas = a sessão provavelmente caiu →
                // reabre a PRÓPRIA sessão (e remapeia números) e continua
                if (consecFail >= 3 && myReconnects < 3 && !this.isCancelled(recvTask)) {
                  try {
                    liveSessions.delete(sess);
                    try {
                      await sess.end();
                    } catch {
                      // sessão já morta
                    }
                    sess = await pop3.openSyncSession();
                    liveSessions.add(sess);
                    myNums = new Map((await sess.uidl()).map((e) => [e.uidl, e.num]));
                    myReconnects++;
                    consecFail = 0;
                  } catch {
                    return; // não reconectou — os outros workers cobrem a fila
                  }
                }
              } else {
                consecFail = 0;
              }
              done++;
              emitProgress(uidl, env, delivered);
            }
          } finally {
            // Sessão extra fecha aqui; a primária fecha no finally externo
            // (ainda é usada pela limpeza "remover após N dias" abaixo).
            if (!primary && sess) {
              liveSessions.delete(sess);
              await sess.end().catch(() => {});
            }
          }
        };
        const env0 = { from: '', subject: '' };
        const workers = [worker(true)];
        for (let i = 1; i < workerCount; i++) workers.push(worker(false));
        await Promise.all(workers);
        // "Remover do servidor após N dias"
        if (account.leaveOnServer !== false && (account.leaveDays ?? 0) > 0) {
          const cutoff = Date.now() - (account.leaveDays || 0) * 86400000;
          for (const { uidl } of list) {
            const n = numByUidl.get(uidl);
            if (!n) continue;
            const ts = seen[uidl];
            if (ts && ts < cutoff) {
              try {
                await session.dele(n);
                delete seen[uidl];
              } catch {
                // mantém
              }
            }
          }
        }
      } finally {
        await session.end();
      }
      // Mensagens que não puderam ser entregues ficam no servidor — mas o
      // usuário precisa saber (senão parecem "perdidas").
      if (undelivered > 0) {
        result.error = popInboxDir()
          ? `${undelivered} mensagem(ns) não puderam ser gravadas na base de entrega — serão tentadas de novo no próximo envio/recebimento.`
          : `${undelivered} mensagem(ns) sem base de entrega — vincule um arquivo de dados à conta em Configurações → Email.`;
      }
      this.emit({
        taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv',
        status: this.isCancelled(recvTask) ? 'cancelled' : 'done',
      });
      flushState();
      await flushDedup();
      this.currentTask = null;
      return result;
    }

    // ---- Caminho IMAP: último UID processado ----
    const state = loadImapState();
    state[accountId] = state[accountId] || {};

    // Regras aplicadas a mensagens que chegam na caixa de entrada
    const folderId = 'INBOX';
    const lastUid = state[accountId][folderId] || 0;

    // Cancelamento derruba a conexão IMAP → downloadRaw em voo rejeita na hora
    abortInFlight = () => {
      try {
        imap.abortNow?.();
      } catch {
        // já fechada
      }
    };
    const envelopes = await imap.fetchNewEnvelopes(folderId, lastUid);
    log.info('imap', `${email}: último UID=${lastUid} — ${envelopes.length} envelope(s) novos`);
    let maxUid = lastUid;
    let done = 0;
    this.emit({ taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing', total: envelopes.length, done });

    // IMAP nunca tem destino padrão em disco — a mensagem mora no servidor.
    // Regras do usuário continuam podendo copiar para a base que ele escolher.
    const imapInboxDir = () => null;

    // Mensagens cujo download/entrega falhou na sincronização anterior —
    // maxUid já passou por elas, então são tentadas de novo aqui.
    type Pending = { uid: number; from: string; to: string; subject: string; tries?: number };
    const retryKey = `${folderId}:retry`;
    const pending = ((state[accountId][retryKey] as unknown as Pending[]) || []).filter(
      (p) => p && typeof p.uid === 'number',
    );
    const stillPending: Pending[] = [];
    let processed = 0;
    for (const p of pending) {
      if (this.isCancelled(recvTask)) break;
      const ok = await deliverMessage(
        () => imap.downloadRaw(folderId, p.uid),
        { from: p.from, to: p.to, subject: p.subject },
        p.uid,
        imapInboxDir,
      );
      if (!ok && (p.tries ?? 0) < 5) stillPending.push({ ...p, tries: (p.tries ?? 0) + 1 });
    }
    result.scanned = envelopes.length + pending.length;

    let lastEmit = 0;
    for (const env of envelopes) {
      if (this.isCancelled(recvTask)) break;
      if (env.uid > maxUid) maxUid = env.uid;
      const ok = await deliverMessage(
        () => imap.downloadRaw(folderId, env.uid),
        env,
        env.uid,
        imapInboxDir,
        // IMAP nunca exige destino local — sem regra casada a mensagem fica
        // no servidor e é dada como processada (não fica pendurada).
        false,
      );
      if (!ok) stillPending.push({ uid: env.uid, from: env.from, to: env.to, subject: env.subject });
      done++;
      // Throttle — um evento por mensagem inundava o renderer de IPCs.
      const now = Date.now();
      if (now - lastEmit >= 150 || done === envelopes.length) {
        lastEmit = now;
        this.emit({
          taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing',
          total: envelopes.length, done,
          item: { key: String(env.uid), subject: env.subject, from: env.from, status: ok ? 'ok' : 'fail', time: new Date().toISOString() },
        });
      }
    }

    this.emit({
      taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv',
      status: this.isCancelled(recvTask) ? 'cancelled' : 'done',
    });
    state[accountId][folderId] = maxUid;
    (state[accountId] as unknown as Record<string, unknown>)[retryKey] = stillPending;
    saveImapState(state);
    await flushDedup();
    this.currentTask = null;
    return result;
  }

  syncAll(): Promise<SyncResult[]> {
    return this.serialize(() => this.syncAllNow());
  }

  private async syncAllNow(): Promise<SyncResult[]> {
    this.cancelFlag = false;
    this.cancelledTasks.clear();
    const sources = sourceManager.listAll().filter((s) => s.kind === 'imap');
    // Como o Outlook: todas as tarefas aparecem de uma vez no diálogo
    // ("Aguardando") — antes a conta só aparecia quando a vez dela chegava,
    // dando a impressão de itens pipocando na lista.
    for (const s of sources) {
      const accountId = s.id.slice('acct:'.length);
      const email = accountStore.getFull(accountId)?.email || accountId;
      this.emit({
        taskId: `${accountId}:recv`, label: `${email} - Recebendo`,
        accountId, phase: 'recv', status: 'pending',
      });
    }
    const results: SyncResult[] = [];
    for (const s of sources) {
      const accountId = s.id.slice('acct:'.length);
      const email2 = accountStore.getFull(accountId)?.email || accountId;
      const recvT = `${accountId}:recv`;
      // Conta não processada — cancelamento geral ou cancelamento individual
      // ANTES dela começar: marca "Cancelado" em vez de ficar "Aguardando"
      // para sempre no diálogo (parecia que o botão não funcionava).
      if (this.cancelFlag || this.cancelledTasks.has(recvT)) {
        this.emit({
          taskId: recvT, label: `${email2} - Recebendo`,
          accountId, phase: 'recv', status: 'cancelled',
        });
        continue;
      }
      try {
        results.push(await this.syncAccountNow(accountId, false));
      } catch (e) {
        log.error(`sync conta ${accountId}`, e);
        const acct = accountStore.getFull(accountId);
        const label = acct?.label || acct?.email || accountId;
        this.emit({
          taskId: `${accountId}:recv`, label: `${label} - Recebendo`, accountId,
          phase: 'recv', status: 'error', error: (e as Error).message,
        });
        results.push({ accountId, scanned: 0, archived: 0, error: (e as Error).message });
      }
    }
    return results;
  }
}

export const ruleEngine = new RuleEngine();
