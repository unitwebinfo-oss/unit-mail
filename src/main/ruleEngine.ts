import { createHash, randomUUID } from 'crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { mkdir, open, readdir, readFile, writeFile } from 'fs/promises';
import { dirname, join } from 'path';
import { sourceManager } from './sourceManager';
import { accountStore } from './accountStore';
import { archiveRootForPst } from './paths';
import { isLocalStorePath } from './localStore';
import { dataDir } from './dataDir';
import { log } from './logger';
import { recordRecipients } from './contactHistory';
import { scanEmailBuffer } from './antivirus';
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

/** Pasta de quarentena dentro da base de entrega (como a "Quarentena" do AV). */
const QUARANTINE_FOLDER = 'Itens infectados';

/** Configurações → Geral → "Verificar vírus" (padrão: ligado). */
function virusScanEnabled(): boolean {
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
  /** nome do .eml → chave (Message-ID/sha1) */
  files: Map<string, string>;
  /** chave → nome do .eml (para validar arquivo que sumiu) */
  byKey: Map<string, string>;
  dirty: boolean;
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
 * Chave estável de dedupe: o Message-ID do cabeçalho, quando presente.
 * Mensagens migradas de PST têm bytes diferentes do raw POP3/IMAP, então o
 * hash do arquivo inteiro não basta — o Message-ID é o mesmo nas duas cópias.
 * Fallback: hash do conteúdo (mensagens sem Message-ID).
 */
function keyOf(buf: Buffer): string {
  const head = buf.subarray(0, Math.min(buf.length, 16384)).toString('latin1');
  const m = /^message-id:[ \t]*<?([^>\r\n]+)>?/im.exec(head);
  return m ? `mid:${m[1].trim().toLowerCase()}` : `sha1:${hashOf(buf)}`;
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

/** Lê só o cabeçalho do .eml para extrair o Message-ID; sem ele, hash do arquivo inteiro. */
async function fileKey(fp: string): Promise<string | null> {
  try {
    const fh = await open(fp, 'r');
    try {
      const head = Buffer.alloc(16384);
      const { bytesRead } = await fh.read(head, 0, 16384, 0);
      const m = /^message-id:[ \t]*<?([^>\r\n]+)>?/im.exec(
        head.subarray(0, bytesRead).toString('latin1'),
      );
      if (m) return `mid:${m[1].trim().toLowerCase()}`;
    } finally {
      await fh.close();
    }
    return `sha1:${hashOf(await readFile(fp))}`;
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
      const files = new Map<string, string>();
      let stored: Record<string, string> = {};
      try {
        stored =
          (JSON.parse(await readFile(join(dir, DEDUP_FILE), 'utf8')) as { files?: Record<string, string> })
            .files || {};
      } catch {
        // sem índice ainda — será montado agora e persistido
      }
      let names: string[] = [];
      try {
        names = (await readdir(dir)).filter((f) => f.endsWith('.eml'));
      } catch {
        // diretório ainda não existe — vazio
      }
      // Reconcilia: índice só vale para arquivo que ainda existe no disco
      for (const f of names) if (stored[f]) files.set(f, stored[f]);
      // Arquivos sem entrada no índice (importados, movidos manualmente, ou
      // pasta de uma instalação antiga sem _dedup.json) → indexados em
      // BACKGROUND. A entrega não pode esperar a varredura de dezenas de
      // milhares de arquivos — era o "travado na primeira mensagem" numa
      // base restaurada. Durante a janela de montagem, o mapa de vistos +
      // Message-IDs em memória cobre a deduplicação.
      const missing = names.filter((f) => !files.has(f));
      const idx: DedupIdx = {
        files,
        byKey: new Map([...files].map(([f, k]) => [k, f])),
        dirty: missing.length > 0 || Object.keys(stored).length !== files.size,
      };
      if (missing.length > 0) {
        const build = async () => {
          for (let i = 0; i < missing.length; i += 64) {
            const slice = missing.slice(i, i + 64);
            const batch = await Promise.all(slice.map((f) => fileKey(join(dir, f))));
            slice.forEach((f, j) => {
              const k = batch[j];
              if (k) {
                files.set(f, k);
                idx.byKey.set(k, f);
              }
            });
          }
          scheduleDedupSave(dir, idx);
        };
        void build();
      }
      return idx;
    })();
    dedupCache.set(dir, p);
  }
  return p;
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
 * Grava o .eml em `dir` só se ainda não existir uma cópia idêntica lá.
 * Retorna true quando a mensagem está na pasta (já estava ou gravou agora),
 * false quando a gravação falhou.
 */
async function writeEmlDedup(dir: string, fileName: string, buf: Buffer): Promise<boolean> {
  try {
    const idx = await dirKeys(dir);
    const k = keyOf(buf);
    if (dedupHas(idx, dir, k)) return true; // já entregue — conta como sucesso
    await mkdir(dir, { recursive: true });
    await writeFile(join(dir, fileName), buf);
    idx.files.set(fileName, k);
    idx.byKey.set(k, fileName);
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

/**
 * Base de entrega de uma conta POP3: a escolhida pelo usuário ou, se nenhuma,
 * uma base local criada automaticamente (equivalente ao PST padrão do Outlook).
 * Retorna o diretório físico onde gravar a pasta.
 */
export function ensureDeliveryStore(account: AccountConfig): string {
  if (account.deliveryPath) return account.deliveryPath;
  // Sem base vinculada à conta → usa o arquivo de dados marcado como principal
  const principal = defaultStorePath();
  if (principal) {
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
    const hasDeliveryTarget = deliveryRules.length > 0 || !!account?.deliveryPath || !!defaultStorePath();
    // IMAP sem nenhum destino/regra: nada a fazer (fica só no servidor).
    // POP3 sempre baixa (entrega padrão cria base se necessário).
    if (userRules.length === 0 && !hasDeliveryTarget && account?.protocol !== 'pop3') {
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
        const targetOk =
          t && (t.toLowerCase().endsWith('.pst') || isLocalStorePath(t));
        try {
          if (!targetOk) throw new Error('destino de regra inválido');
          const buf = await raw();
          const dir = ruleDestDir(rule);
          // Dedupe: cópia (mesmo Message-ID) já presente na pasta → entregue
          if (dedupHas(await dirKeys(dir), dir, keyOf(buf))) return 'delivered';
          if (await writeEmlDedup(dir, `${Date.now()}-${randomUUID().slice(0, 8)}.eml`, buf)) {
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
            const qdir = join(dirname(dir), QUARANTINE_FOLDER);
            const ok = await writeEmlDedup(
              qdir,
              `${Date.now()}-${uidTag}-virus-${randomUUID().slice(0, 8)}.eml`,
              raw,
            );
            if (ok) {
              result.archived++;
              result.infected = (result.infected ?? 0) + 1;
              log.info(
                'antivirus',
                `mensagem de "${env.from}" movida para "${QUARANTINE_FOLDER}"` +
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
        return p ? folderDirFor(p, account.deliveryFolder || 'Caixa de Entrada') : null;
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
      // reabre e continua a partir da mensagem atual — o UIDL novo resolve os
      // números, que podem mudar se nossas DELE já foram commitadas.
      let reconnects = 0;
      let numByUidl = new Map<string, number>();
      const reopen = async () => {
        try {
          await session.end();
        } catch {
          // sessão já morta
        }
        session = await pop3.openSyncSession();
        numByUidl = new Map((await session.uidl()).map((e) => [e.uidl, e.num]));
        reconnects++;
      };
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
        let done = 0;
        let consecFail = 0;
        let lastEmit = 0;
        this.emit({ taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing', total: pendingList.length, done });
        for (const { uidl } of pendingList) {
          if (this.isCancelled(recvTask)) break;
          const num = numByUidl.get(uidl);
          if (!num) { done++; continue; }
          let failed = false;
          let delivered = false;
          let env = { uid: num, from: '', to: '', subject: '', messageId: '' };
          try {
            // TOP é opcional no POP3 — se o servidor recusar numa mensagem,
            // baixa mesmo assim: regras só não casam e cai no destino padrão
            if (needEnvelope) {
              try {
                env = await session.envelope(num);
              } catch {
                // segue sem envelope
              }
            }
            // UIDL desconhecido mas Message-ID já entregue (servidor gera
            // UIDL novo por sessão, ou o mapa "vistos" se perdeu): conta como
            // vista sem baixar o corpo — é o que impede a re-descarga total.
            if (env.messageId && seenMids.has(env.messageId.toLowerCase())) {
              seen[uidl] = Date.now();
              if (++sinceFlush >= 100) { flushState(); sinceFlush = 0; }
              done++;
              continue;
            }
            result.scanned++;
            delivered = await deliverMessage(
              () => session.retr(num),
              env,
              num,
              popInboxDir,
              true, // POP3: destino obrigatório — falha = tenta de novo no próximo sync
            );
            // Só marca como vista e só remove do servidor quando a mensagem
            // foi REALMENTE gravada — senão ela se perderia para sempre.
            if (delivered) {
              seen[uidl] = Date.now();
              if (env.messageId) seenMids.add(env.messageId.toLowerCase());
              if (++sinceFlush >= 100) { flushState(); sinceFlush = 0; }
              // "Deixar uma cópia no servidor" desmarcado → DELE após baixar
              if (account.leaveOnServer === false) {
                await session.dele(num).catch(() => {});
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
            // Várias falhas seguidas = a sessão provavelmente caiu → reabre
            // e continua o loop (em vez de desistir do resto da caixa)
            if (consecFail >= 3 && reconnects < 3) {
              try {
                await reopen();
                consecFail = 0;
              } catch {
                break; // não reconectou — próxima sync continua de onde parou
              }
            }
          } else {
            consecFail = 0;
          }
          done++;
          // Progresso com throttle (~150 ms): emitir um evento por mensagem
          // mandava milhares de IPCs ao renderer, que re-renderizava o app
          // inteiro a cada uma — era a principal fonte da lentidão no sync.
          const now = Date.now();
          if (now - lastEmit >= 150 || done === pendingList.length) {
            lastEmit = now;
            this.emit({
              taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing',
              total: pendingList.length, done,
              item: { key: uidl, subject: env.subject, from: env.from, status: delivered ? 'ok' : 'fail', time: new Date().toISOString() },
            });
          }
        }
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
    let maxUid = lastUid;
    let done = 0;
    this.emit({ taskId: recvTask, label: `${email} - Recebendo`, accountId, phase: 'recv', status: 'processing', total: envelopes.length, done });

    const imapInboxDir = () => {
      const storePath = account?.deliveryPath || defaultStorePath();
      return storePath
        ? folderDirFor(storePath, account?.deliveryFolder || 'Caixa de Entrada')
        : null;
    };

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
