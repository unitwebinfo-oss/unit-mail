import { ImapFlow, type FetchMessageObject, type ListTreeResponse, type MessageStructureObject } from 'imapflow';
import { simpleParser } from 'mailparser';
import type { MailBackend } from './mailBackend';
import { normSearch } from '../shared/textNorm';
import { folderKindOf } from '../shared/folderNames';
import type {
  AccountConfig,
  AttachmentData,
  FolderNode,
  ItemDetail,
  ItemSummary,
  SearchResult,
} from '../shared/types';

function addrList(list?: Array<{ name?: string; address?: string }>): string {
  if (!list) return '';
  return list
    .map((a) => (a.name ? `${a.name} <${a.address || ''}>` : a.address || ''))
    .filter(Boolean)
    .join('; ');
}

function structureHasAttachment(node: MessageStructureObject | undefined): boolean {
  if (!node) return false;
  const isAttachmentPart =
    node.disposition === 'attachment' ||
    (node.dispositionParameters && !!node.dispositionParameters.filename) ||
    (node.parameters && !!node.parameters.name && !node.type?.startsWith('text/'));
  if (isAttachmentPart && node.part) return true;
  if (node.childNodes) return node.childNodes.some(structureHasAttachment);
  return false;
}

function treeToFolderNode(node: ListTreeResponse, idPrefix = ''): FolderNode[] {
  const folders = node.folders || [];
  return folders
    .filter((f) => !f.disabled)
    .map((f) => {
      const id = f.path || `${idPrefix}/${f.name}`;
      return {
        id,
        name: f.name || id,
        containerClass: 'IPF.Note',
        contentCount: f.status?.messages ?? 0,
        unreadCount: f.status?.unseen ?? 0,
        specialUse: f.specialUse || undefined,
        children: treeToFolderNode(f, id),
      };
    });
}

function findTreeNode(node: ListTreeResponse, path: string): ListTreeResponse | null {
  return findTreeBy(node, (f) => f.path === path);
}

/** Separador de pastas do servidor — nem todo servidor o informa, então
 *  deduz pelos caminhos reais (ex.: INBOX.Sent filho de INBOX ⇒ '.'). */
function inferDelim(nodes: ListTreeResponse[]): string | null {
  for (const n of nodes) {
    for (const c of n.folders || []) {
      if (n.path && c.path?.startsWith(n.path)) {
        const d = c.path.slice(n.path.length, n.path.length + 1);
        if (d) return d;
      }
    }
    const r = inferDelim(n.folders || []);
    if (r) return r;
  }
  return null;
}

function findTreeBy(
  node: ListTreeResponse,
  pred: (f: ListTreeResponse) => boolean,
): ListTreeResponse | null {
  if (node.path && pred(node)) return node;
  for (const child of node.folders || []) {
    const found = findTreeBy(child, pred);
    if (found) return found;
  }
  return null;
}

/** Flag SPECIAL-USE IMAP ↔ tipo semântico de pasta do app. */
const SPECIAL_USE_FLAGS: Record<string, string> = {
  sent: '\\sent',
  drafts: '\\drafts',
  trash: '\\trash',
  junk: '\\junk',
  archive: '\\archive',
};

export class ImapMailSource implements MailBackend {
  /** Cache das pastas especiais — listTree por mensagem seria um roundtrip
   *  a mais por item em exclusões/movimentações em lote. Invalidado quando
   *  a estrutura de pastas muda (createFolder). */
  private specialCache = new Map<string, string | null>();
  private client: ImapFlow;
  private connected = false;
  private readonly password: string;
  private connecting: Promise<void> | null = null;

  constructor(private account: AccountConfig, password: string) {
    this.password = password;
    this.client = this.newClient();
  }

  private newClient(): ImapFlow {
    const client = new ImapFlow({
      host: this.account.imapHost,
      port: this.account.imapPort,
      secure: this.account.imapSecure,
      auth: { user: this.account.imapUser || this.account.email, pass: this.password },
      logger: false,
      disableAutoIdle: true,
      // Servidores próprios/hospedagens com certificado autoassinado
      tls: this.account.allowSelfSigned ? { rejectUnauthorized: false } : undefined,
    });
    // Se a conexão cair por qualquer motivo, esta instância não pode ser
    // reusada — a próxima operação cria outra.
    client.on('close', () => {
      this.connected = false;
    });
    client.on('error', () => {
      this.connected = false;
    });
    return client;
  }

  async connect(): Promise<void> {
    if (this.connected) return;
    if (!this.connecting) {
      this.connecting = (async () => {
        try {
          await this.client.connect();
          this.connected = true;
        } finally {
          this.connecting = null;
        }
      })();
    }
    try {
      await this.connecting;
    } catch (e) {
      if ((e as { code?: string })?.code === 'InstanceReused') {
        this.client = this.newClient();
        await this.client.connect();
        this.connected = true;
        return;
      }
      throw e;
    }
  }

  async getFolderTree(): Promise<FolderNode[]> {
    await this.connect();
    const tree = await this.client.listTree({
      statusQuery: { messages: true, unseen: true },
    } as never);
    return treeToFolderNode(tree as ListTreeResponse);
  }

  async getItemCount(folderId: string): Promise<number> {
    await this.connect();
    const status = await this.client.status(folderId, { messages: true });
    return status ? status.messages || 0 : 0;
  }

  async getItems(folderId: string, offset: number, limit: number): Promise<ItemSummary[]> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      const exists = this.client.mailbox && typeof this.client.mailbox === 'object' ? (this.client.mailbox as { exists: number }).exists : 0;
      if (!exists) return [];
      const end = exists - offset;
      const start = Math.max(1, end - limit + 1);
      if (end < 1) return [];
      const items: ItemSummary[] = [];
      for await (const msg of this.client.fetch(
        `${start}:${end}`,
        { envelope: true, flags: true, uid: true, bodyStructure: true, size: true },
      )) {
        items.push(this.summarize(msg));
      }
      return items.reverse();
    } finally {
      lock.release();
    }
  }

  /** Marca lida/não-lida NO SERVIDOR (flag \Seen) — como todo cliente IMAP. */
  async setItemRead(folderId: string, uid: number, read: boolean): Promise<void> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      if (read) await this.client.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
      else await this.client.messageFlagsRemove(String(uid), ['\\Seen'], { uid: true });
    } finally {
      lock.release();
    }
  }

  /** Marca TODAS as não lidas da pasta como lidas no servidor. */
  async markAllRead(folderId: string): Promise<void> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      const uids = await this.client.search({ seen: false }, { uid: true });
      if (uids && uids.length > 0) {
        await this.client.messageFlagsAdd(uids.join(','), ['\\Seen'], { uid: true });
      }
    } finally {
      lock.release();
    }
  }

  async getItem(folderId: string, uid: number): Promise<ItemDetail> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      const dl = await this.client.download(String(uid), undefined, { uid: true });
      if (!dl.content) throw new Error('Mensagem não encontrada');
      const chunks: Buffer[] = [];
      for await (const chunk of dl.content) chunks.push(chunk as Buffer);
      const raw = Buffer.concat(chunks);
      const parsed = await simpleParser(raw);

      try {
        await this.client.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true });
      } catch {
        // falha ao marcar como lida não é crítica
      }

      const attachments = (parsed.attachments || []).map((a, i) => ({
        index: i,
        filename: a.filename || `anexo-${i}`,
        size: a.size || a.content?.length || 0,
        mimeTag: a.contentType || '',
        contentId: a.contentId || '',
        isEmbeddedMessage: false,
        isInline: a.contentDisposition === 'inline' && !!a.contentId,
      }));

      let bodyHTML = parsed.html || '';
      if (bodyHTML) {
        for (const a of parsed.attachments || []) {
          if (a.contentId && a.content) {
            const dataUri = `data:${a.contentType || 'application/octet-stream'};base64,${a.content.toString('base64')}`;
            bodyHTML = bodyHTML
              .split(`cid:${a.contentId}`)
              .join(dataUri)
              .split(`cid:${a.contentId.replace(/[<>]/g, '')}`)
              .join(dataUri);
          }
        }
      }

      const fromAddr = parsed.from?.value?.[0];
      const toAddr = (parsed.to as { value?: Array<{ name?: string; address?: string }> } | undefined)?.value;
      const ccAddr = (parsed.cc as { value?: Array<{ name?: string; address?: string }> } | undefined)?.value;

      return {
        index: uid,
        messageClass: 'IPM.Note',
        subject: parsed.subject || '(sem assunto)',
        senderName: fromAddr?.name || fromAddr?.address || '',
        senderEmailAddress: fromAddr?.address || '',
        displayTo: addrList(toAddr),
        displayCC: addrList(ccAddr),
        sentTime: parsed.date ? parsed.date.toISOString() : null,
        deliveryTime: parsed.date ? parsed.date.toISOString() : null,
        hasAttachments: attachments.filter((a) => !a.isInline).length > 0,
        isRead: true,
        importance: 0,
        preview: (parsed.text || '').slice(0, 200),
        body: parsed.text || '',
        bodyHTML,
        transportMessageHeaders: parsed.headerLines?.map((h) => `${h.key}: ${h.line}`).join('\n') || '',
        attachments: attachments.filter((a) => !a.isInline),
      };
    } finally {
      lock.release();
    }
  }

  /** Importa uma mensagem RFC822 na pasta IMAP, preservando o estado lido.
   *  Lança erro se o servidor não confirmou o APPEND — sem isso um "mover"
   *  excluiria a origem de uma mensagem que nunca chegou ao destino. */
  async appendItem(folderId: string, raw: Buffer, isRead = false): Promise<void> {
    await this.connect();
    const flags = isRead ? ['\\Seen'] : [];
    const res = await this.client.append(folderId, raw, flags, new Date());
    if (!res) throw new Error(`O servidor recusou a mensagem em "${folderId}"`);
  }

  /**
   * Caminho da pasta especial do servidor (Itens Enviados, Rascunhos,
   * Lixeira…) — primeiro pela flag SPECIAL-USE (imapflow já resolve nomes
   * localizados), depois pelo nome conhecido da pasta.
   */
  async specialFolderPath(kind: 'sent' | 'drafts' | 'trash' | 'junk' | 'archive'): Promise<string | null> {
    if (this.specialCache.has(kind)) return this.specialCache.get(kind) ?? null;
    await this.connect();
    const tree = (await this.client.listTree()) as ListTreeResponse;
    const flag = SPECIAL_USE_FLAGS[kind];
    const hit =
      (flag &&
        findTreeBy(
          tree,
          (f) => typeof f.specialUse === 'string' && f.specialUse.toLowerCase() === flag,
        )) ||
      findTreeBy(tree, (f) => folderKindOf(f.name || '') === kind);
    const path = hit?.path || null;
    this.specialCache.set(kind, path);
    return path;
  }

  /** Exclui de verdade (\\Deleted + expurgo) — usada internamente depois de
   *  copiar a mensagem para outra pasta num "mover". */
  async expungeItem(folderId: string, uid: number): Promise<void> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      await this.client.messageDelete(String(uid), { uid: true });
    } finally {
      lock.release();
    }
  }

  /**
   * Move a mensagem para outra pasta DA MESMA conta — usa a extensão MOVE
   * quando o servidor tem, senão COPY + expurgo (messageMove devolve false
   * em servidores sem MOVE, e sem o fallback a mensagem ficava na origem).
   */
  async moveWithin(srcFolderId: string, uid: number, dstFolderId: string): Promise<void> {
    await this.connect();
    const lock = await this.client.getMailboxLock(srcFolderId);
    try {
      await this.moveWithinUnlocked(srcFolderId, uid, dstFolderId);
    } finally {
      lock.release();
    }
  }

  /** Excluir pelo usuário — como o Outlook: a mensagem vai para Itens
   *  Excluídos do servidor quando a pasta existe; senão expurga. */
  async deleteItem(folderId: string, uid: number): Promise<void> {
    await this.connect();
    const trash = await this.specialFolderPath('trash').catch(() => null);
    const lock = await this.client.getMailboxLock(folderId);
    try {
      if (trash && trash !== folderId) {
        await this.moveWithinUnlocked(folderId, uid, trash);
      } else {
        await this.client.messageDelete(String(uid), { uid: true });
      }
    } finally {
      lock.release();
    }
  }

  /** Move dentro do lock já adquirido pelo chamador. */
  private async moveWithinUnlocked(srcFolderId: string, uid: number, dstFolderId: string): Promise<void> {
    const moved = await this.client.messageMove(String(uid), dstFolderId, { uid: true });
    if (!moved) {
      // Sem MOVE: COPY primeiro e só expurga se o servidor confirmou a cópia.
      const copied = await this.client.messageCopy(String(uid), dstFolderId, { uid: true });
      if (!copied) throw new Error(`O servidor não copiou a mensagem para "${dstFolderId}"`);
      await this.client.messageDelete(String(uid), { uid: true });
    }
  }

  /** Cria uma pasta no servidor e retorna a árvore atualizada. */
  async createFolder(name: string, parentFolderId: string | null): Promise<FolderNode[]> {
    await this.connect();
    const tree = (await this.client.listTree()) as ListTreeResponse;
    const parent = parentFolderId ? findTreeNode(tree, parentFolderId) : null;
    const delimiter =
      parent?.delimiter ||
      inferDelim(tree.folders || []) ||
      (tree.folders || []).find((f) => f.delimiter)?.delimiter ||
      '/';
    const clean = name.trim().replace(/[\\/]/g, delimiter);
    if (!clean) throw new Error('O nome da pasta não pode ficar vazio');
    const path = parentFolderId ? `${parentFolderId}${delimiter}${clean}` : clean;
    await this.client.mailboxCreate(path);
    this.specialCache.clear();
    return treeToFolderNode((await this.client.listTree({
      statusQuery: { messages: true, unseen: true },
    }) as ListTreeResponse));
  }

  /** Guarda: Caixa de Entrada e pastas especiais (\Sent, \Trash…) não
   *  podem ser renomeadas nem excluídas — protege o próprio servidor. */
  private async assertMutableFolder(folderId: string, action: string): Promise<void> {
    if (folderId.toUpperCase() === 'INBOX') {
      throw new Error(`A Caixa de Entrada não pode ser ${action}`);
    }
    const tree = (await this.client.listTree()) as ListTreeResponse;
    const node = findTreeNode(tree, folderId);
    if (!node) throw new Error('Pasta não encontrada no servidor');
    if (node.specialUse) {
      throw new Error(`"${node.name}" é uma pasta especial do servidor e não pode ser ${action}`);
    }
  }

  /** Renomeia a pasta no servidor (mantém o pai: INBOX.X → INBOX.Novo). */
  async renameFolder(folderId: string, newName: string): Promise<FolderNode[]> {
    await this.connect();
    await this.assertMutableFolder(folderId, 'renomeada');
    const clean = newName.trim();
    if (!clean) throw new Error('O nome da pasta não pode ficar vazio');
    // Separador pelo próprio caminho (INBOX.X → '.', INBOX/X → '/').
    const delim = folderId.includes('/') ? '/' : folderId.includes('.') ? '.' : '/';
    const sep = folderId.lastIndexOf(delim);
    const parentPath = sep > 0 ? folderId.slice(0, sep) : '';
    const leaf = clean.replace(/[\\/]/g, delim);
    const newPath = parentPath ? `${parentPath}${delim}${leaf}` : leaf;
    await this.client.mailboxRename(folderId, newPath);
    this.specialCache.clear();
    return treeToFolderNode((await this.client.listTree({
      statusQuery: { messages: true, unseen: true },
    }) as ListTreeResponse));
  }

  /** Exclui a pasta no servidor com todo o conteúdo. */
  async deleteFolder(folderId: string): Promise<FolderNode[]> {
    await this.connect();
    await this.assertMutableFolder(folderId, 'excluída');
    await this.client.mailboxDelete(folderId);
    this.specialCache.clear();
    return treeToFolderNode((await this.client.listTree({
      statusQuery: { messages: true, unseen: true },
    }) as ListTreeResponse));
  }

  async getAttachmentData(folderId: string, uid: number, attachIndex: number): Promise<AttachmentData> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      const dl = await this.client.download(String(uid), undefined, { uid: true });
      if (!dl.content) throw new Error('Mensagem não encontrada');
      const chunks: Buffer[] = [];
      for await (const chunk of dl.content) chunks.push(chunk as Buffer);
      const raw = Buffer.concat(chunks);
      const parsed = await simpleParser(raw);
      const att = (parsed.attachments || []).filter((a) => a.contentDisposition !== 'inline')[attachIndex];
      if (!att) throw new Error('Anexo não encontrado');
      return {
        filename: att.filename || `anexo-${attachIndex}`,
        mimeTag: att.contentType || '',
        dataBase64: (att.content || Buffer.alloc(0)).toString('base64'),
      };
    } finally {
      lock.release();
    }
  }

  async search(
    query: string,
    _containerClass?: string,
    _maxResults?: number,
    _onBatch?: (hits: SearchResult[]) => void,
    folderId?: string,
  ): Promise<SearchResult[]> {
    void _containerClass;
    void _maxResults;
    void _onBatch;
    await this.connect();
    // Multi-token sem acentos — mesmo critério da busca dos arquivos de dados
    const tokens = normSearch(query).split(/\s+/).filter(Boolean);
    const results: SearchResult[] = [];
    const tree = await this.getFolderTree();
    const important: FolderNode[] = [];
    const collect = (nodes: FolderNode[]) => {
      for (const n of nodes) {
        important.push(n);
        collect(n.children);
      }
    };
    collect(tree);
    // Escopo "pasta atual" — busca só na pasta selecionada
    const scope = folderId ? important.filter((n) => n.id === folderId) : important;
    for (const folder of scope) {
      if (results.length >= 1000) break;
      try {
        const items = await this.getItems(folder.id, 0, 2000);
        for (const it of items) {
          if (results.length >= 1000) break;
          const text = normSearch(
            `${it.subject}\n${it.senderName}\n${it.senderEmailAddress}\n${it.displayTo}\n${it.preview || ''}`,
          );
          if (tokens.every((t) => text.includes(t))) {
            results.push({ folderId: folder.id, folderName: folder.name, item: it });
          }
        }
      } catch {
        // pasta inacessível — ignora
      }
    }
    return results;
  }

  /**
   * Lista envelopes novos na pasta (UID > sinceUid) — usado pelo motor de regras.
   * Retorna apenas metadados leves (envelope), sem baixar o corpo.
   */
  async fetchNewEnvelopes(
    folderId: string,
    sinceUid: number,
  ): Promise<Array<{ uid: number; from: string; to: string; subject: string }>> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      const out: Array<{ uid: number; from: string; to: string; subject: string }> = [];
      for await (const msg of this.client.fetch(`${sinceUid + 1}:*`, {
        envelope: true,
        uid: true,
      }, { uid: true })) {
        if (!msg || !msg.uid || msg.uid <= sinceUid) continue;
        const env = msg.envelope;
        out.push({
          uid: msg.uid,
          from: addrList(env?.from),
          to: addrList(env?.to),
          subject: env?.subject || '',
        });
      }
      return out;
    } finally {
      lock.release();
    }
  }

  /** Baixa a mensagem RFC822 completa como Buffer — usado pelo arquivamento. */
  async downloadRaw(folderId: string, uid: number): Promise<Buffer> {
    await this.connect();
    const lock = await this.client.getMailboxLock(folderId);
    try {
      const dl = await this.client.download(String(uid), undefined, { uid: true });
      if (!dl.content) throw new Error('Mensagem não encontrada');
      const chunks: Buffer[] = [];
      for await (const chunk of dl.content) chunks.push(chunk as Buffer);
      return Buffer.concat(chunks);
    } finally {
      lock.release();
    }
  }

  /**
   * Derruba a conexão para abortar um download em andamento (botão
   * "Cancelar" do Enviar/Receber). A conexão é reaberta sob demanda na
   * próxima operação (`connect()` verifica `connected`).
   */
  abortNow(): void {
    try {
      void this.client.close();
    } catch {
      // já fechada
    }
    this.connected = false;
  }

  async close(): Promise<void> {
    try {
      if (this.connected) await this.client.logout();
    } catch {
      // ignora erro de logout
    }
    this.connected = false;
  }

  private summarize(msg: FetchMessageObject): ItemSummary {
    const env = msg.envelope;
    const flags = msg.flags || new Set<string>();
    return {
      index: msg.uid,
      messageClass: 'IPM.Note',
      subject: env?.subject || '(sem assunto)',
      senderName: env?.from?.[0]?.name || env?.from?.[0]?.address || '',
      senderEmailAddress: env?.from?.[0]?.address || '',
      displayTo: addrList(env?.to),
      displayCC: addrList(env?.cc),
      sentTime: env?.date ? new Date(env.date).toISOString() : null,
      deliveryTime: env?.date ? new Date(env.date).toISOString() : null,
      hasAttachments: structureHasAttachment(msg.bodyStructure),
      isRead: flags.has('\\Seen'),
      importance: 0,
      preview: '',
    };
  }
}
