import { ImapFlow, type FetchMessageObject, type ListTreeResponse, type MessageStructureObject } from 'imapflow';
import { simpleParser } from 'mailparser';
import type { MailBackend } from './mailBackend';
import { normSearch } from '../shared/textNorm';
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
        children: treeToFolderNode(f, id),
      };
    });
}

export class ImapMailSource implements MailBackend {
  private client: ImapFlow;
  private connected = false;

  constructor(private account: AccountConfig, password: string) {
    this.client = new ImapFlow({
      host: account.imapHost,
      port: account.imapPort,
      secure: account.imapSecure,
      auth: { user: account.imapUser || account.email, pass: password },
      logger: false,
      disableAutoIdle: true,
    });
  }

  async connect(): Promise<void> {
    if (this.connected) return;
    await this.client.connect();
    this.connected = true;
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
