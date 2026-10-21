import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import { simpleParser } from 'mailparser';
import { Pop3Client } from './pop3Client';
import { dataDir } from './dataDir';
import { normSearch } from '../shared/textNorm';
import type { MailBackend } from './mailBackend';
import type {
  AccountConfig,
  AttachmentData,
  FolderNode,
  ItemDetail,
  ItemSummary,
  SearchResult,
} from '../shared/types';

interface HeaderMeta {
  subject: string;
  fromName: string;
  fromAddr: string;
  to: string;
  date: string | null;
  read: boolean;
}

type HeaderCache = Record<string, HeaderMeta>; // chave: uidl

import { decodeMimeWords, getHeader, parseFrom } from './mailHeaders';

/** Backend POP3 — o servidor não tem pastas nem flags; só a caixa de entrada. */
export class Pop3MailSource implements MailBackend {
  private headerCache: HeaderCache = {};
  private uidlMap = new Map<string, number>(); // uidl → num da última listagem
  private numUidl = new Map<number, string>();

  constructor(
    private account: AccountConfig,
    private password: string,
  ) {
    this.loadCache();
  }

  private cachePath(): string {
    return join(dataDir(), `pop3-cache-${this.account.id}.json`);
  }

  private loadCache(): void {
    try {
      const p = this.cachePath();
      if (existsSync(p)) this.headerCache = JSON.parse(readFileSync(p, 'utf-8'));
    } catch {
      this.headerCache = {};
    }
  }

  private saveCache(): void {
    try {
      mkdirSync(dataDir(), { recursive: true });
      writeFileSync(this.cachePath(), JSON.stringify(this.headerCache));
    } catch {
      // não crítico
    }
  }

  private async withClient<T>(fn: (c: Pop3Client) => Promise<T>): Promise<T> {
    const c = new Pop3Client();
    await c.connect(
      this.account.imapHost, // campo reutilizado como "servidor de entrada"
      this.account.imapPort,
      this.account.imapSecure,
      this.account.imapUser || this.account.email,
      this.password,
      !!this.account.allowSelfSigned,
    );
    try {
      return await fn(c);
    } finally {
      await c.quit();
    }
  }

  private async refreshUidl(c: Pop3Client): Promise<Array<{ num: number; uidl: string }>> {
    const list = await c.uidl();
    this.uidlMap.clear();
    this.numUidl.clear();
    for (const { num, uidl } of list) {
      this.uidlMap.set(uidl, num);
      this.numUidl.set(num, uidl);
    }
    return list;
  }

  private metaToSummary(num: number, meta: HeaderMeta): ItemSummary {
    return {
      index: num,
      messageClass: 'IPM.Note',
      subject: meta.subject || '(sem assunto)',
      senderName: meta.fromName,
      senderEmailAddress: meta.fromAddr,
      displayTo: meta.to,
      displayCC: '',
      sentTime: meta.date,
      deliveryTime: meta.date,
      hasAttachments: false, // só sabemos baixando a mensagem inteira
      isRead: meta.read,
      importance: 0,
      preview: '',
    };
  }

  getFolderTree(): FolderNode[] {
    return [
      {
        id: 'INBOX',
        name: 'Caixa de Entrada',
        containerClass: 'IPF.Note',
        contentCount: 0, // atualizado por getItemCount na UI
        unreadCount: 0,
        children: [],
      },
    ];
  }

  async getItemCount(_folderId: string): Promise<number> {
    void _folderId;
    return this.withClient((c) => c.stat());
  }

  async getItems(_folderId: string, offset: number, limit: number): Promise<ItemSummary[]> {
    void _folderId;
    return this.withClient(async (c) => {
      const total = await c.stat();
      if (total === 0) return [];
      await this.refreshUidl(c);
      const end = total - offset; // mais novo primeiro
      const start = Math.max(1, end - limit + 1);
      const items: ItemSummary[] = [];
      for (let num = start; num <= end; num++) {
        const uidl = this.numUidl.get(num);
        if (!uidl) continue;
        let meta = this.headerCache[uidl];
        if (!meta) {
          try {
            meta = await this.fetchMeta(c, num);
            this.headerCache[uidl] = meta;
          } catch {
            meta = { subject: '(erro ao ler)', fromName: '', fromAddr: '', to: '', date: null, read: true };
          }
        }
        items.push(this.metaToSummary(num, meta));
      }
      this.saveCache();
      return items.reverse();
    });
  }

  async getItem(_folderId: string, index: number): Promise<ItemDetail> {
    void _folderId;
    const detail = await this.withClient(async (c) => {
      const raw = await c.retr(index);
      const parsed = await simpleParser(raw);
      return this.detailFromParsed(parsed, index);
    });
    // marca como lida localmente (POP3 não tem flags no servidor)
    const uidl = this.numUidl.get(index);
    if (uidl && this.headerCache[uidl]) {
      this.headerCache[uidl].read = true;
      this.saveCache();
    }
    return detail;
  }

  async getAttachmentData(_folderId: string, index: number, attachIndex: number): Promise<AttachmentData> {
    void _folderId;
    return this.withClient(async (c) => {
      const raw = await c.retr(index);
      const parsed = await simpleParser(raw);
      const att = (parsed.attachments || []).filter(
        (a) => !(a.contentDisposition === 'inline' && a.contentId),
      )[attachIndex];
      if (!att) throw new Error('Anexo não encontrado');
      return {
        filename: att.filename || `anexo-${attachIndex}`,
        mimeTag: att.contentType || '',
        dataBase64: att.content ? att.content.toString('base64') : '',
      };
    });
  }

  /** Cabeçalhos de uma mensagem via TOP (lança erro se o servidor recusar). */
  private async fetchMeta(c: Pop3Client, num: number): Promise<HeaderMeta> {
    const raw = await c.top(num);
    const from = parseFrom(getHeader(raw, 'From'));
    const dateStr = getHeader(raw, 'Date');
    const d = dateStr ? new Date(dateStr) : null;
    return {
      subject: getHeader(raw, 'Subject'),
      fromName: from.name,
      fromAddr: from.addr,
      to: getHeader(raw, 'To'),
      date: d && !isNaN(d.getTime()) ? d.toISOString() : null,
      read: false,
    };
  }

  /**
   * Busca em TODA a caixa do servidor: cabeçalhos ainda não conhecidos são
   * baixados por TOP numa única sessão e entram no cache (buscas seguintes
   * são instantâneas). Mensagens com erro de leitura são ignoradas.
   */
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
    // POP3 só tem a caixa de entrada no servidor — outra pasta não tem o que buscar
    if (folderId && folderId !== 'INBOX') return [];
    // Multi-token sem acentos — mesmo critério da busca dos arquivos de dados
    const tokens = normSearch(query).split(/\s+/).filter(Boolean);
    const results: SearchResult[] = [];
    await this.withClient(async (c) => {
      const list = await this.refreshUidl(c);
      for (const { num, uidl } of list) {
        let meta = this.headerCache[uidl];
        if (!meta) {
          try {
            meta = await this.fetchMeta(c, num);
            this.headerCache[uidl] = meta;
          } catch {
            continue;
          }
        }
        const text = normSearch(
          `${meta.subject}\n${meta.fromName}\n${meta.fromAddr}\n${meta.to}`,
        );
        if (tokens.every((t) => text.includes(t))) {
          results.push({
            folderId: 'INBOX',
            folderName: 'Caixa de Entrada',
            item: this.metaToSummary(num, meta),
          });
        }
        if (results.length >= 2000) break;
      }
      this.saveCache();
    });
    return results;
  }

  // ---------- Extras usados pelo motor de regras ----------

  /** UIDLs presentes no servidor (para diff de novidades). */
  async listUids(): Promise<Array<{ num: number; uidl: string }>> {
    return this.withClient((c) => c.uidl());
  }

  /** Envelope leve de uma mensagem, para avaliação de regra. */
  async fetchEnvelope(num: number): Promise<{ uid: number; from: string; to: string; subject: string }> {
    return this.withClient(async (c) => {
      const raw = await c.top(num);
      return {
        uid: num,
        from: getHeader(raw, 'From'),
        to: getHeader(raw, 'To'),
        subject: decodeMimeWords(getHeader(raw, 'Subject')),
      };
    });
  }

  /** Mensagem RFC822 completa, para arquivamento por regra. */
  async downloadRaw(_folderId: string, num: number): Promise<Buffer> {
    void _folderId;
    return this.withClient((c) => c.retr(num));
  }

  /**
   * Sessão de sincronização: UMA conexão POP3 para todo o ciclo
   * Enviar/Receber (UIDL + TOP + RETR + DELE). Abrir uma conexão por
   * mensagem fazia o servidor cortar/rate-limitar e o sync travava.
   */
  async openSyncSession(): Promise<{
    uidl(): Promise<Array<{ num: number; uidl: string }>>;
    envelope(num: number): Promise<{ uid: number; from: string; to: string; subject: string; messageId: string }>;
    retr(num: number): Promise<Buffer>;
    dele(num: number): Promise<void>;
    abort(): void;
    end(): Promise<void>;
  }> {
    const c = new Pop3Client();
    await c.connect(
      this.account.imapHost,
      this.account.imapPort,
      this.account.imapSecure,
      this.account.imapUser || this.account.email,
      this.password,
      !!this.account.allowSelfSigned,
    );
    return {
      uidl: () => c.uidl(),
      envelope: async (num: number) => {
        const raw = await c.top(num);
        // Message-ID vem de graça no TOP e serve de segunda chave de dedupe
        // (UIDL instável por sessão em alguns provedores).
        const mid = getHeader(raw, 'Message-ID').replace(/[<>]/g, '').trim();
        return {
          uid: num,
          from: getHeader(raw, 'From'),
          to: getHeader(raw, 'To'),
          subject: decodeMimeWords(getHeader(raw, 'Subject')),
          messageId: mid,
        };
      },
      retr: (num: number) => c.retr(num),
      dele: (num: number) => c.dele(num),
      abort: () => c.abort(),
      end: () => c.quit(),
    };
  }

  close(): void {
    // POP3 é conexão por operação — nada a fechar persistentemente
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private detailFromParsed(parsed: any, index: number): ItemDetail {
    const attachments = (parsed.attachments || []).map(
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
    const toVal = (parsed.to as { value?: Array<{ name?: string; address?: string }> } | undefined)?.value;
    const ccVal = (parsed.cc as { value?: Array<{ name?: string; address?: string }> } | undefined)?.value;
    const addrList = (l?: Array<{ name?: string; address?: string }>) =>
      (l || []).map((a) => (a.name ? `${a.name} <${a.address}>` : a.address || '')).join('; ');
    const uidl = this.numUidl.get(index);
    const cached = uidl ? this.headerCache[uidl] : undefined;
    return {
      index,
      messageClass: 'IPM.Note',
      subject: parsed.subject || '(sem assunto)',
      senderName: fromAddr?.name || fromAddr?.address || '',
      senderEmailAddress: fromAddr?.address || '',
      displayTo: addrList(toVal),
      displayCC: addrList(ccVal),
      sentTime: parsed.date ? parsed.date.toISOString() : null,
      deliveryTime: parsed.date ? parsed.date.toISOString() : null,
      hasAttachments: attachments.filter((a: { isInline: boolean }) => !a.isInline).length > 0,
      isRead: cached?.read ?? true,
      importance: 0,
      preview: (parsed.text || '').slice(0, 200),
      body: parsed.text || '',
      bodyHTML,
      transportMessageHeaders:
        parsed.headerLines?.map((h: { key: string; line: string }) => `${h.key}: ${h.line}`).join('\n') || '',
      attachments: attachments.filter((a: { isInline: boolean }) => !a.isInline),
    };
  }
}
