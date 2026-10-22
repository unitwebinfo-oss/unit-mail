import { createHash } from 'crypto';
import { basename, resolve } from 'path';
import { mkdirSync } from 'fs';
import { PstService } from './pstService';
import { archiveRootForPst, adoptSidecarForMovedPst } from './paths';
import { ImapMailSource } from './imapSource';
import { explainConnError } from './smtpService';
import { Pop3MailSource } from './pop3Source';
import { accountStore } from './accountStore';
import { folderKindOf } from '../shared/folderNames';
import type { MailBackend } from './mailBackend';
import type { FolderNode, MailSource, NewSpecialItem } from '../shared/types';

interface Entry {
  info: MailSource;
  backend: MailBackend;
}

/** Procura um nó de pasta pelo id em qualquer nível da árvore. */
function findNodeById(nodes: FolderNode[], id: string): FolderNode | null {
  for (const n of nodes) {
    if (n.id === id) return n;
    const hit = findNodeById(n.children, id);
    if (hit) return hit;
  }
  return null;
}

/**
 * Pastas reservadas do sistema — nunca excluídas pelo usuário. A quarentena
 * de antivírus em qualquer idioma ("Itens infectados" / "Infected Items").
 */
function isProtectedFolderName(name: string): boolean {
  return folderKindOf(name) === 'quarantine';
}

/**
 * Centraliza todas as bases abertas (arquivos PST e contas IMAP) e roteia
 * chamadas por sourceId para o backend correto.
 */
export class SourceManager {
  private entries = new Map<string, Entry>();

  /**
   * Chamado quando um PST aberto é reconhecido como o mesmo arquivo que já
   * existia em outro caminho (movido) e sua sidecar foi adotada no novo local.
   * O chamador remapeia referências persistidas (ordem de pastas, entrega…).
   */
  onPstMoved?: (oldPath: string, newPath: string) => void;

  /** Fontes visíveis na árvore (PSTs, bases locais, contas IMAP). POP3 é oculto. */
  list(): MailSource[] {
    return Array.from(this.entries.values())
      .map((e) => e.info)
      .filter((s) => !s.hidden);
  }

  /** Todas as fontes registradas, incluindo transportes ocultos (POP3) — uso interno. */
  listAll(): MailSource[] {
    return Array.from(this.entries.values()).map((e) => e.info);
  }

  private getEntry(sourceId: string): Entry {
    const e = this.entries.get(sourceId);
    if (!e) throw new Error(`Base não encontrada: ${sourceId}`);
    return e;
  }

  backend(sourceId: string): MailBackend {
    return this.getEntry(sourceId).backend;
  }

  /**
   * Id estável por arquivo — o mesmo caminho sempre gera o mesmo sourceId,
   * para que estado persistido por base (ordem/expandido/favoritos, chaveado
   * por sourceId) sobreviva a reinícios. Também evita abrir duplicado.
   */
  private stableId(prefix: string, filePath: string): string {
    const h = createHash('sha1')
      .update(resolve(filePath).toLowerCase())
      .digest('hex')
      .slice(0, 16);
    return `${prefix}:${h}`;
  }

  openPst(filePath: string): MailSource {
    const id = this.stableId('pst', filePath);
    const existing = this.entries.get(id);
    if (existing) return existing.info;
    // Se este arquivo já foi aberto de outro caminho (foi movido), adota a
    // sidecar antiga — preserva itens excluídos, lidos e mensagens entregues.
    const movedFrom = adoptSidecarForMovedPst(filePath);
    const archiveDir = archiveRootForPst(filePath);
    if (movedFrom) this.onPstMoved?.(movedFrom, filePath);
    mkdirSync(archiveDir, { recursive: true });
    const svc = new PstService(archiveDir);
    const result = svc.open(filePath);
    const info: MailSource = { id, kind: 'pst', label: result.storeName, filePath };
    this.entries.set(id, { info, backend: svc });
    return info;
  }

  /** Abre uma base de dados local (diretório de .eml criado pelo app). */
  openLocalStore(storePath: string, name?: string): MailSource {
    const existing = Array.from(this.entries.values()).find(
      (e) => e.info.kind === 'local' && e.info.filePath === storePath,
    );
    if (existing) return existing.info;
    const id = this.stableId('local', storePath);
    const label = name || basename(storePath);
    const svc = new PstService(storePath);
    svc.openLocal(label, storePath);
    const info: MailSource = { id, kind: 'local', label, filePath: storePath };
    this.entries.set(id, { info, backend: svc });
    return info;
  }

  async openAccount(accountId: string): Promise<MailSource> {
    const account = accountStore.getFull(accountId);
    if (!account) throw new Error('Conta não encontrada');
    const id = `acct:${accountId}`;
    // Evita abrir conexão duplicada se já estiver aberta
    const existing = this.entries.get(id);
    if (existing) return existing.info;

    const password = accountStore.getPassword(accountId);
    const backend: MailBackend =
      account.protocol === 'pop3'
        ? new Pop3MailSource(account, password)
        : new ImapMailSource(account, password);
    const info: MailSource = {
      id,
      kind: 'imap',
      label: account.email,
      connecting: true,
      // POP3 não é uma base de armazenamento — entrega na base vinculada
      hidden: account.protocol === 'pop3',
    };
    this.entries.set(id, { info, backend });
    if (backend instanceof ImapMailSource) {
      try {
        await backend.connect();
        info.connecting = false;
      } catch (e) {
        info.connecting = false;
        info.error = explainConnError(e, 'IMAP', {
          host: account.imapHost,
          port: account.imapPort,
          secure: account.imapSecure,
          user: account.imapUser || account.email,
        });
      }
    } else {
      info.connecting = false; // POP3 conecta por operação
    }
    return info;
  }

  /**
   * Move um item entre pastas (arrastar e soltar). Retorna 'movido' ou
   * 'copiado' — de pasta PST só dá para copiar (PST é somente leitura).
   */
  async moveItem(
    srcId: string,
    srcFolderId: string,
    index: number,
    dstId: string,
    dstFolderId: string,
  ): Promise<'movido' | 'copiado'> {
    const src = this.backend(srcId);
    const dst = this.backend(dstId);
    const srcImap = src instanceof ImapMailSource;
    const dstImap = dst instanceof ImapMailSource;
    if (!srcImap && !(src instanceof PstService)) throw new Error('Fonte de origem não aceita mensagens');
    if (!dstImap && !(dst instanceof PstService)) throw new Error('Destino não aceita mensagens');
    if (!dstImap && dst instanceof PstService && !dst.writableDirFor(dstFolderId)) {
      throw new Error('A pasta de destino não aceita itens');
    }
    // Mesma conta IMAP → MOVE nativo no servidor (sem baixar/subir a
    // mensagem). Contas IMAP diferentes caem no fluxo genérico de upload.
    if (srcImap && src === dst) {
      await src.moveWithin(srcFolderId, index, dstFolderId);
      return 'movido';
    }
    // IMAP não usa diretórios locais. Exporta a mensagem e faz APPEND no
    // servidor; só remove a origem depois que o upload terminou. Isso permite
    // migrar .eml/PST para uma conta online e também mover entre contas.
    if (dstImap) {
      const raw = srcImap
        ? await src.downloadRaw(srcFolderId, index)
        : await (src as PstService).exportItemEml(srcFolderId, index);
      const wasRead = src instanceof PstService
        ? src.isItemRead?.(srcFolderId, index)
        : true;
      await dst.appendItem(dstFolderId, raw, wasRead);
      if (srcImap) await src.expungeItem(srcFolderId, index);
      else if ((src as PstService).isLocalItem(srcFolderId, index)) (src as PstService).deleteItem(srcFolderId, index);
      else (src as PstService).hideItem(srcFolderId, index);
      return 'movido';
    }
    if (srcImap) {
      const raw = await src.downloadRaw(srcFolderId, index);
      await (dst as PstService).importItemEml(dstFolderId, raw, true);
      await src.expungeItem(srcFolderId, index);
      return 'movido';
    }
    // Lê o estado de leitura ANTES do export — para itens do PST o export
    // já marca como lido (efeito colateral de getItem). Assim um item não
    // lido deletado continua não lido na lixeira, e vice-versa.
    const wasRead = src.isItemRead ? src.isItemRead(srcFolderId, index) : undefined;
    const cats = src.getItemCategories?.(srcFolderId, index) ?? [];
    const raw = await src.exportItemEml(srcFolderId, index);
    const name = await dst.importItemEml(dstFolderId, raw, wasRead);
    if (cats.length) dst.setItemCategoriesByFile?.(dstFolderId, name, cats);
    if (!src.isLocalItem(srcFolderId, index)) {
      // Item gravado no PST: copia para o destino e oculta o original —
      // na interface parece um "mover" de verdade (o .pst não pode ser
      // modificado fisicamente, então a exclusão é lógica).
      src.hideItem(srcFolderId, index);
      return 'movido';
    }
    src.deleteItem(srcFolderId, index);
    return 'movido';
  }

  /**
   * Move vários itens num lote — as gravações de estado em disco
   * (readstate/hidden) e os recounts da árvore acontecem uma única vez no
   * fim, em vez de por item. Retorna quantos foram movidos.
   */
  async moveItems(
    srcId: string,
    srcFolderId: string,
    indexes: number[],
    dstId: string,
    dstFolderId: string,
  ): Promise<number> {
    const src = this.backend(srcId);
    const dst = this.backend(dstId);
    if (src instanceof ImapMailSource || dst instanceof ImapMailSource) {
      let moved = 0;
      const run = async () => {
        for (const ix of [...indexes].sort((a, b) => b - a)) {
          await this.moveItem(srcId, srcFolderId, ix, dstId, dstFolderId);
          moved++;
        }
      };
      // Origem local → batch grava readstate/hidden uma vez no fim do lote,
      // não por item (migração de centenas de mensagens para o servidor).
      if (src instanceof PstService) await src.batch(run);
      else await run();
      return moved;
    }
    if (!(src instanceof PstService) || !(dst instanceof PstService)) {
      throw new Error('Mover itens só é suportado em arquivos de dados');
    }
    if (!dst.writableDirFor(dstFolderId)) throw new Error('A pasta de destino não aceita itens');
    let moved = 0;
    // Ordem decrescente: remover um item não desloca os índices menores que
    // ainda serão processados — e itens do PST (índices mais altos) são
    // resolvidos antes do arcCount encolher com as remoções de .eml locais.
    const run = async () => {
      for (const ix of [...indexes].sort((a, b) => b - a)) {
        try {
          await this.moveItem(srcId, srcFolderId, ix, dstId, dstFolderId);
          moved++;
        } catch {
          // item sumiu/índice já deslocado — segue com os demais
        }
      }
    };
    await src.batch(() => dst.batch(run));
    return moved;
  }

  /** Exclui vários itens de uma pasta num lote (mesma economia de I/O). */
  async deleteItems(sourceId: string, folderId: string, indexes: number[]): Promise<number> {
    const svc = this.backend(sourceId);
    // Conta IMAP: exclusão vai para Itens Excluídos do servidor (ou expurgo
    // quando a pasta não existe) — o índice do item é o UID da mensagem.
    if (svc instanceof ImapMailSource) {
      let n = 0;
      for (const ix of [...indexes].sort((a, b) => b - a)) {
        try {
          await svc.deleteItem(folderId, ix);
          n++;
        } catch {
          // mensagem já removida no servidor — segue com as demais
        }
      }
      return n;
    }
    if (!(svc instanceof PstService)) {
      throw new Error('Excluir só é suportado em arquivos de dados (PST/base local)');
    }
    let n = 0;
    await svc.batch(() => {
      for (const ix of [...indexes].sort((a, b) => b - a)) {
        try {
          svc.deleteItem(folderId, ix);
          n++;
        } catch {
          // item já inexistente — segue
        }
      }
    });
    return n;
  }

  /** Move uma pasta para dentro de outra — local reparenta na base, IMAP
   *  renomeia o caminho no servidor. Entre bases diferentes copia a pasta
   *  inteira (itens + subpastas) para o destino e depois exclui a origem.
   *  newParentId null = raiz. Retorna as árvores atualizadas por fonte. */
  async moveFolder(
    sourceId: string,
    folderId: string,
    newParentFolderId: string | null,
    dstSourceId?: string,
  ): Promise<Record<string, FolderNode[]>> {
    const src = this.backend(sourceId);
    const dstId = dstSourceId || sourceId;
    const dst = this.backend(dstId);
    // Mesma base: reparent nativo (mover dentro da mesma árvore).
    if (dst === src) {
      if (src instanceof ImapMailSource) {
        return { [sourceId]: await src.moveFolder(folderId, newParentFolderId) };
      }
      if (!(src instanceof PstService)) {
        throw new Error('Mover pastas só é suportado em arquivos de dados locais');
      }
      src.moveArchiveFolder(folderId, newParentFolderId);
      return { [sourceId]: src.getFolderTree() };
    }
    // Entre bases: copia tudo e só então exclui a origem — sem perda se a
    // cópia falhar no meio (a pasta original permanece).
    if (!(dst instanceof ImapMailSource) && !(dst instanceof PstService)) {
      throw new Error('O destino não aceita pastas');
    }
    if (!(src instanceof ImapMailSource) && !(src instanceof PstService)) {
      throw new Error('Mover pastas só é suportado de arquivos de dados ou IMAP');
    }
    const tree = await Promise.resolve(src.getFolderTree());
    const node = findNodeById(tree, folderId);
    if (!node) throw new Error('Pasta não encontrada');
    await this.copyFolderRecursive(src, node, dst, newParentFolderId);
    if (src instanceof ImapMailSource) await src.deleteFolder(folderId);
    else (src as PstService).deleteArchiveFolder(folderId);
    return {
      [sourceId]: await Promise.resolve(src.getFolderTree()),
      [dstId]: await Promise.resolve(dst.getFolderTree()),
    };
  }

  /** Copia recursiva de pasta entre backends: cria a pasta no destino,
   *  copia todos os itens (.eml bruto preserva lido/categorias) e repete
   *  para cada subpasta. Retorna nada — falha aborta antes do delete. */
  private async copyFolderRecursive(
    src: MailBackend,
    node: FolderNode,
    dst: MailBackend,
    parentId: string | null,
  ): Promise<void> {
    // 1. Cria a pasta no destino e descobre o id novo (procura na árvore
    //    do pai — mesmo nome pode já existir em outro nível).
    const findCreated = (tree: FolderNode[]): string | null => {
      const parentNode = parentId ? findNodeById(tree, parentId) : null;
      const kids = parentNode ? parentNode.children : tree;
      return kids.find((n) => n.name === node.name)?.id || null;
    };
    let newId: string | null = null;
    if (dst instanceof ImapMailSource) {
      newId = findCreated(await dst.createFolder(node.name, parentId));
    } else if (dst instanceof PstService) {
      dst.createArchiveFolder(node.name, parentId);
      newId = findCreated(dst.getFolderTree());
    }
    if (!newId) throw new Error(`Não foi possível criar "${node.name}" no destino`);
    // 2. Copia todos os itens — index IMAP é UID, então usa os índices
    //    reais do listItems, nunca 0..n-1.
    const dstImap = dst instanceof ImapMailSource;
    const items = await Promise.resolve(src.getItems(node.id, 0, 1000000));
    const copyAll = async () => {
      for (const it of items) {
        const wasRead = src.isItemRead ? src.isItemRead(node.id, it.index) : true;
        const raw = src instanceof ImapMailSource
          ? await src.downloadRaw(node.id, it.index)
          : await (src as PstService).exportItemEml(node.id, it.index);
        if (dstImap) {
          await (dst as ImapMailSource).appendItem(newId, raw, wasRead ?? true);
        } else {
          const fname = await (dst as PstService).importItemEml(newId, raw, wasRead);
          const cats = (src as PstService).getItemCategories?.(node.id, it.index) ?? [];
          if (cats.length) (dst as PstService).setItemCategoriesByFile?.(newId, fname, cats);
        }
      }
    };
    if (dst instanceof PstService) await dst.batch(copyAll);
    else await copyAll();
    // 3. Subpastas por último — mesmo processo recursivo.
    for (const c of node.children) {
      await this.copyFolderRecursive(src, c, dst, newId);
    }
  }

  /** Cria compromisso/contato/tarefa na base — retorna a árvore atualizada. */
  async createSpecialItem(sourceId: string, item: NewSpecialItem): Promise<FolderNode[]> {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) {
      throw new Error('Criar itens só é suportado em arquivos de dados');
    }
    await svc.createSpecialItem(item);
    return svc.getFolderTree();
  }

  /** Edita compromisso/contato/tarefa — regrava o .eml e retorna árvore + item
   *  efetivo (com uid/seq do compromisso, para reenviar o convite). */
  async updateSpecialItem(
    sourceId: string,
    folderId: string,
    index: number,
    item: NewSpecialItem,
  ): Promise<{ tree: FolderNode[]; item: NewSpecialItem }> {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) {
      throw new Error('Editar itens só é suportado em arquivos de dados');
    }
    const merged = await svc.updateSpecialItem(folderId, index, item);
    return { tree: svc.getFolderTree(), item: merged };
  }

  /** Cria uma pasta gravável na base (raiz ou dentro de outra pasta). */
  async createFolder(sourceId: string, name: string, parentFolderId: string | null): Promise<FolderNode[]> {
    const svc = this.backend(sourceId);
    if (svc instanceof ImapMailSource) {
      return svc.createFolder(name, parentFolderId);
    }
    if (!(svc instanceof PstService)) {
      throw new Error('Esta fonte não aceita criação de pastas');
    }
    svc.createArchiveFolder(name, parentFolderId);
    return svc.getFolderTree();
  }

  /** Renomeia uma pasta — local renomeia na base, IMAP renomeia no servidor. */
  async renameFolder(sourceId: string, folderId: string, newName: string): Promise<FolderNode[]> {
    const svc = this.backend(sourceId);
    if (svc instanceof ImapMailSource) {
      return svc.renameFolder(folderId, newName);
    }
    if (!(svc instanceof PstService)) throw new Error('Só arquivos de dados');
    const node = findNodeById(svc.getFolderTree(), folderId);
    if (node && isProtectedFolderName(node.name)) {
      throw new Error(
        `A pasta "${node.name}" é protegida pelo sistema (quarentena de antivírus) e não pode ser renomeada.`,
      );
    }
    svc.renameArchiveFolder(folderId, newName);
    return svc.getFolderTree();
  }

  /** Exclui uma pasta com todo o conteúdo — local apaga da base, IMAP apaga no servidor. */
  async deleteFolder(sourceId: string, folderId: string): Promise<FolderNode[]> {
    const svc = this.backend(sourceId);
    if (svc instanceof ImapMailSource) {
      return svc.deleteFolder(folderId);
    }
    if (!(svc instanceof PstService)) throw new Error('Só arquivos de dados');
    const tree = svc.getFolderTree();
    const node = findNodeById(tree, folderId);
    if (node && isProtectedFolderName(node.name)) {
      throw new Error(
        `A pasta "${node.name}" é protegida pelo sistema (quarentena de antivírus) e não pode ser excluída.`,
      );
    }
    svc.deleteArchiveFolder(folderId);
    return svc.getFolderTree();
  }

  /** Copia uma pasta local para "<nome> - Cópia". */
  copyFolder(sourceId: string, folderId: string): FolderNode[] {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) throw new Error('Só arquivos de dados');
    svc.copyArchiveFolder(folderId);
    return svc.getFolderTree();
  }

  /** Oculta/reexibe uma pasta na árvore (não apaga conteúdo). */
  setFolderHidden(sourceId: string, folderId: string, hidden: boolean): FolderNode[] {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) throw new Error('Só arquivos de dados');
    return svc.setFolderHidden(folderId, hidden);
  }

  /** Remove todos os itens locais de uma pasta. */
  emptyFolder(sourceId: string, folderId: string): number {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) throw new Error('Só arquivos de dados');
    return svc.emptyFolder(folderId);
  }

  /** Marca todos os itens da pasta como lidos (local em disco, IMAP no servidor). */
  async markAllRead(sourceId: string, folderId: string): Promise<void> {
    const svc = this.backend(sourceId);
    if (svc instanceof PstService) {
      svc.markAllRead(folderId);
      return;
    }
    if (svc instanceof ImapMailSource) {
      await svc.markAllRead(folderId);
      return;
    }
    throw new Error('Marcar como lidas não é suportado nesta fonte.');
  }

  /** Define lido/não-lido — local grava na base, IMAP muda a flag \Seen no servidor. */
  async setItemRead(sourceId: string, folderId: string, index: number, read: boolean): Promise<void> {
    const svc = this.backend(sourceId);
    if (svc instanceof PstService) {
      svc.setItemRead(folderId, index, read);
      return;
    }
    if (svc instanceof ImapMailSource) {
      await svc.setItemRead(folderId, index, read);
    }
  }

  async close(sourceId: string): Promise<void> {
    const e = this.entries.get(sourceId);
    if (!e) return;
    await e.backend.close();
    this.entries.delete(sourceId);
  }

  async closeAll(): Promise<void> {
    await Promise.all(Array.from(this.entries.keys()).map((id) => this.close(id)));
  }

  /** PstService já aberto cuja base é `storePath` (PST ou base local). */
  serviceForPath(storePath: string): PstService | null {
    for (const e of this.entries.values()) {
      if (
        (e.info.kind === 'local' || e.info.kind === 'pst') &&
        e.info.filePath === storePath &&
        e.backend instanceof PstService
      )
        return e.backend;
    }
    return null;
  }

  labelForPath(filePath: string): string {
    return basename(filePath);
  }
}

export const sourceManager = new SourceManager();
