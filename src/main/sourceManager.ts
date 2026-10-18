import { createHash } from 'crypto';
import { basename, resolve } from 'path';
import { mkdirSync } from 'fs';
import { PstService } from './pstService';
import { archiveRootForPst, adoptSidecarForMovedPst } from './paths';
import { ImapMailSource } from './imapSource';
import { Pop3MailSource } from './pop3Source';
import { accountStore } from './accountStore';
import type { MailBackend } from './mailBackend';
import type { FolderNode, MailSource, NewSpecialItem } from '../shared/types';

interface Entry {
  info: MailSource;
  backend: MailBackend;
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
        info.error = (e as Error).message;
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
    if (!(src instanceof PstService) || !(dst instanceof PstService)) {
      throw new Error('Mover itens só é suportado em arquivos de dados (PST/base local)');
    }
    if (!dst.writableDirFor(dstFolderId)) {
      throw new Error('A pasta de destino não aceita itens');
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
    if (!(src instanceof PstService) || !(dst instanceof PstService)) {
      throw new Error('Mover itens só é suportado em arquivos de dados (PST/base local)');
    }
    if (!dst.writableDirFor(dstFolderId)) {
      throw new Error('A pasta de destino não aceita itens');
    }
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

  /** Move uma pasta local para dentro de outra (reparent). Só pastas graváveis. */
  moveFolder(sourceId: string, folderId: string, newParentFolderId: string | null): FolderNode[] {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) {
      throw new Error('Mover pastas só é suportado em arquivos de dados locais');
    }
    svc.moveArchiveFolder(folderId, newParentFolderId);
    return svc.getFolderTree();
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

  /** Cria uma pasta gravável na base (raiz ou dentro de outra pasta). */
  createFolder(sourceId: string, name: string, parentFolderId: string | null): FolderNode[] {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) {
      throw new Error('Criar pasta só é suportado em arquivos de dados locais');
    }
    svc.createArchiveFolder(name, parentFolderId);
    return svc.getFolderTree();
  }

  /** Renomeia uma pasta local. Retorna a árvore atualizada. */
  renameFolder(sourceId: string, folderId: string, newName: string): FolderNode[] {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) throw new Error('Só arquivos de dados');
    svc.renameArchiveFolder(folderId, newName);
    return svc.getFolderTree();
  }

  /** Exclui uma pasta local com todo o conteúdo. */
  deleteFolder(sourceId: string, folderId: string): FolderNode[] {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) throw new Error('Só arquivos de dados');
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

  /** Marca todos os itens locais da pasta como lidos. */
  markAllRead(sourceId: string, folderId: string): void {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) return;
    svc.markAllRead(folderId);
  }

  /** Define lido/não-lido de um item local de uma base. */
  setItemRead(sourceId: string, folderId: string, index: number, read: boolean): void {
    const svc = this.backend(sourceId);
    if (!(svc instanceof PstService)) return;
    svc.setItemRead(folderId, index, read);
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
