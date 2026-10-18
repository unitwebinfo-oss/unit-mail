import type {
  AttachmentData,
  FolderNode,
  ItemDetail,
  ItemSummary,
  SearchResult,
} from '../shared/types';

/**
 * Contrato comum implementado por qualquer fonte de email (arquivo PST ou
 * conta IMAP). O SourceManager roteia as chamadas de IPC para a implementação
 * correta com base no prefixo do sourceId.
 */
export interface MailBackend {
  getFolderTree(): FolderNode[] | Promise<FolderNode[]>;
  getItemCount(folderId: string): number | Promise<number>;
  getItems(folderId: string, offset: number, limit: number): ItemSummary[] | Promise<ItemSummary[]>;
  getItem(folderId: string, index: number): ItemDetail | Promise<ItemDetail>;
  getAttachmentData(
    folderId: string,
    index: number,
    attachIndex: number,
  ): AttachmentData | Promise<AttachmentData>;
  /**
   * Busca na fonte. `onBatch` (opcional) recebe lotes de hits conforme as
   * pastas são varridas — a UI mostra resultados ao vivo em vez de esperar
   * a varredura inteira terminar. `folderId` restringe à pasta atual.
   */
  search(
    query: string,
    containerClass?: string,
    maxResults?: number,
    onBatch?: (hits: SearchResult[]) => void,
    folderId?: string,
  ): SearchResult[] | Promise<SearchResult[]>;
  /** Estado de leitura de um item (bases locais/PST). */
  isItemRead?(folderId: string, index: number): boolean;
  close(): void | Promise<void>;
}
