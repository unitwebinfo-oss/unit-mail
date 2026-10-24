import { createLocalStore } from './localStore';
import { sourceManager } from './sourceManager';
import { PstService } from './pstService';
import { dedupContains, dedupNote } from './ruleEngine';
import type { FolderNode, MailSource } from '../shared/types';
import { safeFsName } from '../shared/fsName';

export interface MigrateProgress {
  phase: 'folders' | 'items' | 'done';
  currentFolder: string;
  done: number;
  total: number;
  /** Mensagens puladas por já existirem no destino (dedupe por Message-ID). */
  skipped?: number;
}

/** Mesma sanitização usada por PstService.createArchiveFolder — mantém os ids previsíveis. */
function sanitizeFolderName(name: string): string {
  return safeFsName(name);
}

/**
 * Migra TODO o conteúdo de um PST (ou base já aberta) para uma base nativa
 * do Unit Mail — diretório de .eml comuns, sem nenhuma dependência do
 * formato .pst. Depois da migração, excluir um item apaga o arquivo de
 * verdade (sem a limitação de "arquivo somente leitura" do PST).
 *
 * O PST original NUNCA é alterado — a migração só lê dele. Itens de
 * calendário/contatos/tarefas são convertidos para um .eml com o resumo
 * (assunto + corpo) — não preservam o formato ICS/vCard original.
 */
export async function migratePstToLocalStore(
  sourceId: string,
  destParentDir: string,
  storeName: string,
  onProgress: (p: MigrateProgress) => void,
  /** Caminho de uma base local EXISTENTE — importa dentro dela em vez de criar uma nova. */
  existingStorePath?: string,
  /**
   * Ids das pastas (FolderNode.id) a importar — marcar uma pasta inclui
   * toda a subárvore; undefined = importar tudo. Uma pasta marcada cujo
   * ancestral não foi marcado ainda preserva a hierarquia no destino: a
   * pasta-pai é criada (vazia) só para abrigar o caminho, mas os itens
   * dela não entram.
   */
  folderIds?: ReadonlySet<string>,
): Promise<MailSource> {
  const srcEntry = sourceManager.listAll().find((s) => s.id === sourceId);
  if (!srcEntry) throw new Error('Base de origem não encontrada');
  const srcSvc = sourceManager.backend(sourceId);
  if (!(srcSvc instanceof PstService)) throw new Error('Só é possível migrar arquivos .pst ou bases locais');
  if (existingStorePath === srcEntry.filePath) {
    throw new Error('O destino não pode ser o próprio PST de origem');
  }

  // Destino: base existente escolhida pelo usuário, ou uma nova criada no diretório pai.
  let destInfo: MailSource;
  if (existingStorePath) {
    destInfo = sourceManager.openLocalStore(existingStorePath);
  } else {
    const def = createLocalStore(storeName, destParentDir);
    destInfo = sourceManager.openLocalStore(def.path, def.name);
  }
  const destSvc = sourceManager.backend(destInfo.id);
  if (!(destSvc instanceof PstService)) throw new Error('Falha ao abrir a base de destino');

  const tree = srcSvc.getFolderTree();

  // Seleção de pastas: includeItems = marcadas + descendentes (as mensagens
  // que entram); includePath = includeItems + ancestrais (pastas criadas no
  // destino só para preservar o caminho — ficam vazias).
  let includeItems: Set<string> | null = null;
  let includePath: Set<string> | null = null;
  if (folderIds) {
    includeItems = new Set();
    includePath = new Set();
    const markItems = (nodes: FolderNode[], force: boolean) => {
      for (const n of nodes) {
        const inc = force || folderIds.has(n.id);
        if (inc) includeItems!.add(n.id);
        markItems(n.children, inc);
      }
    };
    markItems(tree, false);
    const markPath = (nodes: FolderNode[], chain: string[]) => {
      for (const n of nodes) {
        const next = [...chain, n.id];
        if (includeItems!.has(n.id)) for (const a of next) includePath!.add(a);
        markPath(n.children, next);
      }
    };
    markPath(tree, []);
  }

  // Primeira passada: conta o total de itens (para a barra de progresso).
  let total = 0;
  const countAll = (nodes: FolderNode[]) => {
    for (const n of nodes) {
      // Pasta fora do caminho selecionado não tem descendentes incluídos —
      // a subárvore inteira é podada (e nem é contada).
      if (includePath && !includePath.has(n.id)) continue;
      if (!includeItems || includeItems.has(n.id)) total += srcSvc.getItemCount(n.id);
      countAll(n.children);
    }
  };
  countAll(tree);
  onProgress({ phase: 'folders', currentFolder: '', done: 0, total });

  let done = 0;
  let skipped = 0;
  // Segunda passada: recria a hierarquia de pastas e copia os itens.
  const walk = async (nodes: FolderNode[], destParentArc: string | null) => {
    for (const node of nodes) {
      // Fora do caminho selecionado — nem a pasta nem os descendentes entram.
      if (includePath && !includePath.has(node.id)) continue;
      const clean = sanitizeFolderName(node.name);
      const destArcId = destParentArc ? `arc:${destParentArc}/${clean}` : `arc:${clean}`;
      // mkdir idempotente — se a base nova já tem a pasta padrão com esse
      // nome (ex.: "Caixa de Entrada"), os itens simplesmente se juntam.
      destSvc.createArchiveFolder(clean, destParentArc ? `arc:${destParentArc}` : null);

      // Pasta criada só como caminho (ancestral de uma pasta marcada) —
      // as mensagens dela não entram na seleção.
      const importItems = !includeItems || includeItems.has(node.id);
      // Diretório da pasta de destino — para o dedupe por Message-ID/hash.
      // Reimportar o mesmo PST (ou um PST que compartilha mensagens com a
      // base) não cria cópias duplicadas.
      const destDir = destSvc.writableDirFor(destArcId);
      const count = importItems ? srcSvc.getItemCount(node.id) : 0;
      const pageSize = 100;
      for (let offset = 0; offset < count; offset += pageSize) {
        const items = await srcSvc.getItems(node.id, offset, pageSize);
        for (const item of items) {
          try {
            const raw = await srcSvc.exportItemEml(node.id, item.index);
            if (destDir && (await dedupContains(destDir, raw))) {
              skipped++;
            } else {
              const name = await destSvc.importItemEml(destArcId, raw, item.isRead);
              if (destDir) await dedupNote(destDir, name, raw);
            }
          } catch {
            // item corrompido/ilegível — pula, não interrompe a migração
          }
          done++;
          if (done % 20 === 0 || done === total) {
            onProgress({ phase: 'items', currentFolder: node.name, done, total, skipped });
          }
        }
      }
      await walk(node.children, destParentArc ? `${destParentArc}/${clean}` : clean);
    }
  };
  await walk(tree, null);

  onProgress({ phase: 'done', currentFolder: '', done, total, skipped });
  return destInfo;
}
