import { createLocalStore } from './localStore';
import { sourceManager } from './sourceManager';
import { PstService } from './pstService';
import {
  countMboxMessages,
  mboxMessages,
  mboxMessageToEml,
  scanThunderbirdMail,
} from './thunderbirdImport';
import { dedupContains, dedupNote } from './ruleEngine';
import type { MigrateProgress } from './pstImport';
import type { MailSource, TbMailFolder } from '../shared/types';
import { safeFsName } from '../shared/fsName';

function sanitizeFolderName(name: string): string {
  return safeFsName(name);
}

/**
 * Importa TODAS as mensagens de um perfil do Thunderbird para uma base nativa
 * do Unit Mail — cada armazenamento (Pastas Locais, conta POP/IMAP) vira uma
 * pasta de primeiro nível na base de destino, com a hierarquia preservada.
 * O perfil do Thunderbird nunca é alterado — só leitura dos .mbox.
 */
export async function importThunderbirdMail(
  profilePath: string,
  destParentDir: string,
  storeName: string,
  onProgress: (p: MigrateProgress) => void,
  existingStorePath?: string,
): Promise<MailSource> {
  const profile = scanThunderbirdMail().find((p) => p.path === profilePath);
  if (!profile) throw new Error('Perfil do Thunderbird não encontrado ou sem mensagens');

  let destInfo: MailSource;
  if (existingStorePath) {
    destInfo = sourceManager.openLocalStore(existingStorePath);
  } else {
    const def = createLocalStore(storeName, destParentDir);
    destInfo = sourceManager.openLocalStore(def.path, def.name);
  }
  const destSvc = sourceManager.backend(destInfo.id);
  if (!(destSvc instanceof PstService)) throw new Error('Falha ao abrir a base de destino');

  // Primeira passada: conta mensagens para a barra de progresso.
  let total = 0;
  const countTree = async (folders: TbMailFolder[]) => {
    for (const f of folders) {
      try {
        total += await countMboxMessages(f.mbox);
      } catch {
        // arquivo ilegível — será pulado na importação também
      }
      await countTree(f.children);
    }
  };
  for (const s of profile.stores) await countTree(s.folders);
  onProgress({ phase: 'folders', currentFolder: '', done: 0, total });

  let done = 0;
  let skipped = 0;
  await destSvc.batch(async () => {
    const walk = async (folders: TbMailFolder[], parentArc: string) => {
      for (const f of folders) {
        const clean = sanitizeFolderName(f.name);
        const arc = `${parentArc}/${clean}`;
        destSvc.createArchiveFolder(clean, `arc:${parentArc}`);
        // Dedupe por Message-ID/hash — reimportar o mesmo perfil (ou uma
        // base que já tem parte das mensagens) não cria duplicatas.
        const destDir = destSvc.writableDirFor(`arc:${arc}`);
        try {
          for await (const raw of mboxMessages(f.mbox)) {
            try {
              const conv = mboxMessageToEml(raw);
              if (conv) {
                if (destDir && (await dedupContains(destDir, conv.eml))) {
                  skipped++;
                } else {
                  const name = await destSvc.importItemEml(`arc:${arc}`, conv.eml, conv.read);
                  if (destDir) await dedupNote(destDir, name, conv.eml);
                }
              }
            } catch {
              // mensagem corrompida — pula sem interromper a importação
            }
            done++;
            if (done % 20 === 0 || done === total) {
              onProgress({ phase: 'items', currentFolder: f.name, done, total, skipped });
            }
          }
        } catch {
          // mbox ilegível — segue para a próxima pasta
        }
        await walk(f.children, arc);
      }
    };
    for (const s of profile.stores) {
      const storeArc = sanitizeFolderName(s.label);
      destSvc.createArchiveFolder(storeArc, null);
      await walk(s.folders, storeArc);
    }
  });

  onProgress({ phase: 'done', currentFolder: '', done, total, skipped });
  return destInfo;
}
