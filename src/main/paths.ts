import { createHash } from 'crypto';
import {
  existsSync,
  mkdirSync,
  accessSync,
  constants,
  readdirSync,
  cpSync,
  readFileSync,
  writeFileSync,
  statSync,
  openSync,
  readSync,
  closeSync,
} from 'fs';
import { basename, dirname, join, resolve } from 'path';
import { dataDir } from './dataDir';

/** Hash estável do caminho do PST — vincula o arquivo local à base. */
export function pstKeyForPath(filePath: string): string {
  return createHash('sha1').update(filePath.toLowerCase()).digest('hex').slice(0, 16);
}

// ---------- Registro de identidade do PST (sobrevive a mover o arquivo) ----------
//
// Tudo que o app grava de um PST vai para a pasta-irmã "<nome>.unitmail-data".
// Quando o usuário MOVE o .pst para outro diretório e o abre de lá, o caminho
// muda e — sem este registro — o app criaria uma sidecar nova e vazia: itens
// "excluídos" voltavam, leitura zerava e a ordem de pastas se perdia. O
// fingerprint (nome + tamanho + trechos de início/fim) identifica o MESMO
// arquivo mesmo num caminho novo; se o caminho antigo não existe mais, a
// sidecar registrada é adotada (copiada) para o novo local.

interface PstIndexEntry {
  fp: string;
  pstPath: string;
  sidecar: string;
}

function pstIndexPath(): string {
  return join(dataDir(), 'pst-index.json');
}

function loadPstIndex(): PstIndexEntry[] {
  try {
    const raw = JSON.parse(readFileSync(pstIndexPath(), 'utf8'));
    return Array.isArray(raw) ? raw : [];
  } catch {
    return [];
  }
}

function savePstIndex(entries: PstIndexEntry[]): void {
  try {
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(pstIndexPath(), JSON.stringify(entries, null, 1));
  } catch {
    // best-effort
  }
}

const normPath = (p: string) => resolve(p).toLowerCase();

/** Fingerprint de conteúdo — não leva o caminho em conta. ~rápido (lê 128KB). */
export function pstFingerprint(filePath: string): string {
  const st = statSync(filePath);
  const h = createHash('sha1').update(`${basename(filePath).toLowerCase()}|${st.size}|`);
  const fd = openSync(filePath, 'r');
  try {
    const headLen = Math.min(65536, st.size);
    const head = Buffer.alloc(headLen);
    readSync(fd, head, 0, headLen, 0);
    h.update(head);
    if (st.size > 65536) {
      const tail = Buffer.alloc(65536);
      readSync(fd, tail, 0, 65536, st.size - 65536);
      h.update(tail);
    }
  } finally {
    closeSync(fd);
  }
  return h.digest('hex').slice(0, 24);
}

/** Registra (ou atualiza) o par PST → sidecar no índice. */
function registerPstSidecar(pstPath: string, sidecar: string): void {
  try {
    const fp = pstFingerprint(pstPath);
    const entries = loadPstIndex().filter((e) => normPath(e.pstPath) !== normPath(pstPath));
    entries.push({ fp, pstPath, sidecar });
    savePstIndex(entries);
  } catch {
    // fingerprint falhou (arquivo travado?) — ignora
  }
}

/**
 * Se o PST neste caminho é o MESMO arquivo que já foi aberto de outro caminho
 * (fingerprint igual e o arquivo antigo não existe mais = foi movido), copia a
 * sidecar registrada para junto do novo caminho. Retorna o caminho antigo do
 * PST quando houve adoção — o chamador usa para remapear ordens de pasta.
 */
export function adoptSidecarForMovedPst(filePath: string): string | null {
  const sibling = join(
    dirname(filePath),
    basename(filePath).replace(/\.pst$/i, '') + '.unitmail-data',
  );
  try {
    if (existsSync(sibling) && readdirSync(sibling).length > 0) return null; // já tem dados
  } catch {
    return null;
  }
  let fp: string;
  try {
    fp = pstFingerprint(filePath);
  } catch {
    return null;
  }
  const cand = loadPstIndex().find(
    (e) =>
      e.fp === fp &&
      normPath(e.pstPath) !== normPath(filePath) &&
      !existsSync(e.pstPath) && // o caminho antigo sumiu = o arquivo foi movido
      existsSync(e.sidecar) &&
      readdirSync(e.sidecar).length > 0,
  );
  if (!cand) return null;
  try {
    mkdirSync(sibling, { recursive: true });
    for (const sub of readdirSync(cand.sidecar)) {
      const dst = join(sibling, sub);
      if (!existsSync(dst)) cpSync(join(cand.sidecar, sub), dst, { recursive: true });
    }
    return cand.pstPath;
  } catch {
    return null;
  }
}

/** Local legado da pasta-sombra (dentro de userData — escondido do usuário). */
function legacyArchiveRoot(filePath: string): string {
  return join(dataDir(), 'archives', pstKeyForPath(filePath));
}

/**
 * Diretório de dados graváveis de um PST — fica AO LADO do arquivo .pst
 * ("Outlook2.unitmail-data"), visível e portável: copiar o .pst + esta pasta
 * para o Linux preserva tudo que o app entregou ali. Se o diretório do PST
 * não for gravável (CD, pasta protegida), cai de volta para userData.
 */
export function archiveRootForPst(filePath: string): string {
  const sibling = join(
    dirname(filePath),
    basename(filePath).replace(/\.pst$/i, '') + '.unitmail-data',
  );
  try {
    mkdirSync(sibling, { recursive: true });
    accessSync(sibling, constants.W_OK);
    // Migra conteúdo da pasta-sombra legada (userData) se existir —
    // copia item a item; o legado permanece intacto como segurança.
    const legacy = legacyArchiveRoot(filePath);
    if (existsSync(legacy) && legacy !== sibling) {
      try {
        for (const sub of readdirSync(legacy)) {
          const dst = join(sibling, sub);
          if (!existsSync(dst)) cpSync(join(legacy, sub), dst, { recursive: true });
        }
      } catch {
        // migração best-effort — o conteúdo continua no legado
      }
    }
    registerPstSidecar(filePath, sibling);
    return sibling;
  } catch {
    const legacy = legacyArchiveRoot(filePath);
    registerPstSidecar(filePath, legacy);
    return legacy;
  }
}
