import { randomUUID } from 'crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'fs';
import { join } from 'path';
import { dataDir } from './dataDir';
import { safeFsName } from '../shared/fsName';
import { folderKindOf, folderName, type FolderKind } from '../shared/folderNames';
import { appLanguage } from './language';

/**
 * Base de dados local ("novo arquivo de dados"). Como não é possível gerar
 * um .pst binário com bibliotecas abertas, a base local é um diretório que
 * o próprio usuário escolhe, contendo store.json + uma subpasta por pasta
 * de email, cada uma guardando mensagens .eml. É portável: copiar a pasta
 * para o Linux preserva tudo.
 */
export interface LocalStoreDef {
  id: string;
  name: string;
  path: string;
}

/**
 * Garante que as pastas padrão existem — cobre bases criadas por versões
 * antigas. Para cada tipo semântico só cria se NENHUMA variante de nome já
 * existir (ex.: base com "Inbox" não ganha uma "Caixa de Entrada" vazia ao
 * lado) — pastas existentes nunca são renomeadas nem duplicadas.
 */
function ensureDefaultFolders(path: string): void {
  const kinds: FolderKind[] = [
    'inbox', 'sent', 'deleted', 'quarantine', 'drafts', 'outbox',
    'junk', 'archive', 'calendar', 'contacts', 'tasks',
  ];
  try {
    const existing = readdirSync(path, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => folderKindOf(d.name));
    const lang = appLanguage();
    for (const kind of kinds) {
      if (!existing.includes(kind)) mkdirSync(join(path, folderName(kind, lang)), { recursive: true });
    }
  } catch {
    // sem permissão — a pasta será criada sob demanda na primeira entrega
  }
}

function registryPath(): string {
  return join(dataDir(), 'localstores.json');
}

export function listLocalStores(): LocalStoreDef[] {
  try {
    if (existsSync(registryPath())) return JSON.parse(readFileSync(registryPath(), 'utf-8'));
  } catch {
    // corrompido — começa do zero
  }
  return [];
}

function saveRegistry(stores: LocalStoreDef[]): void {
  mkdirSync(dataDir(), { recursive: true });
  writeFileSync(registryPath(), JSON.stringify(stores, null, 2));
}

export function isLocalStorePath(p: string): boolean {
  return listLocalStores().some((s) => s.path.toLowerCase() === p.toLowerCase());
}

function sanitizeDirName(name: string): string {
  return safeFsName(name, 'Base de Dados');
}

/** Cria um novo arquivo de dados local dentro de parentDir. */
export function createLocalStore(name: string, parentDir: string): LocalStoreDef {
  const dirName = sanitizeDirName(name);
  let path = join(parentDir, dirName);
  let finalName = name;
  let i = 2;
  // Quando o nome colide com uma pasta existente (ex.: várias contas com o
  // mesmo nome de exibição), o nome EXIBIDO precisa acompanhar o sufixo da
  // pasta real — senão bases diferentes aparecem com o mesmo rótulo na UI.
  while (existsSync(path)) {
    path = join(parentDir, `${dirName} (${i})`);
    finalName = `${name} (${i})`;
    i++;
  }

  mkdirSync(path, { recursive: true });
  ensureDefaultFolders(path);
  writeFileSync(
    join(path, 'store.json'),
    JSON.stringify({ name: finalName, format: 'outlook-linux-store@1', created: new Date().toISOString() }, null, 2),
  );

  const def: LocalStoreDef = { id: `ls-${randomUUID()}`, name: finalName, path };
  const stores = listLocalStores();
  stores.push(def);
  saveRegistry(stores);
  return def;
}

/** Apaga de verdade uma base local: fecha o registro e remove a pasta do disco. */
export function deleteLocalStore(path: string): void {
  unregisterLocalStore(path);
  try {
    rmSync(path, { recursive: true, force: true });
  } catch {
    // já removida ou sem permissão — registro já foi limpo
  }
}

/** Registra um diretório já existente como base local (ex.: ao restaurar sessão). */
export function registerLocalStore(path: string): LocalStoreDef | null {
  const existing = listLocalStores().find((s) => s.path.toLowerCase() === path.toLowerCase());
  if (existing) {
    ensureDefaultFolders(path);
    return existing;
  }
  let name = path.split(/[\\/]/).pop() || 'Base Local';
  try {
    const meta = join(path, 'store.json');
    if (existsSync(meta)) name = JSON.parse(readFileSync(meta, 'utf-8')).name || name;
  } catch {
    // usa o nome da pasta
  }
  const def: LocalStoreDef = { id: `ls-${randomUUID()}`, name, path };
  const stores = listLocalStores();
  stores.push(def);
  saveRegistry(stores);
  return def;
}

export function unregisterLocalStore(path: string): void {
  saveRegistry(listLocalStores().filter((s) => s.path.toLowerCase() !== path.toLowerCase()));
}

/** Atualiza o caminho registrado de uma base — usado quando um backup é
 *  restaurado noutra máquina/SO e o caminho antigo (ex.: "D:\…") muda. */
export function remapLocalStorePath(oldPath: string, newPath: string): void {
  const stores = listLocalStores();
  const def = stores.find(
    (s) => s.path.toLowerCase() === oldPath.toLowerCase() ||
      s.path.split(/[\\/]+/).join('/').toLowerCase() ===
        oldPath.split(/[\\/]+/).join('/').toLowerCase(),
  );
  if (!def) return;
  def.path = newPath;
  saveRegistry(stores);
}

/** Renomeia a base local (atualiza store.json + registro). */
export function renameLocalStore(path: string, newName: string): LocalStoreDef | null {
  const stores = listLocalStores();
  const def = stores.find((s) => s.path.toLowerCase() === path.toLowerCase());
  if (!def) return null;
  def.name = newName;
  saveRegistry(stores);
  try {
    const meta = join(path, 'store.json');
    if (existsSync(meta)) {
      const j = JSON.parse(readFileSync(meta, 'utf-8'));
      j.name = newName;
      writeFileSync(meta, JSON.stringify(j, null, 2));
    }
  } catch {
    // metadados secundários — registro já foi atualizado
  }
  return def;
}
