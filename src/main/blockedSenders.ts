import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'fs';
import { join } from 'path';
import { dataDir } from './dataDir';
import { parseFrom } from './mailHeaders';

/**
 * Lista de remetentes bloqueados ("Blocked Senders" do Outlook).
 * Persistida em blocked-senders.json na pasta de dados — cada entrada é um
 * endereço completo ("spam@x.com") ou um domínio ("x.com", cobre também
 * subdomínios como "mail.x.com"). Tudo normalizado em minúsculas.
 */

const FILE = 'blocked-senders.json';

function filePath(): string {
  return join(dataDir(), FILE);
}

let cache: Set<string> | null = null;

function load(): Set<string> {
  if (cache) return cache;
  cache = new Set<string>();
  try {
    if (existsSync(filePath())) {
      const raw = JSON.parse(readFileSync(filePath(), 'utf-8')) as string[];
      for (const e of raw) {
        const n = normalize(e);
        if (n) cache.add(n);
      }
    }
  } catch {
    // corrompido — começa vazio
  }
  return cache;
}

function save(): void {
  try {
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(filePath(), JSON.stringify([...load()].sort(), null, 2));
  } catch {
    // falha de gravação — a lista em memória continua valendo nesta sessão
  }
}

/** Normaliza uma entrada: minúsculas, trim, sem "mailto:", sem "*@" inicial. */
function normalize(entry: string): string {
  let e = (entry || '').trim().toLowerCase();
  e = e.replace(/^mailto:/, '').trim();
  e = e.replace(/^[*@]+/, '').trim();
  return e;
}

/** Extrai o endereço de um cabeçalho From ("Nome <a@b.com>" ou "a@b.com"). */
function addrOf(fromHeader: string): string {
  const f = (fromHeader || '').trim();
  if (!f) return '';
  return parseFrom(f).addr.toLowerCase();
}

export const blockedSenders = {
  list(): string[] {
    return [...load()].sort();
  },

  add(entry: string): string[] {
    const n = normalize(entry);
    if (n) {
      load().add(n);
      save();
    }
    return this.list();
  },

  remove(entry: string): string[] {
    load().delete(normalize(entry));
    save();
    return this.list();
  },

  /**
   * O remetente está bloqueado? Casa endereço exato, domínio exato e
   * subdomínios do domínio bloqueado ("x.com" pega "a@x.com" e "a@m.x.com").
   */
  isBlocked(fromHeader: string): boolean {
    const addr = addrOf(fromHeader);
    if (!addr) return false;
    const at = addr.lastIndexOf('@');
    const domain = at >= 0 ? addr.slice(at + 1) : '';
    for (const e of load()) {
      if (addr === e) return true;
      if (domain && (domain === e || domain.endsWith(`.${e}`))) return true;
    }
    return false;
  },
};
