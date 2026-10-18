import { readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import { dataDir } from './dataDir';

/**
 * Cache de contatos para autocompletar — todo endereço para o qual o usuário
 * enviou email fica salvo em contact-history.json (userData) e reaparece como
 * sugestão ao digitar em Para/Cc/Cco, como o cache do Outlook (NK2).
 */
export interface HistoryContact {
  name: string;
  email: string;
  /** Quantas vezes foi usado — ordena as sugestões */
  count: number;
  lastUsed: number;
}

let cache: Map<string, HistoryContact> | null = null;

function file(): string {
  return join(dataDir(), 'contact-history.json');
}

function load(): Map<string, HistoryContact> {
  if (!cache) {
    cache = new Map();
    try {
      const arr = JSON.parse(readFileSync(file(), 'utf8')) as HistoryContact[];
      for (const c of arr) {
        if (c?.email) cache.set(c.email.toLowerCase(), c);
      }
    } catch {
      // primeira execução — histórico vazio
    }
  }
  return cache;
}

let saveTimer: ReturnType<typeof setTimeout> | null = null;

/**
 * Gravação com debounce: durante um Enviar/Receber o recordRecipients roda
 * POR mensagem — gravar o JSON a cada uma era milhares de escritas de disco.
 * Agora uma gravação acontece 3s depois da última atualização.
 */
function save(): void {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    saveTimer = null;
    try {
      writeFileSync(file(), JSON.stringify([...load().values()]));
    } catch {
      // falha de escrita não pode quebrar o envio
    }
  }, 3000);
}

const EMAIL_RE = /^[^\s@;,<>]+@[^\s@;,<>]+\.[^\s@;,<>]+$/;

/**
 * Registra os destinatários de um envio. Aceita lista separada por `;` ou `,`
 * e formatos "Nome <email>" / "Nome" <email> / email puro.
 */
export function recordRecipients(raw: string | undefined | null): void {
  if (!raw) return;
  const parts = raw
    .split(/[;,]/)
    .map((s) => s.trim())
    .filter(Boolean);
  let changed = false;
  for (const p of parts) {
    const m = p.match(/^(.*?)\s*<([^>]+)>\s*$/);
    const email = (m ? m[2] : p).trim();
    const name = m ? m[1].trim().replace(/^["']|["']$/g, '') : '';
    if (!EMAIL_RE.test(email)) continue;
    const key = email.toLowerCase();
    const cur = load().get(key);
    if (cur) {
      cur.count++;
      cur.lastUsed = Date.now();
      if (name && !cur.name) cur.name = name;
    } else {
      load().set(key, { name, email, count: 1, lastUsed: Date.now() });
    }
    changed = true;
  }
  if (changed) save();
}

/** Busca por trecho do nome ou do email — mais usados primeiro. */
export function searchContacts(q: string, limit = 8): HistoryContact[] {
  const needle = q.trim().toLowerCase();
  if (needle.length < 1) return [];
  return [...load().values()]
    .filter(
      (c) =>
        c.email.toLowerCase().includes(needle) ||
        c.name.toLowerCase().includes(needle),
    )
    .sort((a, b) => b.count - a.count || b.lastUsed - a.lastUsed)
    .slice(0, limit);
}
