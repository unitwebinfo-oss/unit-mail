import { randomUUID } from 'crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import { dataDir } from './dataDir';
import type { Signature } from '../shared/types';

function sigPath(): string {
  return join(dataDir(), 'signatures.json');
}

export class SignatureStore {
  list(): Signature[] {
    try {
      if (existsSync(sigPath())) return JSON.parse(readFileSync(sigPath(), 'utf-8'));
    } catch {
      // corrompido — começa do zero
    }
    return [];
  }

  get(id: string): Signature | undefined {
    return this.list().find((s) => s.id === id);
  }

  upsert(sig: Omit<Signature, 'id'> & { id?: string }): Signature {
    const all = this.list();
    // Sem id explícito: se já existe assinatura idêntica (mesmo nome e
    // conteúdo), reaproveita — evita duplicar a cada importação/migração.
    if (!sig.id) {
      const dup = all.find((s) => s.name === sig.name && s.html === sig.html);
      if (dup) return dup;
    }
    const full: Signature = { ...sig, id: sig.id || `sig-${randomUUID()}` };
    const idx = all.findIndex((s) => s.id === full.id);
    if (idx >= 0) all[idx] = full;
    else all.push(full);
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(sigPath(), JSON.stringify(all, null, 2));
    return full;
  }

  remove(id: string): void {
    mkdirSync(dataDir(), { recursive: true });
    writeFileSync(sigPath(), JSON.stringify(this.list().filter((s) => s.id !== id), null, 2));
  }
}

export const signatureStore = new SignatureStore();
