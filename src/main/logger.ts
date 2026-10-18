import { appendFileSync, existsSync, mkdirSync, renameSync, statSync } from 'fs';
import { join } from 'path';
import { dataDir } from './dataDir';

/** Log em arquivo na pasta de dados (logs/unitmail.log), com rotação por tamanho. */
const MAX_BYTES = 5 * 1024 * 1024;

function logPath(): string {
  return join(dataDir(), 'logs', 'unitmail.log');
}

function write(level: 'INFO' | 'ERROR', scope: string, detail: unknown): void {
  try {
    const p = logPath();
    mkdirSync(join(dataDir(), 'logs'), { recursive: true });
    if (existsSync(p) && statSync(p).size > MAX_BYTES) renameSync(p, `${p}.1`);
    const msg = detail instanceof Error ? detail.stack || detail.message : String(detail);
    appendFileSync(p, `${new Date().toISOString()} ${level} [${scope}] ${msg}\n`);
  } catch {
    // log nunca pode derrubar o app
  }
}

export const log = {
  info: (scope: string, detail: unknown) => write('INFO', scope, detail),
  error: (scope: string, detail: unknown) => write('ERROR', scope, detail),
};
