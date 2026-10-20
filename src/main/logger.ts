import { app } from 'electron';
import { appendFileSync, existsSync, mkdirSync, renameSync, statSync } from 'fs';
import { join } from 'path';
import { dataDir, profileEstablished } from './dataDir';

/** Log em arquivo na pasta de dados (logs/unitmail.log), com rotação por tamanho. */
const MAX_BYTES = 5 * 1024 * 1024;

function logDir(): string {
  // Antes do wizard escolher a pasta de dados (primeiro boot), NÃO cria a
  // raiz padrão só para registrar log — grava no userData (invisível).
  return profileEstablished()
    ? join(dataDir(), 'logs')
    : join(app.getPath('userData'), 'logs');
}

function write(level: 'INFO' | 'ERROR', scope: string, detail: unknown): void {
  try {
    const dir = logDir();
    const p = join(dir, 'unitmail.log');
    mkdirSync(dir, { recursive: true });
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
