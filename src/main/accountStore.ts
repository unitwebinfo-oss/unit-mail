import { safeStorage } from 'electron';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import { createCipheriv, createDecipheriv, randomBytes, randomUUID } from 'crypto';
import { dataDir } from './dataDir';
import type { AccountConfig, AccountDraft } from '../shared/types';

interface StoredAccount extends AccountConfig {
  /** Senha criptografada — formato "v1:iv:tag:ct" (AES-256-GCM portátil).
   *  Legado: base64 via safeStorage (Chromium/DPAPI — não portátil). */
  secret: string;
  secretEncrypted: boolean;
  /** Senha de SMTP separada — só presente quando difere da senha de entrada. */
  smtpSecret?: string;
  smtpSecretEncrypted?: boolean;
}

/**
 * Criptografia de senhas PORTÁTIL — AES-256-GCM com uma chave de 32 bytes
 * guardada em <raiz-de-dados>/secrets.key. A chave viaja junto da pasta de
 * dados, então "apontar o app para a pasta" em outro PC restaura as senhas
 * — é o mesmo modelo do perfil do Thunderbird (key4.db). O safeStorage do
 * Chromium amarra a chave a esta máquina (DPAPI), o que quebrava a portabilidade.
 */
function secretsKeyPath(): string {
  return join(dataDir(), 'secrets.key');
}

let cachedKey: Buffer | null = null;

function secretsKey(): Buffer {
  if (cachedKey) return cachedKey;
  const p = secretsKeyPath();
  if (existsSync(p)) {
    const k = readFileSync(p);
    if (k.length === 32) {
      cachedKey = k;
      return k;
    }
  }
  const k = randomBytes(32);
  mkdirSync(dataDir(), { recursive: true });
  writeFileSync(p, k, { mode: 0o600 });
  cachedKey = k;
  return k;
}

function encryptPortable(plain: string): string {
  const iv = randomBytes(12);
  const c = createCipheriv('aes-256-gcm', secretsKey(), iv);
  const ct = Buffer.concat([c.update(plain, 'utf-8'), c.final()]);
  return `v1:${iv.toString('base64')}:${c.getAuthTag().toString('base64')}:${ct.toString('base64')}`;
}

function decryptPortable(blob: string): string {
  const m = /^v1:([A-Za-z0-9+/=]+):([A-Za-z0-9+/=]+):([A-Za-z0-9+/=]+)$/.exec(blob);
  if (!m) return '';
  try {
    const d = createDecipheriv('aes-256-gcm', secretsKey(), Buffer.from(m[1], 'base64'));
    d.setAuthTag(Buffer.from(m[2], 'base64'));
    return Buffer.concat([d.update(Buffer.from(m[3], 'base64')), d.final()]).toString('utf-8');
  } catch {
    return '';
  }
}

function accountsPath(): string {
  return join(dataDir(), 'accounts.json');
}

function loadAll(): StoredAccount[] {
  try {
    const p = accountsPath();
    if (existsSync(p)) return JSON.parse(readFileSync(p, 'utf-8'));
  } catch {
    // arquivo corrompido — começa do zero
  }
  return [];
}

function saveAll(list: StoredAccount[]): void {
  mkdirSync(dataDir(), { recursive: true });
  writeFileSync(accountsPath(), JSON.stringify(list, null, 2));
}

/** As senhas sempre usam a cifra portátil — a informação de keyring do SO
 *  continua sendo exposta apenas como aviso informativo na UI. */
export function secureStorageAvailable(): boolean {
  return safeStorage.isEncryptionAvailable();
}

function encryptSecret(plain: string): { secret: string; secretEncrypted: boolean } {
  return { secret: encryptPortable(plain), secretEncrypted: true };
}

function decryptSecretValue(encrypted: boolean | undefined, value: string): string {
  try {
    if (value.startsWith('v1:')) return decryptPortable(value);
    if (encrypted && safeStorage.isEncryptionAvailable()) {
      return safeStorage.decryptString(Buffer.from(value, 'base64'));
    }
    return Buffer.from(value, 'base64').toString('utf-8');
  } catch {
    return '';
  }
}

/**
 * Regrava senhas antigas (safeStorage/DPAPI ou base64) no formato portátil —
 * roda uma vez no boot: depois disso a pasta de dados carrega TUDO, senhas
 * incluídas, para qualquer máquina.
 */
export function migrateSecrets(): void {
  const all = loadAll();
  let dirty = false;
  const fix = (a: StoredAccount, val: 'secret' | 'smtpSecret') => {
    const blob = a[val];
    if (!blob || blob.startsWith('v1:')) return;
    const enc = val === 'secret' ? a.secretEncrypted : a.smtpSecretEncrypted;
    const plain = decryptSecretValue(enc, blob);
    if (!plain) return; // blob ilegível aqui — mantém como está
    a[val] = encryptPortable(plain);
    if (val === 'secret') a.secretEncrypted = true;
    else a.smtpSecretEncrypted = true;
    dirty = true;
  };
  for (const a of all) {
    fix(a, 'secret');
    fix(a, 'smtpSecret');
  }
  if (dirty) saveAll(all);
}

function decryptSecret(a: StoredAccount): string {
  return decryptSecretValue(a.secretEncrypted, a.secret);
}

function stripSecret(a: StoredAccount): AccountConfig {
  const { secret, secretEncrypted, smtpSecret, smtpSecretEncrypted, ...rest } = a;
  void secretEncrypted;
  void smtpSecretEncrypted;
  return { ...rest, hasPassword: !!secret, hasSmtpPassword: !!smtpSecret };
}

export class AccountStore {
  list(): AccountConfig[] {
    return loadAll().map(stripSecret);
  }

  getPassword(accountId: string): string {
    const found = loadAll().find((a) => a.id === accountId);
    return found ? decryptSecret(found) : '';
  }

  /** Senha de SMTP — se a conta não tem uma separada, usa a senha de entrada. */
  getSmtpPassword(accountId: string): string {
    const found = loadAll().find((a) => a.id === accountId);
    if (!found) return '';
    return found.smtpSecret
      ? decryptSecretValue(found.smtpSecretEncrypted, found.smtpSecret)
      : decryptSecret(found);
  }

  getFull(accountId: string): AccountConfig | null {
    const found = loadAll().find((a) => a.id === accountId);
    return found ? stripSecret(found) : null;
  }

  upsert(draft: AccountDraft): AccountConfig {
    const all = loadAll();
    const id = draft.id || `acct-${randomUUID()}`;
    const { secret, secretEncrypted } = draft.password
      ? encryptSecret(draft.password)
      : { secret: '', secretEncrypted: false };
    const smtpSec = draft.smtpPassword
      ? encryptSecret(draft.smtpPassword)
      : { secret: '', secretEncrypted: false };
    const existingIdx = all.findIndex((a) => a.id === id);
    const isImap = (draft.protocol || 'imap') === 'imap';
    const base: StoredAccount = {
      id,
      label: draft.label,
      email: draft.email,
      protocol: draft.protocol || 'imap',
      signature: draft.signature || '',
      signatureNewId: draft.signatureNewId,
      signatureReplyId: draft.signatureReplyId,
      // IMAP não tem arquivo de dados — qualquer vínculo vindo de outra tela
      // (wizard, importação, backup) é ignorado aqui na gravação.
      deliveryPath: isImap ? undefined : draft.deliveryPath,
      deliveryFolder: isImap ? undefined : draft.deliveryFolder,
      leaveOnServer: draft.leaveOnServer ?? true,
      leaveDays: draft.leaveDays ?? 0,
      allowSelfSigned: draft.allowSelfSigned ?? false,
      imapHost: draft.imapHost,
      imapPort: draft.imapPort,
      imapSecure: draft.imapSecure,
      imapUser: draft.imapUser,
      smtpHost: draft.smtpHost,
      smtpPort: draft.smtpPort,
      smtpSecure: draft.smtpSecure,
      smtpUser: draft.smtpUser,
      secret: draft.password ? secret : existingIdx >= 0 ? all[existingIdx].secret : '',
      secretEncrypted: draft.password
        ? secretEncrypted
        : existingIdx >= 0
          ? all[existingIdx].secretEncrypted
          : false,
      smtpSecret: draft.smtpPassword
        ? smtpSec.secret
        : existingIdx >= 0
          ? all[existingIdx].smtpSecret
          : undefined,
      smtpSecretEncrypted: draft.smtpPassword
        ? smtpSec.secretEncrypted
        : existingIdx >= 0
          ? all[existingIdx].smtpSecretEncrypted
          : undefined,
    };
    if (existingIdx >= 0) all[existingIdx] = base;
    else all.push(base);
    saveAll(all);
    return stripSecret(base);
  }

  /** Atualiza campos de uma conta existente sem tocar na senha. */
  patch(accountId: string, fields: Partial<AccountConfig>): void {
    const all = loadAll();
    const idx = all.findIndex((a) => a.id === accountId);
    if (idx < 0) return;
    all[idx] = { ...all[idx], ...fields, id: accountId };
    saveAll(all);
  }

  remove(accountId: string): void {
    saveAll(loadAll().filter((a) => a.id !== accountId));
  }
}

export const accountStore = new AccountStore();
