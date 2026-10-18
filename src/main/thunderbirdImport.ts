import { createReadStream, existsSync, readdirSync, readFileSync, statSync, openSync, readSync, closeSync } from 'fs';
import { homedir } from 'os';
import { join, isAbsolute } from 'path';
import type { ImportedAccountGuess, TbMailFolder, TbMailProfile, TbMailStore } from '../shared/types';
import { extractBody } from './outlookImportWin';

/**
 * Importa contas do Thunderbird lendo os perfis em texto puro (profiles.ini e
 * prefs.js). Funciona em Linux, Windows e macOS.
 *
 * As senhas NÃO são recuperadas: o Thunderbird as guarda cifradas em
 * logins.json/key4.db (NSS). O usuário digita a senha ao usar a conta.
 */

type Prefs = Map<string, string | number | boolean>;

/** Raízes onde o Thunderbird guarda perfis (instalação padrão, Flatpak e Snap). */
function thunderbirdRoots(): string[] {
  const home = homedir();
  if (process.platform === 'win32') {
    return [join(process.env.APPDATA || join(home, 'AppData', 'Roaming'), 'Thunderbird')];
  }
  if (process.platform === 'darwin') return [join(home, 'Library', 'Thunderbird')];
  return [
    join(home, '.thunderbird'),
    join(home, '.var', 'app', 'org.mozilla.Thunderbird', '.thunderbird'),
    join(home, 'snap', 'thunderbird', 'common', '.thunderbird'),
  ];
}

/** Lê profiles.ini e devolve os perfis (diretório + nome) — ou varre subpastas com prefs.js. */
function profileInfos(root: string): Array<{ path: string; name: string }> {
  const dirs = new Map<string, string>();
  const ini = join(root, 'profiles.ini');
  if (existsSync(ini)) {
    let section = '';
    let path = '';
    let name = '';
    let relative = true;
    const flush = () => {
      if (section.startsWith('Profile') && path) {
        const p = relative ? join(root, path) : path;
        if (!dirs.has(p)) dirs.set(p, name);
      }
      path = '';
      name = '';
      relative = true;
    };
    for (const raw of readFileSync(ini, 'utf-8').split(/\r?\n/)) {
      const line = raw.trim();
      if (line.startsWith('[')) {
        flush();
        section = line.slice(1, -1);
      } else if (/^Path=/i.test(line)) path = line.slice(5);
      else if (/^Name=/i.test(line)) name = line.slice(5);
      else if (/^IsRelative=/i.test(line)) relative = line.slice(11) !== '0';
    }
    flush();
  }
  if (!dirs.size && existsSync(root)) {
    for (const f of readdirSync(root, { withFileTypes: true })) {
      const p = join(root, f.name);
      if (f.isDirectory() && existsSync(join(p, 'prefs.js'))) dirs.set(p, f.name);
    }
  }
  return [...dirs]
    .filter(([d]) => existsSync(join(d, 'prefs.js')))
    .map(([p, n]) => ({ path: p, name: n || p.split(/[\\/]/).pop() || p }));
}

function profileDirs(root: string): string[] {
  return profileInfos(root).map((p) => p.path);
}

/** Converte valores do prefs.js (string entre aspas, número ou booleano). */
function parseValue(v: string): string | number | boolean {
  if (v === 'true') return true;
  if (v === 'false') return false;
  if (v.startsWith('"')) {
    try {
      return JSON.parse(v);
    } catch {
      return v.slice(1, -1);
    }
  }
  const n = Number(v);
  return Number.isNaN(n) ? v : n;
}

export function parsePrefs(text: string): Prefs {
  const prefs: Prefs = new Map();
  const re = /^\s*user_pref\("([^"]+)",\s*(.+?)\);\s*$/;
  for (const line of text.split(/\r?\n/)) {
    const m = line.match(re);
    if (m) prefs.set(m[1], parseValue(m[2]));
  }
  return prefs;
}

const str = (p: Prefs, k: string): string => {
  const v = p.get(k);
  return typeof v === 'string' ? v : '';
};
const num = (p: Prefs, k: string): number | undefined => {
  const v = p.get(k);
  return typeof v === 'number' ? v : undefined;
};
const bool = (p: Prefs, k: string): boolean | undefined => {
  const v = p.get(k);
  return typeof v === 'boolean' ? v : undefined;
};
const list = (p: Prefs, k: string): string[] =>
  str(p, k)
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);

/** Assinatura: arquivo texto ou HTML apontado pela identidade (sig_file). */
function readSignature(path: string): string | undefined {
  if (!path) return undefined;
  const file = isAbsolute(path) ? path : '';
  if (!file || !existsSync(file)) return undefined;
  try {
    const raw = readFileSync(file, 'utf-8');
    if (/\.html?$/i.test(file) || /<html|<body/i.test(raw)) return extractBody(raw);
    return raw.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r?\n/g, '<br>');
  } catch {
    return undefined;
  }
}

/** socketType do Thunderbird: 0 = nenhuma, 2 = STARTTLS, 3 = SSL/TLS. */
const usesSsl = (socketType: number | undefined) => socketType === 3;

function portFor(configured: number | undefined, protocol: 'imap' | 'pop3', ssl: boolean): number {
  if (configured && configured > 0) return configured;
  if (protocol === 'pop3') return ssl ? 995 : 110;
  return ssl ? 993 : 143;
}

export function accountsFromPrefs(p: Prefs): ImportedAccountGuess[] {
  const out: ImportedAccountGuess[] = [];

  for (const acc of list(p, 'mail.accountmanager.accounts')) {
    const serverKey = str(p, `mail.account.${acc}.server`);
    const type = str(p, `mail.server.${serverKey}.type`);
    if (!serverKey || (type !== 'pop3' && type !== 'imap')) continue; // Pastas Locais, RSS, NNTP...

    const identityKey = list(p, `mail.account.${acc}.identities`)[0];
    const email = identityKey ? str(p, `mail.identity.${identityKey}.useremail`) : '';
    if (!email) continue;

    const protocol: 'imap' | 'pop3' = type;
    const inSocket = num(p, `mail.server.${serverKey}.socketType`);
    const inSsl = usesSsl(inSocket);
    const smtpKey =
      (identityKey && str(p, `mail.identity.${identityKey}.smtpServer`)) ||
      str(p, 'mail.smtp.defaultserver');
    const smtpHost = smtpKey ? str(p, `mail.smtpserver.${smtpKey}.hostname`) : '';
    const smtpSocket = smtpKey ? num(p, `mail.smtpserver.${smtpKey}.socketType`) : undefined;
    const smtpSsl = usesSsl(smtpSocket);
    const inHost = str(p, `mail.server.${serverKey}.hostname`);
    const leave = bool(p, `mail.server.${serverKey}.leave_on_server`);
    const deleteAfter = bool(p, `mail.server.${serverKey}.delete_mail_left_on_server`);

    out.push({
      email,
      displayName: str(p, `mail.identity.${identityKey}.fullName`) || email,
      protocol,
      imapHost: inHost,
      imapPort: portFor(num(p, `mail.server.${serverKey}.port`), protocol, inSsl),
      imapSecure: inSsl,
      imapUser: str(p, `mail.server.${serverKey}.userName`) || undefined,
      smtpHost: smtpHost || inHost,
      smtpUser: smtpKey ? str(p, `mail.smtpserver.${smtpKey}.username`) || undefined : undefined,
      smtpPort: smtpKey ? num(p, `mail.smtpserver.${smtpKey}.port`) || (smtpSsl ? 465 : 587) : 587,
      smtpSecure: smtpSsl,
      signatureHtml: identityKey ? readSignature(str(p, `mail.identity.${identityKey}.sig_file`)) : undefined,
      leaveOnServer: leave ?? true,
      leaveDays: deleteAfter ? num(p, `mail.server.${serverKey}.num_days_to_leave_on_server`) || 0 : 0,
    });
  }
  return out;
}

/** Procura as contas do Thunderbird em todos os perfis encontrados neste sistema. */
export async function scanThunderbirdAccounts(): Promise<ImportedAccountGuess[]> {
  const guesses: ImportedAccountGuess[] = [];
  for (const root of thunderbirdRoots()) {
    if (!existsSync(root)) continue;
    for (const dir of profileDirs(root)) {
      try {
        guesses.push(...accountsFromPrefs(parsePrefs(readFileSync(join(dir, 'prefs.js'), 'utf-8'))));
      } catch {
        // perfil ilegível ou em uso — ignora e segue para o próximo
      }
    }
  }
  const seen = new Set<string>();
  return guesses.filter((g) => {
    const key = `${g.email.toLowerCase()}|${g.imapHost}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// ---------------------------------------------------------------------------
// Importação de MENSAGENS — o Thunderbird guarda os emails em arquivos .mbox
// dentro de <perfil>/Mail/ (POP + Pastas Locais) e <perfil>/ImapMail/ (IMAP).
// Cada pasta de email é um arquivo sem extensão acompanhado de um índice .msf;
// subpastas ficam num diretório "<nome>.sbd".
// ---------------------------------------------------------------------------

/** True se o arquivo parece ser um mbox: sem extensão, com .msf irmão ou começando por "From ". */
function isMboxFile(dir: string, name: string): boolean {
  if (name.includes('.')) return false;
  const p = join(dir, name);
  try {
    if (!statSync(p).isFile()) return false;
    if (existsSync(p + '.msf')) return true;
    if (statSync(p).size === 0) return false;
    const fd = openSync(p, 'r');
    try {
      const buf = Buffer.alloc(5);
      return readSync(fd, buf, 0, 5, 0) === 5 && buf.toString('latin1') === 'From ';
    } finally {
      closeSync(fd);
    }
  } catch {
    return false;
  }
}

/** Árvore de pastas de um armazenamento (dir de Mail/ImapMail) — recusa via .sbd. */
function mboxTree(dir: string): TbMailFolder[] {
  let ents;
  try {
    ents = readdirSync(dir, { withFileTypes: true });
  } catch {
    return [];
  }
  const sbd = new Map<string, string>();
  for (const e of ents) {
    if (e.isDirectory() && e.name.endsWith('.sbd')) sbd.set(e.name.slice(0, -4), join(dir, e.name));
  }
  const out: TbMailFolder[] = [];
  for (const e of ents) {
    if (!e.isFile() || !isMboxFile(dir, e.name)) continue;
    const kidsDir = sbd.get(e.name);
    out.push({
      name: e.name,
      mbox: join(dir, e.name),
      children: kidsDir ? mboxTree(kidsDir) : [],
    });
  }
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Procura perfis do Thunderbird neste sistema e devolve, para cada um, os
 * armazenamentos de email (Pastas Locais, contas POP e IMAP) com a árvore
 * de arquivos .mbox — usado pelo diálogo "Importar do Thunderbird".
 */
export function scanThunderbirdMail(): TbMailProfile[] {
  const profiles: TbMailProfile[] = [];
  for (const root of thunderbirdRoots()) {
    if (!existsSync(root)) continue;
    for (const prof of profileInfos(root)) {
      const stores: TbMailStore[] = [];
      const addStores = (base: string, imap: boolean) => {
        let subs;
        try {
          subs = readdirSync(base, { withFileTypes: true });
        } catch {
          return;
        }
        for (const e of subs) {
          if (!e.isDirectory() || e.name === 'smart mailboxes') continue;
          const folders = mboxTree(join(base, e.name));
          if (!folders.length) continue;
          const label =
            e.name === 'Local Folders' ? 'Pastas Locais' : imap ? `${e.name} (IMAP)` : e.name;
          stores.push({ label, folders });
        }
      };
      addStores(join(prof.path, 'Mail'), false);
      addStores(join(prof.path, 'ImapMail'), true);
      if (stores.length) profiles.push({ path: prof.path, name: prof.name, stores });
    }
  }
  return profiles;
}

/** Conta as mensagens de um mbox (linhas "^From ") numa passada em stream. */
export async function countMboxMessages(path: string): Promise<number> {
  let count = 0;
  let first = true;
  let tail = Buffer.alloc(0);
  for await (const chunk of createReadStream(path)) {
    const buf = tail.length ? Buffer.concat([tail, chunk]) : (chunk as Buffer);
    if (first) {
      first = false;
      if (buf.length >= 5 && buf.subarray(0, 5).toString('latin1') === 'From ') count++;
    }
    let idx = 0;
    while ((idx = buf.indexOf('\nFrom ', idx)) !== -1) {
      count++;
      idx += 6;
    }
    tail = Buffer.from(buf.subarray(Math.max(0, buf.length - 5)));
  }
  return count;
}

const FROM_LINE = Buffer.from('From ', 'latin1');

/**
 * Gera as mensagens de um mbox uma a uma, em stream — cada item é o corpo
 * bruto da mensagem (sem a linha separadora "From "). Não carrega o arquivo
 * inteiro em memória, então funciona com mboxes de vários GB.
 */
export async function* mboxMessages(path: string): AsyncGenerator<Buffer> {
  let leftover = Buffer.alloc(0);
  let msg: Buffer[] | null = null;
  for await (const chunk of createReadStream(path)) {
    const data = leftover.length ? Buffer.concat([leftover, chunk]) : (chunk as Buffer);
    let start = 0;
    for (;;) {
      const nl = data.indexOf(0x0a, start); // '\n'
      if (nl === -1) break;
      const line = data.subarray(start, nl + 1);
      if (line.length >= 5 && line.subarray(0, 5).equals(FROM_LINE)) {
        if (msg) yield Buffer.concat(msg);
        msg = [];
      } else if (msg) {
        msg.push(Buffer.from(line));
      }
      start = nl + 1;
    }
    leftover = Buffer.from(data.subarray(start));
  }
  if (leftover.length && msg) msg.push(leftover);
  if (msg) yield Buffer.concat(msg);
}

/**
 * Converte uma mensagem mbox em .eml: remove os cabeçalhos internos do
 * Thunderbird (X-Mozilla-*) e desescapa linhas ">From ". Devolve também a
 * flag de lida gravada no X-Mozilla-Status (bit 0x0001).
 */
export function mboxMessageToEml(raw: Buffer): { eml: Buffer; read: boolean } | null {
  if (!raw.length) return null;
  const s = raw.toString('latin1');
  const headEnd = s.search(/\r?\n\r?\n/);
  const head = headEnd === -1 ? s : s.slice(0, headEnd);
  const m = head.match(/^X-Mozilla-Status:\s*([0-9a-fA-F]+)/m);
  const read = m ? (parseInt(m[1], 16) & 1) !== 0 : false;
  const eml = s
    .replace(/^X-Mozilla-(?:Status|Status2|Keys|Draft-Info):[^\n]*(?:\r?\n[ \t][^\n]*)*\r?\n/gm, '')
    .replace(/^>(>*From )/gm, '$1');
  return { eml: Buffer.from(eml, 'latin1'), read };
}
