import { execFile } from 'child_process';
import { promisify } from 'util';
import { existsSync, readdirSync, readFileSync } from 'fs';
import { join } from 'path';
import type { ImportedAccountGuess } from '../shared/types';

const execFileAsync = promisify(execFile);

// Chave fixa (GUID conhecido) onde o Outlook 2013/2016/2019/365 guarda a
// lista de contas de cada perfil — verificado contra o Registro real.
const ACCOUNTS_KEY = '9375cff0413111d3b88a00104b2a6676';
const PROFILE_ROOTS = [
  'HKCU\\Software\\Microsoft\\Office\\16.0\\Outlook\\Profiles',
  'HKCU\\Software\\Microsoft\\Office\\15.0\\Outlook\\Profiles',
  'HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows Messaging Subsystem\\Profiles',
];

interface RegValue {
  type: 'sz' | 'dword' | 'binary';
  data: string;
}

/** Seção = subchave de conta; mapa nome→valor */
type Section = Map<string, RegValue>;

async function regQuery(key: string, recursive: boolean): Promise<string> {
  try {
    const args = ['query', key];
    if (recursive) args.push('/s');
    const { stdout } = await execFileAsync('reg', args, { maxBuffer: 1024 * 1024 * 64 });
    return stdout;
  } catch {
    return '';
  }
}

/**
 * Parseia a saída de `reg query ... /s` agrupando por seção (linha HKEY_*).
 * Retorna subchave completa → mapa de valores nomeados.
 */
function parseRegSections(raw: string): Map<string, Section> {
  const sections = new Map<string, Section>();
  let current: Section | null = null;
  for (const line of raw.split(/\r?\n/)) {
    if (/^HKEY_/i.test(line.trim())) {
      current = new Map();
      sections.set(line.trim(), current);
      continue;
    }
    if (!current) continue;
    const m = line.match(/^\s+(.+?)\s{2,}(REG_SZ|REG_DWORD|REG_BINARY|REG_MULTI_SZ)\s{2,}(.*)$/i);
    if (!m) continue;
    const name = m[1].trim().toLowerCase();
    const type = m[2].toLowerCase();
    const data = m[3].trim();
    current.set(name, {
      type: type === 'reg_sz' || type === 'reg_multi_sz' ? 'sz' : type === 'reg_dword' ? 'dword' : 'binary',
      data,
    });
  }
  return sections;
}

function getStr(sec: Section, ...names: string[]): string {
  for (const n of names) {
    const v = sec.get(n);
    if (v && v.type === 'sz') return v.data;
  }
  return '';
}

function getBinary(sec: Section, ...names: string[]): string | undefined {
  for (const n of names) {
    const v = sec.get(n);
    if (v && v.type === 'binary') return v.data;
  }
  return undefined;
}

function getDword(sec: Section, ...names: string[]): number | null {
  for (const n of names) {
    const v = sec.get(n);
    if (v && v.type === 'dword') {
      const n0 = v.data.startsWith('0x') ? parseInt(v.data, 16) : parseInt(v.data, 10);
      if (!isNaN(n0)) return n0;
    }
  }
  return null;
}

/** Extrai o caminho do PST embutido em UTF-16LE no Delivery Store EntryID. */
function pstPathFromBinary(hex: string): string | undefined {
  try {
    const buf = Buffer.from(hex.replace(/\s+/g, ''), 'hex');
    const text = buf.toString('utf16le');
    const m = text.match(/[A-Za-z]:\\[^\x00]+?\.(?:pst|ost)/i);
    return m ? m[0] : undefined;
  } catch {
    return undefined;
  }
}

/** Extrai o conteúdo do <body> de um HTML gerado pelo Word (assinatura do Outlook). */
export function extractBody(html: string): string {
  const m = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  const inner = m ? m[1] : html;
  // Remove tags condicionais do Office e comentários MSO
  return inner
    .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, '')
    .replace(/<o:p>\s*<\/o:p>/gi, ' ')
    .trim();
}

/**
 * Decripta uma senha que o Outlook guarda no Registro protegida por DPAPI
 * (CryptProtectData) — o mesmo mecanismo que o próprio Outlook usa para
 * mostrar a senha no botão "Reparar conta" (ícone de olho). Só funciona
 * rodando com o MESMO usuário do Windows que salvou a senha: é o usuário
 * acessando a própria credencial, não quebra nem contorna proteção alheia.
 */
async function decryptDpapi(hexBytes: string): Promise<string | undefined> {
  try {
    let clean = hexBytes.replace(/\s+/g, '');
    // O valor do Registro pode ter um byte de prefixo (0x02) antes do blob
    // DPAPI real — localiza o cabeçalho do blob: versão 01 00 00 00 + GUID do
    // provedor DPAPI {D08C9DDF-0115-D111-8C7A-00C04FC297EB}.
    const marker = clean.toLowerCase().indexOf('01000000d08c9ddf0115d1118c7a00c04fc297eb');
    if (marker >= 0) clean = clean.slice(marker);
    if (!clean || clean.length % 2 !== 0) return undefined;
    const b64 = Buffer.from(clean, 'hex').toString('base64');
    const script =
      "$b=[Convert]::FromBase64String('" + b64 + "');" +
      "Add-Type -AssemblyName System.Security;" +
      "$d=[System.Security.Cryptography.ProtectedData]::Unprotect($b,$null,[System.Security.Cryptography.DataProtectionScope]::CurrentUser);" +
      '[Convert]::ToBase64String($d)';
    const { stdout } = await execFileAsync(
      'powershell',
      ['-NoProfile', '-NonInteractive', '-Command', script],
      { maxBuffer: 1024 * 1024, windowsHide: true },
    );
    const decoded = Buffer.from(stdout.trim(), 'base64');
    // Outlook grava a senha como UTF-16LE terminada em nulo. Algumas versões
    // gravam ANSI/UTF-8 — escolhe a decodificação que produz mais caracteres
    // imprimíveis, para não salvar lixo que vira "senha errada" no servidor.
    // NÃO remover caracteres das bordas: uma senha que começa/termina em '*'
    // é válida — cortar os asteriscos gera uma senha errada.
    const printable = (s: string) =>
      (s.match(/[\x20-\x7eÀ-ž]/g)?.length ?? 0) / Math.max(1, s.length);
    const u16 = decoded.toString('utf16le').replace(/\0+$/, '');
    const u8 = decoded.toString('utf-8').replace(/\0+$/, '');
    const text = printable(u16) >= printable(u8) ? u16 : u8;
    return text || undefined;
  } catch {
    // não-DPAPI, usuário diferente, ou PowerShell indisponível
    return undefined;
  }
}

/** Lê o arquivo de assinatura do Outlook (%APPDATA%\Microsoft\Signatures). */
function readSignatureHtml(sigName: string): string | undefined {
  if (!sigName || !process.env.APPDATA) return undefined;
  const dir = join(process.env.APPDATA, 'Microsoft', 'Signatures');
  try {
    for (const ext of ['.htm', '.html', '.txt']) {
      const p = join(dir, sigName + ext);
      if (existsSync(p)) {
        const raw = readFileSync(p, 'utf-8');
        return ext === '.txt' ? raw.replace(/\n/g, '<br>') : extractBody(raw);
      }
    }
    // fallback: primeiro .htm com nome parecido
    const similar = readdirSync(dir).find((f) => f.toLowerCase().includes(sigName.toLowerCase()) && f.endsWith('.htm'));
    if (similar) return extractBody(readFileSync(join(dir, similar), 'utf-8'));
  } catch {
    // sem assinatura — ignora
  }
  return undefined;
}

export async function isWindowsImportAvailable(): Promise<boolean> {
  return process.platform === 'win32';
}

export async function importOutlookAccountsFromRegistry(): Promise<ImportedAccountGuess[]> {
  if (process.platform !== 'win32') return [];

  const guesses: ImportedAccountGuess[] = [];

  for (const root of PROFILE_ROOTS) {
    // Enumera perfis (ex.: "Outlook", "Meu Perfil"...)
    const profilesRaw = await regQuery(root, false);
    const profiles = profilesRaw
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => /^HKEY_/i.test(l))
      .map((l) => l.slice(l.indexOf('Profiles\\') + 'Profiles\\'.length));

    for (const profile of profiles) {
      const accountsKey = `${root}\\${profile}\\${ACCOUNTS_KEY}`;
      const raw = await regQuery(accountsKey, true);
      if (!raw) continue;
      const sections = parseRegSections(raw);

      for (const [, sec] of sections) {
        const email = getStr(sec, 'email');
        const pop3Server = getStr(sec, 'pop3 server');
        const imapServer = getStr(sec, 'imap server');
        const smtpServer = getStr(sec, 'smtp server');
        // Pula seções que não são contas de email (CONTAB, MSUPST etc.)
        if (!email && !pop3Server && !imapServer && !smtpServer) continue;
        // Requer ao menos um servidor identificável
        if (!pop3Server && !imapServer && !smtpServer) continue;

        const protocol: 'imap' | 'pop3' = pop3Server ? 'pop3' : 'imap';
        const inServer = protocol === 'pop3' ? pop3Server : imapServer;
        const inUser =
          getStr(sec, 'pop3 user', 'pop3 user name') ||
          getStr(sec, 'imap user', 'imap user name') ||
          email;
        // O Outlook só grava "POP3/IMAP Use SSL" e "POP3/IMAP Port" quando
        // configurados — ausentes significam conexão sem SSL na porta padrão
        // (110/143), NÃO 995/993.
        const inSsl =
          (getDword(sec, 'pop3 use ssl', 'imap use ssl', 'pop3 secure connection', 'imap secure connection') ?? 0) !== 0;
        const inPort =
          getDword(sec, 'pop3 port', 'imap port') ??
          (inSsl ? (protocol === 'pop3' ? 995 : 993) : protocol === 'pop3' ? 110 : 143);
        // "SMTP Secure Connection": 0=nenhum, 1=SSL implícito (465),
        // 2=STARTTLS, 3=auto. Só o valor 1 usa TLS na conexão (nodemailer
        // secure=true); STARTTLS é upgrade posterior → secure=false.
        const smtpPort = getDword(sec, 'smtp port') ?? 587;
        const smtpConn = getDword(sec, 'smtp secure connection');
        const smtpSsl =
          smtpConn === 1 ||
          (smtpConn === null &&
            (getDword(sec, 'smtp use ssl') ?? 0) !== 0 &&
            smtpPort === 465);
        // "Leave on Server" é um bitmask: bit0 = deixar cópia no servidor,
        // bit1 = remover após N dias, bits 16-23 = N.
        const leaveBits = getDword(sec, 'leave on server');
        const leaveOnServer = leaveBits === null ? true : (leaveBits & 1) !== 0;
        const leaveDays =
          leaveBits !== null && (leaveBits & 2) !== 0 ? (leaveBits >> 16) & 0xff : 0;

        const delivery = sec.get('delivery store entryid');
        const pstPath = delivery?.type === 'binary' ? pstPathFromBinary(delivery.data) : undefined;

        const sigName = getStr(sec, 'new signature');
        const signatureHtml = readSignatureHtml(sigName);
        const replySigName = getStr(sec, 'reply-forward signature');
        const signatureReplyHtml = replySigName ? readSignatureHtml(replySigName) : undefined;

        // Senha de entrada (POP3/IMAP) gravada pelo Outlook protegida por
        // DPAPI — decripta só porque estamos rodando como o mesmo usuário
        // do Windows que a salvou (mesmo mecanismo do "Reparar conta").
        const passBinary = getBinary(sec, 'pop3 password', 'imap password', 'http password');
        const smtpPassBinary = getBinary(sec, 'smtp password');
        const password = passBinary ? await decryptDpapi(passBinary) : undefined;
        const smtpPassword = smtpPassBinary ? await decryptDpapi(smtpPassBinary) : undefined;
        // O usuário de entrada/saída nem sempre é o e-mail — sem ele o login
        // falha mesmo com a senha certa.
        const smtpUser = getStr(sec, 'smtp user', 'smtp user name', 'smtp email');

        guesses.push({
          email: email || inUser,
          displayName: getStr(sec, 'display name', 'account name') || email || inUser,
          protocol,
          imapUser: inUser,
          smtpUser: smtpUser || inUser,
          imapHost: inServer,
          imapPort: inPort,
          imapSecure: inSsl,
          smtpHost: smtpServer || inServer,
          smtpPort,
          smtpSecure: smtpSsl,
          pstPath,
          signatureHtml,
          signatureReplyHtml,
          password,
          smtpPassword,
          leaveOnServer,
          leaveDays,
        });
      }
    }
  }

  // Remove duplicatas por email (mesma conta pode aparecer em mais de um perfil)
  const seen = new Set<string>();
  return guesses.filter((g) => {
    const key = `${g.email}|${g.imapHost}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
