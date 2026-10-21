import nodemailer from 'nodemailer';
import { basename } from 'path';
import { signatureStore } from './signatureStore';
import type { AccountConfig, ComposeInput } from '../shared/types';

/** Resolve a assinatura de uma conta conforme o tipo de mensagem. */
export function resolveSignatureHtml(account: AccountConfig, kind?: ComposeInput['kind']): string {
  const sigId =
    kind === 'reply' || kind === 'forward' ? account.signatureReplyId : account.signatureNewId;
  if (sigId === '__none') return '';
  return (sigId ? signatureStore.get(sigId)?.html : undefined) ?? account.signature ?? '';
}

function buildMailOptions(account: AccountConfig, input: ComposeInput) {
  // Se o editor já embutiu a assinatura no corpo (bloco data-sig), não anexa de novo
  const bodyHasSig = input.html.includes('data-sig');
  const sig = input.noSignature || bodyHasSig ? '' : resolveSignatureHtml(account, input.kind);
  const signatureHtml = sig ? `<br><br>${sig}` : '';
  return {
    from: `"${account.label || account.email}" <${account.email}>`,
    to: input.to,
    cc: input.cc || undefined,
    bcc: input.bcc || undefined,
    subject: input.subject,
    html: input.html + signatureHtml,
    // Rascunho — mesma marca que o Outlook grava nos itens de Rascunhos
    headers: input.draft ? { 'X-Unsent': '1' } : undefined,
    inReplyTo: input.inReplyTo || undefined,
    references: input.references || undefined,
    // iMIP: nodemailer coloca o VCALENDAR como alternativa text/calendar —
    // Outlook/Thunderbird/Gmail mostram botões Aceitar/Recusar nativos.
    icalEvent: input.icalEvent
      ? {
          method: input.icalEvent.method,
          filename: 'invite.ics',
          content: Buffer.from(input.icalEvent.contentBase64, 'base64'),
        }
      : undefined,
    attachments: [
      ...(input.attachmentPaths || []).map((p) => ({
        filename: basename(p),
        path: p,
      })),
      ...(input.rawAttachments || []).map((a) => ({
        filename: a.filename,
        content: Buffer.from(a.contentBase64, 'base64'),
        contentType: a.contentType,
      })),
    ],
  };
}

/** Gera a mensagem MIME crua (para salvar cópia em Itens Enviados/Caixa de Saída). */
export function composeRawMessage(account: AccountConfig, input: ComposeInput): Promise<Buffer> {
  // nodemailer embute o mailcomposer — mesma engine usada no envio real
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const MailComposer = require('nodemailer/lib/mail-composer');
  return new Promise((resolve, reject) => {
    new MailComposer(buildMailOptions(account, input)).compile().build((e: Error | null, msg: Buffer) =>
      e ? reject(e) : resolve(msg),
    );
  });
}

export async function sendMail(
  account: AccountConfig,
  password: string,
  input: ComposeInput,
  /** Senha de SMTP separada — quando omitida, usa a senha de entrada. */
  smtpPassword?: string,
): Promise<void> {
  const transporter = nodemailer.createTransport({
    host: account.smtpHost,
    port: account.smtpPort,
    secure: account.smtpSecure,
    auth: { user: account.smtpUser || account.email, pass: smtpPassword || password },
    // Servidor com certificado autoassinado (opção da conta)
    tls: account.allowSelfSigned ? { rejectUnauthorized: false } : undefined,
  });

  await transporter.sendMail(buildMailOptions(account, input));
}

// ---------------------------------------------------------------------------
// Diagnóstico de conexão — traduz erros crus das libs (imapflow/nodemailer/
// sockets) em causa provável + dica de correção, identificando a etapa.
// ---------------------------------------------------------------------------

const CERT_RE =
  /SELF_SIGNED|UNABLE_TO_VERIFY|CERT_HAS_EXPIRED|CERT_ALTNAME|UNABLE_TO_GET_ISSUER|HOSTNAME_MISMATCH|certificate/i;
const AUTH_RE =
  /auth(entication|enticate)?\s*(failed|error)|invalid credentials|login disabled|incorrect (password|senha)|senha (incorreta|inv[áa]lida)|bad credentials|username or password|denied/i;
const TLS_NEG_RE =
  /wrong version number|ssl3_|tlsv1 alert|SSL routines|packet length too long|unexpected.*(byte|data|message)|EPROTO/i;

function portHint(stage: string): string {
  if (stage === 'SMTP') return 'SMTP: 465 = SSL, 587/25 = STARTTLS (sem SSL)';
  if (stage === 'POP3') return 'POP3: 995 = SSL, 110 = sem criptografia';
  return 'IMAP: 993 = SSL, 143 = STARTTLS (sem SSL)';
}

/** Monta mensagem "ETAPA: causa provável — dica". */
export function explainConnError(
  e: unknown,
  stage: 'IMAP' | 'POP3' | 'SMTP',
  ctx: { host?: string; port?: number; secure?: boolean; user?: string },
): string {
  const err = (e || {}) as {
    message?: string;
    code?: string;
    responseCode?: number;
    responseText?: string;
    serverResponseCode?: string;
    authenticationFailed?: boolean;
  };
  const msg = String(err.message || e || '');
  const blob = `${err.code || ''} ${err.serverResponseCode || ''} ${err.responseText || ''} ${msg}`;
  const where = `${ctx.host || '?'}:${ctx.port || '?'}`;

  if (
    err.authenticationFailed === true ||
    err.code === 'EAUTH' ||
    /^(535|534|530)\b/.test(String(err.responseCode || '')) ||
    AUTH_RE.test(blob)
  )
    return (
      `${stage}: usuário ou senha recusados pelo servidor` +
      (ctx.user ? ` (tentou "${ctx.user}")` : '') +
      '. Confira usuário e senha — em Gmail e Microsoft 365 a senha normal não funciona, é preciso gerar uma "senha de app" no site do provedor.'
    );
  if (/ENOTFOUND|EAI_AGAIN|getaddrinfo/i.test(blob))
    return `${stage}: servidor "${ctx.host || '?'}" não encontrado — o nome do servidor está errado ou sem internet.`;
  if (CERT_RE.test(err.code || '') || CERT_RE.test(msg))
    return `${stage}: certificado TLS do servidor não é confiável (${err.code || msg}) — normalmente o nome do servidor está errado; se for autoassinado, confirme com o provedor.`;
  if (TLS_NEG_RE.test(blob))
    return `${stage}: falha na negociação TLS em ${where} — a criptografia está ${ctx.secure ? 'ativada numa porta sem SSL' : 'desativada numa porta que exige SSL'} (${portHint(stage)}).`;
  if (/ETIMEDOUT|ETIMEOUT|timed?\s*out/i.test(blob))
    return `${stage}: sem resposta de ${where} — porta/criptografia incorretas (${portHint(stage)}) ou firewall bloqueando.`;
  if (/ECONNREFUSED/i.test(blob))
    return `${stage}: conexão recusada por ${where} — porta provavelmente errada (${portHint(stage)}) ou serviço fora do ar.`;
  if (/ECONNRESET|EPIPE|ENETUNREACH|EHOSTUNREACH|ECONNABORTED/i.test(blob))
    return `${stage}: a conexão caiu em ${where} (${err.code || msg}) — rede/firewall ou criptografia incompatível (${portHint(stage)}).`;
  return `${stage}: ${err.responseText || msg || 'erro desconhecido'}`;
}

/** Erros em que vale tentar de novo com a criptografia invertida (porta↔SSL). */
function retriableConnError(e: unknown): boolean {
  const err = (e || {}) as { code?: string; message?: string; serverResponseCode?: string; authenticationFailed?: boolean };
  const blob = `${err.code || ''} ${err.serverResponseCode || ''} ${err.message || ''}`;
  return !(
    err.authenticationFailed === true ||
    AUTH_RE.test(blob) ||
    /ENOTFOUND|EAI_AGAIN|getaddrinfo|ECONNREFUSED/i.test(blob)
  );
}

export interface TestAccountResult {
  message: string;
  /** Quando o teste descobriu que a criptografia do servidor de entrada está
   *  invertida — o formulário deve aplicar este valor em `imapSecure`. */
  imapSecureFix?: boolean;
  /** O servidor usa certificado autoassinado — só conectou aceitando-o;
   *  o formulário deve marcar `allowSelfSigned`. */
  allowSelfSignedFix?: boolean;
}

/** Erro de certificado TLS (autoassinado/cadeia não confiável). */
function isCertError(e: unknown): boolean {
  const err = (e || {}) as { code?: string; message?: string };
  return CERT_RE.test(err.code || '') || CERT_RE.test(err.message || '');
}

export async function testAccount(
  account: AccountConfig,
  password: string,
  smtpPassword?: string,
): Promise<TestAccountResult> {
  const inCtx = {
    host: account.imapHost,
    port: account.imapPort,
    secure: account.imapSecure,
    user: account.imapUser || account.email,
  };
  const outCtx = {
    host: account.smtpHost,
    port: account.smtpPort,
    secure: account.smtpSecure,
    user: account.smtpUser || account.email,
  };

  // ---- Servidor de entrada (POP3 ou IMAP) ----
  let inMsg = '';
  let secureFix: boolean | undefined;
  let selfSignedFix: boolean | undefined;

  const pop3Try = async (secure: boolean, allowSS = !!account.allowSelfSigned): Promise<string> => {
    const { Pop3Client } = await import('./pop3Client');
    const c = new Pop3Client();
    await c.connect(account.imapHost, account.imapPort, secure, inCtx.user, password, allowSS);
    const list = await c.uidl();
    await c.quit();
    const { pop3SeenUids } = await import('./ruleEngine');
    const seen = pop3SeenUids(account);
    const fresh = list.filter((m) => !seen.has(m.uidl)).length;
    return fresh > 0
      ? `POP3 OK — ${list.length} mensagem(ns) no servidor, ${fresh} nova(s) para baixar`
      : `POP3 OK — ${list.length} mensagem(ns) no servidor, todas já baixadas`;
  };

  const imapTry = async (secure: boolean, allowSS = !!account.allowSelfSigned): Promise<string> => {
    const { ImapFlow } = await import('imapflow');
    const client = new ImapFlow({
      host: account.imapHost,
      port: account.imapPort,
      secure,
      auth: { user: inCtx.user, pass: password },
      logger: false,
      disableAutoIdle: true,
      socketTimeout: 20000,
      greetingTimeout: 15000,
      tls: allowSS ? { rejectUnauthorized: false } : undefined,
    });
    await client.connect();
    let exists = -1;
    try {
      const box = await client.mailboxOpen('INBOX');
      exists = box.exists ?? -1;
    } catch {
      // INBOX inacessível — conexão já provou que funciona
    }
    await client.logout();
    return exists >= 0 ? `IMAP OK — caixa de entrada com ${exists} mensagem(ns)` : 'IMAP OK';
  };

  const proto = account.protocol === 'pop3' ? 'POP3' : 'IMAP';
  const tryIn = account.protocol === 'pop3' ? pop3Try : imapTry;
  try {
    inMsg = await tryIn(!!account.imapSecure);
  } catch (e) {
    if (isCertError(e)) {
      // Certificado autoassinado/não confiável — tenta aceitando-o (servidores
      // próprios e hospedagens comuns). O formulário ganha a opção marcada.
      try {
        inMsg = await tryIn(!!account.imapSecure, true);
        selfSignedFix = true;
      } catch {
        throw new Error(explainConnError(e, proto, inCtx));
      }
    } else if (retriableConnError(e)) {
      // Falha "de transporte" → tenta uma vez com a criptografia invertida;
      // se funcionar, avisa exatamente qual opção virar (causa nº1 de "IMAP
      // não funciona": porta 993 sem SSL marcado, ou 143 com SSL marcado).
      try {
        inMsg = await tryIn(!account.imapSecure);
        secureFix = !account.imapSecure;
      } catch (e2) {
        // Porta invertida + certificado autoassinado — combinação comum.
        if (isCertError(e2)) {
          try {
            inMsg = await tryIn(!account.imapSecure, true);
            secureFix = !account.imapSecure;
            selfSignedFix = true;
          } catch {
            throw new Error(explainConnError(e, proto, inCtx));
          }
        } else {
          throw new Error(explainConnError(e, proto, inCtx));
        }
      }
    } else {
      throw new Error(explainConnError(e, proto, inCtx));
    }
  }

  // ---- Servidor de saída (SMTP) — testado para POP3 e IMAP ----
  const smtpTry = async (allowSS: boolean) => {
    const transporter = nodemailer.createTransport({
      host: account.smtpHost,
      port: account.smtpPort,
      secure: account.smtpSecure,
      auth: { user: account.smtpUser || account.email, pass: smtpPassword || password },
      connectionTimeout: 15000,
      greetingTimeout: 15000,
      socketTimeout: 20000,
      tls: allowSS ? { rejectUnauthorized: false } : undefined,
    });
    await transporter.verify();
  };
  try {
    await smtpTry(!!account.allowSelfSigned);
  } catch (e) {
    if (isCertError(e)) {
      try {
        await smtpTry(true);
        selfSignedFix = true;
      } catch {
        throw new Error(`${inMsg} • ${explainConnError(e, 'SMTP', outCtx)}`);
      }
    } else {
      throw new Error(`${inMsg} • ${explainConnError(e, 'SMTP', outCtx)}`);
    }
  }

  return {
    message:
      `${inMsg} • SMTP OK` +
      (secureFix !== undefined
        ? ` — ajuste: ${secureFix ? 'ative' : 'desative'} a criptografia do servidor de entrada (aplicada no formulário)`
        : '') +
      (selfSignedFix
        ? ' — o servidor usa certificado autoassinado; a opção "aceitar certificado autoassinado" foi marcada no formulário'
        : ''),
    imapSecureFix: secureFix,
    allowSelfSignedFix: selfSignedFix,
  };
}
