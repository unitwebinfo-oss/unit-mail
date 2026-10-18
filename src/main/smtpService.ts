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
  const sig = bodyHasSig ? '' : resolveSignatureHtml(account, input.kind);
  const signatureHtml = sig ? `<br><br>${sig}` : '';
  return {
    from: `"${account.label || account.email}" <${account.email}>`,
    to: input.to,
    cc: input.cc || undefined,
    bcc: input.bcc || undefined,
    subject: input.subject,
    html: input.html + signatureHtml,
    inReplyTo: input.inReplyTo || undefined,
    references: input.references || undefined,
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
  });

  await transporter.sendMail(buildMailOptions(account, input));
}

export async function testAccount(
  account: AccountConfig,
  password: string,
  smtpPassword?: string,
): Promise<string> {
  if (account.protocol === 'pop3') {
    const { Pop3Client } = await import('./pop3Client');
    const c = new Pop3Client();
    await c.connect(
      account.imapHost,
      account.imapPort,
      account.imapSecure,
      account.imapUser || account.email,
      password,
    );
    const list = await c.uidl();
    await c.quit();
    // Compara com o estado de sync para mostrar quantas são realmente novas
    const { pop3SeenUids } = await import('./ruleEngine');
    const seen = pop3SeenUids(account);
    const fresh = list.filter((m) => !seen.has(m.uidl)).length;
    return fresh > 0
      ? `POP3 OK — ${list.length} mensagem(ns) no servidor, ${fresh} nova(s) para baixar`
      : `POP3 OK — ${list.length} mensagem(ns) no servidor, todas já baixadas`;
  }
  const { ImapFlow } = await import('imapflow');
  const client = new ImapFlow({
    host: account.imapHost,
    port: account.imapPort,
    secure: account.imapSecure,
    auth: { user: account.imapUser || account.email, pass: password },
    logger: false,
    disableAutoIdle: true,
  });
  await client.connect();
  await client.logout();

  const transporter = nodemailer.createTransport({
    host: account.smtpHost,
    port: account.smtpPort,
    secure: account.smtpSecure,
    auth: { user: account.smtpUser || account.email, pass: smtpPassword || password },
  });
  await transporter.verify();
  return 'Conexão IMAP e SMTP bem-sucedida';
}
