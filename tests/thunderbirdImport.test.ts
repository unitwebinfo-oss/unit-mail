import { describe, expect, it } from 'vitest';
import { accountsFromPrefs, parsePrefs } from '../src/main/thunderbirdImport';

const PREFS = `
# Mozilla User Preferences
user_pref("mail.accountmanager.accounts", "account1,account2,account3");
user_pref("mail.account.account1.server", "server1");
user_pref("mail.account.account1.identities", "id1");
user_pref("mail.account.account2.server", "server2");
user_pref("mail.account.account2.identities", "id2");
user_pref("mail.account.account3.server", "server3");
user_pref("mail.server.server1.type", "pop3");
user_pref("mail.server.server1.hostname", "pop.exemplo.com.br");
user_pref("mail.server.server1.port", 995);
user_pref("mail.server.server1.socketType", 3);
user_pref("mail.server.server1.leave_on_server", true);
user_pref("mail.server.server1.delete_mail_left_on_server", true);
user_pref("mail.server.server1.num_days_to_leave_on_server", 14);
user_pref("mail.server.server2.type", "imap");
user_pref("mail.server.server2.hostname", "imap.exemplo.com.br");
user_pref("mail.server.server2.port", -1);
user_pref("mail.server.server2.socketType", 2);
user_pref("mail.server.server3.type", "none");
user_pref("mail.identity.id1.useremail", "roger@exemplo.com.br");
user_pref("mail.identity.id1.fullName", "Roger \\"RO\\" Oliveira");
user_pref("mail.identity.id1.smtpServer", "smtp1");
user_pref("mail.identity.id2.useremail", "contato@exemplo.com.br");
user_pref("mail.identity.id2.fullName", "Contato");
user_pref("mail.identity.id2.smtpServer", "smtp1");
user_pref("mail.smtpserver.smtp1.hostname", "smtp.exemplo.com.br");
user_pref("mail.smtpserver.smtp1.port", 465);
user_pref("mail.smtpserver.smtp1.socketType", 3);
`;

describe('Thunderbird — leitura de prefs.js', () => {
  const guesses = accountsFromPrefs(parsePrefs(PREFS));

  it('lê o valor de texto com aspas escapadas', () => {
    expect(parsePrefs(PREFS).get('mail.identity.id1.fullName')).toBe('Roger "RO" Oliveira');
  });

  it('importa contas POP3 e IMAP e ignora Pastas Locais (type none)', () => {
    expect(guesses.map((g) => g.email)).toEqual(['roger@exemplo.com.br', 'contato@exemplo.com.br']);
  });

  it('mapeia POP3 com SSL, porta e retenção no servidor', () => {
    const pop = guesses[0];
    expect(pop.protocol).toBe('pop3');
    expect(pop.imapHost).toBe('pop.exemplo.com.br');
    expect(pop.imapPort).toBe(995);
    expect(pop.imapSecure).toBe(true);
    expect(pop.leaveOnServer).toBe(true);
    expect(pop.leaveDays).toBe(14);
  });

  it('usa a porta padrão quando o Thunderbird grava -1, e STARTTLS não vira SSL', () => {
    const imap = guesses[1];
    expect(imap.protocol).toBe('imap');
    expect(imap.imapPort).toBe(143);
    expect(imap.imapSecure).toBe(false);
  });

  it('lê o SMTP da identidade (SSL na porta 465)', () => {
    expect(guesses[0].smtpHost).toBe('smtp.exemplo.com.br');
    expect(guesses[0].smtpPort).toBe(465);
    expect(guesses[0].smtpSecure).toBe(true);
    expect(guesses[1].smtpHost).toBe('smtp.exemplo.com.br');
  });

  it('nunca traz senha (o Thunderbird as guarda cifradas)', () => {
    expect(guesses.every((g) => g.password === undefined)).toBe(true);
  });
});

describe('mbox → .eml', () => {
  it('separa mensagens pela linha "From "', async () => {
    const { mboxMessages } = await import('../src/main/thunderbirdImport');
    const { writeFileSync, mkdtempSync, rmSync } = await import('fs');
    const { join } = await import('path');
    const { tmpdir } = await import('os');
    const dir = mkdtempSync(join(tmpdir(), 'mbox-'));
    const file = join(dir, 'Inbox');
    writeFileSync(
      file,
      'From a@b Sat Jan  1 00:00:00 2022\nSubject: um\n\nbody1\nFrom c@d Sun Jan  2 00:00:00 2022\nSubject: dois\n\nbody2\n',
    );
    try {
      const msgs: Buffer[] = [];
      for await (const m of mboxMessages(file)) msgs.push(m);
      expect(msgs).toHaveLength(2);
      expect(msgs[0].toString('latin1')).toContain('Subject: um');
      expect(msgs[1].toString('latin1')).toContain('Subject: dois');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('lê a flag de lida e remove cabeçalhos X-Mozilla-*', async () => {
    const { mboxMessageToEml } = await import('../src/main/thunderbirdImport');
    const read = mboxMessageToEml(
      Buffer.from('Subject: x\nX-Mozilla-Status: 0001\nX-Mozilla-Status2: 00000000\n\nbody', 'latin1'),
    )!;
    expect(read.read).toBe(true);
    expect(read.eml.toString('latin1')).not.toContain('X-Mozilla');
    const unread = mboxMessageToEml(Buffer.from('Subject: x\nX-Mozilla-Status: 0000\n\nb', 'latin1'))!;
    expect(unread.read).toBe(false);
  });

  it('desescapa linhas ">From " do corpo', async () => {
    const { mboxMessageToEml } = await import('../src/main/thunderbirdImport');
    const r = mboxMessageToEml(Buffer.from('Subject: x\n\n>From dentro do corpo\n>>From duplo', 'latin1'))!;
    const eml = r.eml.toString('latin1');
    expect(eml).toContain('\nFrom dentro do corpo');
    expect(eml).toContain('\n>From duplo');
  });
});
