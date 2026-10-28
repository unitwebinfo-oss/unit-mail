import { mkdtempSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { PstService, parseIcsInvite } from '../src/main/pstService';

const ICS_TZID = [
  'BEGIN:VCALENDAR',
  'PRODID:-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN',
  'VERSION:2.0',
  'METHOD:REQUEST',
  'BEGIN:VTIMEZONE',
  'TZID:E. South America Standard Time',
  'BEGIN:STANDARD',
  'DTSTART:16010101T000000',
  'TZOFFSETFROM:-0200',
  'TZOFFSETTO:-0300',
  'END:STANDARD',
  'END:VTIMEZONE',
  'BEGIN:VEVENT',
  'UID:outlook-uid-1',
  'SUMMARY:Reuniao com cliente',
  'DTSTART;TZID=E. South America Standard Time:20261030T140000',
  'DTEND;TZID=E. South America Standard Time:20261030T150000',
  'LOCATION:Sala 2',
  'ORGANIZER;CN=Contato Unit:mailto:contato@unitrs.com.br',
  'ATTENDEE;RSVP=TRUE:mailto:unitwebinfo@gmail.com',
  'BEGIN:VALARM',
  'TRIGGER:-PT15M',
  'ACTION:DISPLAY',
  'END:VALARM',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

const ICS_UTC = ICS_TZID.replace(
  /DTSTART;TZID=[^:]*:20261030T140000/,
  'DTSTART:20261030T170000Z',
).replace(/DTEND;TZID=[^:]*:20261030T150000/, 'DTEND:20261030T180000Z');

function inviteEml(ics: string): Buffer {
  return Buffer.from(
    [
      'From: Contato Unit <contato@unitrs.com.br>',
      'To: unitwebinfo@gmail.com',
      'Subject: Reuniao com cliente',
      'Date: Mon, 28 Oct 2026 10:00:00 -0300',
      'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="b1"',
      '',
      '--b1',
      'Content-Type: text/plain; charset=utf-8',
      '',
      'Convite para reuniao.',
      '',
      '--b1',
      'Content-Type: text/calendar; method=REQUEST; charset=utf-8',
      'Content-Transfer-Encoding: 7bit',
      'Content-Disposition: attachment; filename="invite.ics"',
      '',
      ics,
      '--b1--',
    ].join('\r\n'),
    'utf8',
  );
}

describe('convite recebido: detecção e datas', () => {
  let dir: string;
  let svc: PstService;
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'um-inv-'));
    svc = new PstService();
    svc.openLocal('teste', dir);
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  it('parseIcsInvite lê convite com TZID (Outlook)', () => {
    const inv = parseIcsInvite(ICS_TZID);
    expect(inv?.method).toBe('REQUEST');
    expect(inv?.uid).toBe('outlook-uid-1');
    expect(inv?.summary).toBe('Reuniao com cliente');
    expect(inv?.location).toBe('Sala 2');
    expect(inv?.organizerEmail).toBe('contato@unitrs.com.br');
  });

  it('data com TZID vira o instante correto (14:00 em Brasília = 17:00Z)', () => {
    const inv = parseIcsInvite(ICS_TZID);
    expect(inv?.start).toBe('2026-10-30T17:00:00.000Z');
  });

  it('data em UTC (Z) continua correta', () => {
    const inv = parseIcsInvite(ICS_UTC);
    expect(inv?.start).toBe('2026-10-30T17:00:00.000Z');
  });

  it('aceitar grava o compromisso no Calendário na data certa, sem virar convite', async () => {
    await svc.importItemEml('arc:Caixa de Entrada', inviteEml(ICS_TZID));
    const d = await svc.getItem('arc:Caixa de Entrada', 0);
    const ics = (d.invite?.ics || '').replace(/^METHOD:[^\r\n]*/m, 'METHOD:PUBLISH');
    const eml = [
      'From: contato@unitrs.com.br',
      'To: unitwebinfo@gmail.com',
      'Subject: Reuniao com cliente',
      'Date: Mon, 28 Oct 2026 10:00:00 GMT',
      'MIME-Version: 1.0',
      'X-UM-Class: IPM.Appointment',
      `X-UM-Appt-Uid: ${d.invite?.uid}`,
      'Content-Type: multipart/mixed; boundary="ol-cal"',
      '',
      '--ol-cal',
      'Content-Type: text/plain; charset=utf-8',
      '',
      'Compromisso',
      '',
      '--ol-cal',
      'Content-Type: text/calendar; method=PUBLISH; name="invite.ics"',
      'Content-Disposition: attachment; filename="invite.ics"',
      '',
      ics,
      '--ol-cal--',
    ].join('\r\n');
    await svc.importItemEml('arc:Calendário', Buffer.from(eml, 'utf8'));
    const cal = svc.getFolderTree().find((n) => n.name === 'Calendário');
    const items = await svc.getItems(cal!.id, 0, 500);
    expect(items).toHaveLength(1);
    expect(items[0].startTime).toBe('2026-10-30T17:00:00.000Z');
    expect(items[0].endTime).toBe('2026-10-30T18:00:00.000Z');
    expect(items[0].location).toBe('Sala 2');
    expect(items[0].inviteMethod).toBeUndefined();
  });

  it('clique no compromisso abre o certo mesmo com outro chegando antes (stableId)', async () => {
    const appt = (uid: string, subject: string, start: string) =>
      Buffer.from(
        [
          'From: x@y.com',
          'To: z@y.com',
          `Subject: ${subject}`,
          'Date: Mon, 28 Oct 2026 10:00:00 GMT',
          'MIME-Version: 1.0',
          'X-UM-Class: IPM.Appointment',
          'Content-Type: multipart/mixed; boundary="ol-cal"',
          '',
          '--ol-cal',
          'Content-Type: text/calendar; method=PUBLISH; name="invite.ics"',
          'Content-Disposition: attachment; filename="invite.ics"',
          '',
          ['BEGIN:VCALENDAR', 'VERSION:2.0', 'METHOD:PUBLISH', 'BEGIN:VEVENT', `UID:${uid}`, `SUMMARY:${subject}`, `DTSTART:${start}`, `DTEND:${start}`, 'END:VEVENT', 'END:VCALENDAR'].join('\r\n'),
          '--ol-cal--',
        ].join('\r\n'),
        'utf8',
      );
    await svc.importItemEml('arc:Calendário', appt('u-a', 'Primeiro', '20261030T120000Z'));
    await new Promise((r) => setTimeout(r, 5));
    const cal = svc.getFolderTree().find((n) => n.name === 'Calendário')!;
    const before = await svc.getItems(cal.id, 0, 500);
    const target = before.find((i) => i.subject === 'Primeiro')!;
    await new Promise((r) => setTimeout(r, 5));
    await svc.importItemEml('arc:Calendário', appt('u-b', 'Segundo', '20261031T120000Z'));
    const d = await svc.getItem(cal.id, target.index, target.stableId);
    expect(d.subject).toBe('Primeiro');
    expect(d.invite).toBeUndefined();
  });

  it('getItem expõe o convite pendente com o método REQUEST', async () => {
    await svc.importItemEml('arc:Caixa de Entrada', inviteEml(ICS_UTC));
    const d = await svc.getItem('arc:Caixa de Entrada', 0);
    expect(d.invite?.method).toBe('REQUEST');
    expect(d.invite?.uid).toBe('outlook-uid-1');
    expect(d.invite?.start).toBe('2026-10-30T17:00:00.000Z');
  });
});
