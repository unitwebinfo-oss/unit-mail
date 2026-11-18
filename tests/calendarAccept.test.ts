import { mkdtempSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { PstService } from '../src/main/pstService';

const ICS = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'METHOD:PUBLISH',
  'BEGIN:VEVENT',
  'UID:teste-uid-1',
  'SUMMARY:Reuniao teste1',
  'DTSTART:20261030T140000Z',
  'DTEND:20261030T150000Z',
  'ORGANIZER;CN=Contato:mailto:contato@unitrs.com.br',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

const EML = [
  'From: contato@unitrs.com.br',
  'To: unitwebinfo@gmail.com',
  'Subject: Reuniao teste1',
  'Date: Mon, 28 Oct 2026 10:00:00 GMT',
  'MIME-Version: 1.0',
  'X-UM-Class: IPM.Appointment',
  'X-UM-Appt-Uid: teste-uid-1',
  'Content-Type: multipart/mixed; boundary="ol-cal"',
  '',
  '--ol-cal',
  'Content-Type: text/plain; charset=utf-8',
  '',
  'Compromisso: Reuniao teste1',
  '',
  '--ol-cal',
  'Content-Type: text/calendar; method=PUBLISH; name="invite.ics"',
  'Content-Disposition: attachment; filename="invite.ics"',
  '',
  ICS,
  '--ol-cal--',
].join('\r\n');

describe('compromisso aceito grava no Calendário', () => {
  let dir: string;
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'um-cal-'));
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  it('aparece na pasta Calendário como compromisso (não como convite)', async () => {
    const svc = new PstService();
    svc.openLocal('teste', dir);
    await svc.importItemEml('arc:Calendário', Buffer.from(EML, 'utf8'));

    const cal = svc.getFolderTree().find((n) => n.name === 'Calendário');
    expect(cal).toBeDefined();
    const items = await svc.getItems(cal!.id, 0, 500);
    expect(items).toHaveLength(1);
    expect(items[0].startTime).toBe('2026-10-30T14:00:00.000Z');
    expect(items[0].apptUid).toBe('teste-uid-1');
    expect(items[0].inviteMethod).toBeUndefined();
  });

  it('preserva corpo HTML, participantes e anexos do convite no compromisso', async () => {
    // Mesma estrutura que respondInvite grava desde que passou a copiar o
    // convite: alternative(plain+html) + ics PUBLISH + anexos comuns.
    const html = Buffer.from(
      '<div><b>Quando</b> quinta 1 out. 2026 — Convidados: Ana, Bruno</div>',
      'utf8',
    ).toString('base64');
    const plain = Buffer.from('Quando: quinta 1 out. 2026', 'utf8').toString('base64');
    const pdf = Buffer.from('%PDF-1.4 fake').toString('base64');
    const eml = [
      'From: contato@unitrs.com.br',
      'To: unitwebinfo@gmail.com',
      'Subject: Reuniao rica',
      'Date: Mon, 28 Oct 2026 10:00:00 GMT',
      'MIME-Version: 1.0',
      'X-UM-Class: IPM.Appointment',
      'X-UM-Appt-Uid: teste-uid-2',
      'X-UM-Appt-Attendees: ana@x.com, bruno@x.com',
      'Content-Type: multipart/mixed; boundary="ol-cal"',
      '',
      '--ol-cal',
      'Content-Type: multipart/alternative; boundary="ol-alt"',
      '',
      '--ol-alt',
      'Content-Type: text/plain; charset=utf-8',
      'Content-Transfer-Encoding: base64',
      '',
      plain,
      '--ol-alt',
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: base64',
      '',
      html,
      '--ol-alt--',
      '--ol-cal',
      'Content-Type: text/calendar; method=PUBLISH; name="invite.ics"',
      'Content-Disposition: attachment; filename="invite.ics"',
      '',
      ICS,
      '--ol-cal',
      'Content-Type: application/pdf; name="alinhamento.pdf"',
      'Content-Transfer-Encoding: base64',
      'Content-Disposition: attachment; filename="alinhamento.pdf"',
      '',
      pdf,
      '--ol-cal--',
    ].join('\r\n');

    const svc = new PstService();
    svc.openLocal('teste', dir);
    await svc.importItemEml('arc:Calendário', Buffer.from(eml, 'utf8'));

    const cal = svc.getFolderTree().find((n) => n.name === 'Calendário');
    const items = await svc.getItems(cal!.id, 0, 500);
    expect(items).toHaveLength(1);
    const detail = await svc.getItem(cal!.id, 0, items[0].stableId);
    expect(detail.bodyHTML).toContain('Convidados: Ana, Bruno');
    expect(detail.requiredAttendees).toBe('ana@x.com, bruno@x.com');
    const pdfAtt = detail.attachments.find((a) => a.filename === 'alinhamento.pdf');
    expect(pdfAtt).toBeDefined();
    const data = await svc.getAttachmentData(cal!.id, 0, pdfAtt!.index, items[0].stableId);
    expect(Buffer.from(data.dataBase64, 'base64').toString()).toBe('%PDF-1.4 fake');
  });
});
