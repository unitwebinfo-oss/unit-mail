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
});
