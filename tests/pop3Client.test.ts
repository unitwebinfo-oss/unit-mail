import net from 'net';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { Pop3Client } from '../src/main/pop3Client';

/**
 * Servidor POP3 falso em TCP local. `retr` responde com o corpo no MESMO
 * pacote do "+OK" (caso real que quebrava) ou fragmentado em dois pacotes.
 */
let server: net.Server;
let port: number;
let splitRetr = false;

const BODY = 'Subject: teste\r\n\r\nLinha\r\n..ponto\r\n';
/**
 * O antivírus ESET (proteção de email) intercepta POP3 em 127.0.0.1 e injeta
 * "X-EsetId:" no cabeçalho. Ignoramos esse cabeçalho na comparação.
 */
const body = (b: Buffer) => b.toString('latin1').replace(/^X-EsetId:[^\r\n]*\r\n/m, '');

beforeEach(async () => {
  server = net.createServer((sock) => {
    sock.write('+OK pronto\r\n');
    let buf = '';
    sock.on('data', (d) => {
      buf += d.toString('latin1');
      let idx;
      while ((idx = buf.indexOf('\r\n')) !== -1) {
        const cmd = buf.slice(0, idx);
        buf = buf.slice(idx + 2);
        if (cmd.startsWith('RETR')) {
          if (splitRetr) {
            sock.write('+OK 40 octets\r\n' + BODY.slice(0, 10));
            setTimeout(() => sock.write(BODY.slice(10) + '.\r\n'), 20);
          } else {
            sock.write('+OK 40 octets\r\n' + BODY + '.\r\n');
          }
        } else if (cmd === 'UIDL') {
          sock.write('+OK\r\n1 abc\r\n2 def\r\n.\r\n');
        } else if (cmd === 'QUIT') {
          sock.write('+OK tchau\r\n');
          sock.end();
        } else {
          sock.write('+OK\r\n');
        }
      }
    });
  });
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()));
  port = (server.address() as net.AddressInfo).port;
});

afterEach(async () => {
  splitRetr = false;
  await new Promise<void>((r) => server.close(() => r()));
});

describe('Pop3Client', () => {
  it('RETR devolve o corpo exato (dot-unstuffing) quando vem no mesmo pacote do +OK', async () => {
    const c = new Pop3Client();
    await c.connect('127.0.0.1', port, false, 'u', 'p');
    const raw = await c.retr(1);
    expect(body(raw)).toBe(BODY.replace('..ponto', '.ponto').slice(0, -2));
    await c.quit();
  });

  it('RETR aguarda o terminador quando a resposta chega fragmentada', async () => {
    splitRetr = true;
    const c = new Pop3Client();
    await c.connect('127.0.0.1', port, false, 'u', 'p');
    const raw = await c.retr(1);
    expect(body(raw)).toBe(BODY.replace('..ponto', '.ponto').slice(0, -2));
    await c.quit();
  });

  it('comandos seguintes funcionam depois de um RETR (buffer não fica corrompido)', async () => {
    const c = new Pop3Client();
    await c.connect('127.0.0.1', port, false, 'u', 'p');
    await c.retr(1);
    const list = await c.uidl();
    expect(list).toEqual([
      { num: 1, uidl: 'abc' },
      { num: 2, uidl: 'def' },
    ]);
    await c.quit();
  });
});
