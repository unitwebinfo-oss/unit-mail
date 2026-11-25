import { describe, expect, it } from 'vitest';
import { decodeMimeWords, getHeader, parseFrom, parseMailDate, parseReceivedDate } from '../src/main/mailHeaders';

describe('decodeMimeWords', () => {
  it('decodifica Base64 UTF-8', () => {
    expect(decodeMimeWords('=?UTF-8?B?QXZpc28gZGUgcmVzdWx0YWRv?=')).toBe('Aviso de resultado');
  });

  it('decodifica Quoted-Printable com acentos e underscore', () => {
    expect(decodeMimeWords('=?utf-8?Q?Confirma=C3=A7=C3=A3o_de_pedido?=')).toBe('Confirmação de pedido');
  });

  it('mantém texto sem codificação', () => {
    expect(decodeMimeWords('Assunto simples')).toBe('Assunto simples');
  });
});

describe('getHeader', () => {
  const raw = [
    'From: Ana <ana@x.com>',
    'Subject: Reunião',
    '  de planejamento',
    'Date: Mon, 06 Oct 2025 10:00:00 +0000',
    '',
    'corpo',
  ].join('\n');

  it('lê cabeçalho simples', () => {
    expect(getHeader(raw, 'From')).toBe('Ana <ana@x.com>');
  });

  it('une cabeçalho dobrado em múltiplas linhas', () => {
    expect(getHeader(raw, 'Subject')).toBe('Reunião de planejamento');
  });

  it('retorna vazio para cabeçalho ausente', () => {
    expect(getHeader(raw, 'Cc')).toBe('');
  });
});

describe('parseFrom', () => {
  it('separa nome e endereço', () => {
    expect(parseFrom('"Silva, João" <joao@x.com>')).toEqual({ name: 'Silva, João', addr: 'joao@x.com' });
  });

  it('aceita endereço puro', () => {
    expect(parseFrom('joao@x.com')).toEqual({ name: 'joao@x.com', addr: 'joao@x.com' });
  });
});

describe('parseMailDate', () => {
  it('lê RFC 2822 padrão', () => {
    expect(parseMailDate('Mon, 06 Oct 2025 10:00:00 +0000')?.toISOString()).toBe(
      '2025-10-06T10:00:00.000Z',
    );
  });

  it('lê zona textual obsoleta (EST/PST)', () => {
    expect(parseMailDate('Mon, 06 Oct 2025 10:00:00 EST')?.toISOString()).toBe(
      '2025-10-06T15:00:00.000Z',
    );
    expect(parseMailDate('06 Oct 2025 10:00:00 PST')?.toISOString()).toBe(
      '2025-10-06T18:00:00.000Z',
    );
  });

  it('ignora comentários RFC no fim da data', () => {
    expect(parseMailDate('Mon, 06 Oct 2025 10:00:00 -0300 (BRT)')?.toISOString()).toBe(
      '2025-10-06T13:00:00.000Z',
    );
  });

  it('lê ano de 2 dígitos', () => {
    expect(parseMailDate('06 Oct 25 10:00:00 +0000')?.getUTCFullYear()).toBe(2025);
  });

  it('lê formato "Dec 02, 2024" e ISO sem T', () => {
    expect(parseMailDate('Dec 02, 2024 10:00:00')?.getFullYear()).toBe(2024);
    // Sem zona explícita → interpretado como hora local (RFC 5322)
    expect(parseMailDate('2024-12-02 10:00:00')?.getTime()).toBe(
      new Date(2024, 11, 2, 10, 0, 0).getTime(),
    );
  });

  it('retorna null para lixo/vazio', () => {
    expect(parseMailDate('')).toBeNull();
    expect(parseMailDate(null)).toBeNull();
    expect(parseMailDate('isto não é data')).toBeNull();
  });
});

describe('parseReceivedDate', () => {
  it('extrai a data após o ponto-e-vírgula do Received:', () => {
    const rec =
      'from mail.ex.com (mail.ex.com [10.0.0.1]) by mx.provedor.com with ESMTPS; Mon, 06 Oct 2025 10:00:00 -0300';
    expect(parseReceivedDate(rec)?.toISOString()).toBe('2025-10-06T13:00:00.000Z');
  });
});
