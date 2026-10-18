import { describe, expect, it } from 'vitest';
import { decodeMimeWords, getHeader, parseFrom } from '../src/main/mailHeaders';

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
