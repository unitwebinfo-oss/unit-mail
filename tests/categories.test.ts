import { existsSync, mkdirSync, mkdtempSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { PstService } from '../src/main/pstService';

const EML = Buffer.from(
  'From: Ana <ana@exemplo.com>\r\n' +
    'To: roger@exemplo.com\r\n' +
    'Subject: Orcamento\r\n' +
    'Date: Mon, 10 Mar 2025 10:00:00 -0300\r\n' +
    'Message-ID: <t1@exemplo>\r\n' +
    '\r\n' +
    'corpo da mensagem\r\n',
);

let dir: string;
let svc: PstService;

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'um-cats-'));
  svc = new PstService(dir);
  svc.openLocal('base', dir);
});

afterEach(() => {
  svc.close();
  rmSync(dir, { recursive: true, force: true });
});

describe('Categorias por item (itemcats.json)', () => {
  it('atribui, lista e persiste entre instâncias', async () => {
    svc.importNamedEml('Caixa de Entrada', EML, true);
    svc.setItemCategories('arc:Caixa de Entrada', 0, ['red', 'blue']);
    expect(svc.getItemCategories('arc:Caixa de Entrada', 0)).toEqual(['red', 'blue']);
    expect(existsSync(join(dir, 'itemcats.json'))).toBe(true);

    // Instância nova lê do disco — categorias sobrevivem ao restart
    const svc2 = new PstService(dir);
    svc2.openLocal('base', dir);
    expect(svc2.getItemCategories('arc:Caixa de Entrada', 0)).toEqual(['red', 'blue']);
    const items = await svc2.getItems('arc:Caixa de Entrada', 0, 50);
    expect(items[0].categories).toEqual(['red', 'blue']);
    svc2.close();
  });

  it('limpar com lista vazia remove a marcação', () => {
    svc.importNamedEml('Caixa de Entrada', EML, true);
    svc.setItemCategories('arc:Caixa de Entrada', 0, ['red']);
    svc.setItemCategories('arc:Caixa de Entrada', 0, []);
    expect(svc.getItemCategories('arc:Caixa de Entrada', 0)).toEqual([]);
  });

  it('deduplica ids repetidos', () => {
    svc.importNamedEml('Caixa de Entrada', EML, true);
    svc.setItemCategories('arc:Caixa de Entrada', 0, ['red', 'red', 'blue']);
    expect(svc.getItemCategories('arc:Caixa de Entrada', 0)).toEqual(['red', 'blue']);
  });

  it('mover o .eml entre pastas leva as categorias junto', () => {
    const name = svc.importNamedEml('Caixa de Entrada', EML, true);
    svc.setItemCategories('arc:Caixa de Entrada', 0, ['green']);
    svc.moveNamedFile('Caixa de Entrada', 'Arquivo Morto', name);
    expect(svc.getItemCategories('arc:Caixa de Entrada', 0)).toEqual([]);
    expect(svc.getItemCategories('arc:Arquivo Morto', 0)).toEqual(['green']);
  });

  it('excluir o .eml limpa a marcação', () => {
    svc.importNamedEml('Caixa de Entrada', EML, true);
    svc.setItemCategories('arc:Caixa de Entrada', 0, ['red']);
    svc.deleteItem('arc:Caixa de Entrada', 0);
    expect(existsSync(join(dir, 'Caixa de Entrada'))).toBe(true);
    // marcação removida do mapa — uma segunda instância não a vê
    const svc2 = new PstService(dir);
    svc2.openLocal('base', dir);
    expect(svc2.getItemCategories('arc:Caixa de Entrada', 0)).toEqual([]);
    svc2.close();
  });
});
