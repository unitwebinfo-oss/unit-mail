import { describe, expect, it } from 'vitest';
import { safeFsName } from '../src/shared/fsName';

describe('safeFsName', () => {
  it('mantém nomes normais, com acento, espaço e hífen', () => {
    expect(safeFsName('Caixa de Entrada')).toBe('Caixa de Entrada');
    expect(safeFsName('Itens Excluídos')).toBe('Itens Excluídos');
    expect(safeFsName('Emails - Gmail')).toBe('Emails - Gmail');
    expect(safeFsName('Relatório 2024')).toBe('Relatório 2024');
  });

  it('remove caracteres proibidos no Windows', () => {
    expect(safeFsName('a/b\\c:d*e?f"g<h>i|j')).toBe('abcdefghij');
    expect(safeFsName('pasta: arquivos')).toBe('pasta arquivos');
  });

  it('corta ponto e espaço no fim — o caso do erro 0x80004005', () => {
    expect(safeFsName('ao arquivo para que as modificações feitas no item.'))
      .toBe('ao arquivo para que as modificações feitas no item');
    expect(safeFsName('pasta ')).toBe('pasta');
    expect(safeFsName('pasta...  ')).toBe('pasta');
    expect(safeFsName('nome. ')).toBe('nome');
  });

  it('corta ponto e espaço no início', () => {
    expect(safeFsName('.oculta')).toBe('oculta');
    expect(safeFsName('  recuo')).toBe('recuo');
  });

  it('protege nomes reservados do DOS', () => {
    expect(safeFsName('CON')).toBe('_CON');
    expect(safeFsName('nul')).toBe('_nul');
    expect(safeFsName('Com1')).toBe('_Com1');
    expect(safeFsName('LPT9')).toBe('_LPT9');
    // com extensão não é reservado
    expect(safeFsName('con.txt')).toBe('con.txt');
  });

  it('usa o fallback quando sobra nada', () => {
    expect(safeFsName('...')).toBe('Pasta');
    expect(safeFsName('')).toBe('Pasta');
    expect(safeFsName('???')).toBe('Pasta');
    expect(safeFsName('...', 'Base')).toBe('Base');
    expect(safeFsName('...', '')).toBe('');
  });

  it('remove caracteres de controle', () => {
    expect(safeFsName('tab\tname')).toBe('tabname');
    expect(safeFsName('line\nbreak')).toBe('linebreak');
  });
});
