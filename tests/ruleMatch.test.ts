import { describe, expect, it } from 'vitest';
import { matchRule } from '../src/main/ruleMatch';
import type { MailRule } from '../src/shared/types';

const rule = (conditions: Partial<MailRule['conditions']>) =>
  ({ id: 'r', name: 'r', enabled: true, conditions }) as unknown as MailRule;

const env = { from: 'Maria <maria@Cliente.com.br>', to: 'roger@unit.com', subject: 'Fatura de Março' };

describe('matchRule', () => {
  it('regra sem condição pega tudo', () => {
    expect(matchRule(rule({}), env)).toBe(true);
    expect(matchRule(rule({ fromContains: '   ' }), env)).toBe(true);
  });

  it('compara sem diferenciar maiúsculas', () => {
    expect(matchRule(rule({ fromContains: 'CLIENTE.com' }), env)).toBe(true);
    expect(matchRule(rule({ subjectContains: 'fatura' }), env)).toBe(true);
  });

  it('exige todas as condições preenchidas (E lógico)', () => {
    expect(matchRule(rule({ fromContains: 'maria', subjectContains: 'fatura' }), env)).toBe(true);
    expect(matchRule(rule({ fromContains: 'maria', subjectContains: 'recibo' }), env)).toBe(false);
  });

  it('não casa quando o campo não contém o texto', () => {
    expect(matchRule(rule({ toContains: 'outro@' }), env)).toBe(false);
  });
});
