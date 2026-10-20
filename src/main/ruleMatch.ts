import type { MailRule } from '../shared/types';

export function matchRule(
  rule: MailRule,
  env: { from: string; to: string; subject: string },
): boolean {
  const c = rule.conditions || {};
  // Condições são cortadas (trim) antes de casar — um campo só com espaços
  // conta como "vazio" (regra sem condição pega tudo), não como texto a
  // procurar — senão from.includes(' ') casava com TODAS as mensagens.
  const fromC = (c.fromContains || '').trim().toLowerCase();
  const toC = (c.toContains || '').trim().toLowerCase();
  const subjC = (c.subjectContains || '').trim().toLowerCase();
  if (!fromC && !toC && !subjC) return true; // regra sem condição = pega tudo (igual "aplicar a todas" do Outlook)
  const from = env.from.toLowerCase();
  const to = env.to.toLowerCase();
  const subj = env.subject.toLowerCase();
  if (fromC && !from.includes(fromC)) return false;
  if (toC && !to.includes(toC)) return false;
  if (subjC && !subj.includes(subjC)) return false;
  return true;
}
