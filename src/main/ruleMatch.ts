import type { MailRule } from '../shared/types';

export function matchRule(
  rule: MailRule,
  env: { from: string; to: string; subject: string },
): boolean {
  const c = rule.conditions;
  const any =
    (c.fromContains || '').trim() ||
    (c.toContains || '').trim() ||
    (c.subjectContains || '').trim();
  if (!any) return true; // regra sem condição = pega tudo (igual "aplicar a todas" do Outlook)
  const from = env.from.toLowerCase();
  const to = env.to.toLowerCase();
  const subj = env.subject.toLowerCase();
  if (c.fromContains && !from.includes(c.fromContains.toLowerCase())) return false;
  if (c.toContains && !to.includes(c.toContains.toLowerCase())) return false;
  if (c.subjectContains && !subj.includes(c.subjectContains.toLowerCase())) return false;
  return true;
}
