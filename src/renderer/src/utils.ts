import { locale } from './i18n';
import { resolveFolderName, type FolderKind, type Lang } from '../../shared/folderNames';
import type { FolderNode } from '../../shared/types';

const now = () => new Date();

/** Aplica o tema ('system' segue o SO via prefers-color-scheme). */
export function applyTheme(t: 'system' | 'light' | 'dark' | string): void {
  const dark =
    t === 'dark' ||
    (t === 'system' && window.matchMedia('(prefers-color-scheme: dark)').matches);
  document.body.classList.toggle('dark', dark);
}

/**
 * Chime de lembrete de compromisso (estilo Outlook) — sintetizado com Web
 * Audio, sem arquivo de som. Três notas curtas em sequência; o Electron não
 * exige gesto do usuário para autoplay, então toca mesmo sem clique recente.
 */
export function playReminderSound(): void {
  try {
    const Ctor =
      window.AudioContext ??
      (window as unknown as { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
    if (!Ctor) return;
    const ctx = new Ctor();
    const t0 = ctx.currentTime;
    // Lá–Mi–Lá — arpejo curto e discreto, como o ding do Outlook
    [880, 659.25, 880].forEach((freq, i) => {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.value = freq;
      const start = t0 + i * 0.22;
      gain.gain.setValueAtTime(0, start);
      gain.gain.linearRampToValueAtTime(0.28, start + 0.02);
      gain.gain.exponentialRampToValueAtTime(0.001, start + 0.5);
      osc.connect(gain).connect(ctx.destination);
      osc.start(start);
      osc.stop(start + 0.55);
    });
    // O contexto fica ocioso depois das notas — fecha para não acumular
    setTimeout(() => void ctx.close().catch(() => {}), 1800);
  } catch {
    // sem áudio disponível — lembrete visual continua funcionando
  }
}

export function fmtSmartDate(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const n = now();
  const sameDay =
    d.getDate() === n.getDate() && d.getMonth() === n.getMonth() && d.getFullYear() === n.getFullYear();
  if (sameDay) {
    return d.toLocaleTimeString(locale(), { hour: '2-digit', minute: '2-digit' });
  }
  const weekAgo = new Date(n);
  weekAgo.setDate(n.getDate() - 6);
  if (d >= weekAgo) {
    return d.toLocaleDateString(locale(), { weekday: 'short' });
  }
  if (d.getFullYear() === n.getFullYear()) {
    return d.toLocaleDateString(locale(), { day: '2-digit', month: 'short' });
  }
  return d.toLocaleDateString(locale());
}

export function fmtFullDate(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleString(locale(), {
    weekday: 'long',
    day: '2-digit',
    month: 'long',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

export function fmtDateTime(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  return d.toLocaleString(locale(), { dateStyle: 'short', timeStyle: 'short' });
}

export function fmtTime(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleTimeString(locale(), { hour: '2-digit', minute: '2-digit' });
}

export function fmtSize(bytes: number): string {
  if (!bytes || bytes <= 0) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** Classe de contêiner esperada por módulo */
export function containerClassForModule(module: string): string {
  switch (module) {
    case 'calendar':
      return 'IPF.Appointment';
    case 'people':
      return 'IPF.Contact';
    case 'tasks':
      return 'IPF.Task';
    default:
      return 'IPF.Note';
  }
}

/** Nome de ícone por pasta (estilo Outlook) */
export function folderIconKey(name: string, containerClass: string): string {
  const n = (name || '').toLowerCase();
  if (n.includes('infectad')) return 'virus';
  if (n.includes('inbox') || n.includes('caixa de entrada')) return 'inbox';
  if (n.includes('sent') || n.includes('enviados') || n.includes('itens enviados')) return 'sent';
  if (n.includes('draft') || n.includes('rascunho')) return 'draft';
  if (n.includes('junk') || n.includes('lixo eletr')) return 'junk';
  if (n.includes('deleted') || n.includes('exclu') || n.includes('lixo')) return 'trash';
  if (n.includes('outbox') || n.includes('caixa de sa')) return 'outbox';
  if (n.includes('archive') || n.includes('arquiv')) return 'archive';
  if (containerClass === 'IPF.Appointment') return 'calendar';
  if (containerClass === 'IPF.Contact') return 'contact';
  if (containerClass === 'IPF.Task') return 'task';
  if (containerClass === 'IPF.StickyNote') return 'note';
  if (containerClass === 'IPF.Journal') return 'journal';
  return 'folder';
}

/** Encontra a primeira pasta de uma classe de contêiner na árvore */
export function guessServers(email: string): { imapHost: string; smtpHost: string } {
  const domain = email.split('@')[1]?.toLowerCase() || '';
  const presets: Record<string, { imapHost: string; smtpHost: string }> = {
    'gmail.com': { imapHost: 'imap.gmail.com', smtpHost: 'smtp.gmail.com' },
    'outlook.com': { imapHost: 'outlook.office365.com', smtpHost: 'smtp.office365.com' },
    'hotmail.com': { imapHost: 'outlook.office365.com', smtpHost: 'smtp.office365.com' },
    'yahoo.com': { imapHost: 'imap.mail.yahoo.com', smtpHost: 'smtp.mail.yahoo.com' },
  };
  return presets[domain] || { imapHost: domain ? `imap.${domain}` : '', smtpHost: domain ? `smtp.${domain}` : '' };
}

export function findFirstFolder<T extends { containerClass: string; children: T[] }>(
  nodes: T[],
  containerClass: string,
): T | null {
  for (const n of nodes) {
    if (n.containerClass === containerClass) return n;
    const found = findFirstFolder(n.children, containerClass);
    if (found) return found;
  }
  return null;
}

/**
 * Id "arc:<nome>" da pasta padrão de um tipo (lixo, lixeira, arquivo…) na
 * base cuja árvore está carregada — usa a pasta EXISTENTE em qualquer
 * idioma (nunca aponta para um nome que não existe); só quando nenhuma
 * variante está na árvore usa o nome do idioma atual, que o main cria
 * sob demanda na primeira gravação.
 */
export function arcFolderId(tree: FolderNode[], kind: FolderKind, lang: Lang): string {
  const names: string[] = [];
  const walk = (nodes: FolderNode[]) => {
    for (const n of nodes) {
      if (n.id.startsWith('arc:')) names.push(n.id.slice(4));
      if (n.archiveName) names.push(n.archiveName);
      walk(n.children);
    }
  };
  walk(tree);
  return `arc:${resolveFolderName(names, kind, lang)}`;
}
