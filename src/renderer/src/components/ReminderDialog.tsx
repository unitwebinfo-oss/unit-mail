import React, { useState } from 'react';
import { Icon } from '../icons';
import { t, useLang, locale } from '../i18n';

/** Lembrete disparado pelo main (calendar:reminder). */
export interface ReminderEvent {
  key: string;
  subject: string;
  startTime: string;
  location: string;
  sourceId: string;
  folderId: string;
  index: number;
  /** Identidade estável do item — resolve a posição atual no backend. */
  stableId?: string;
  /** Módulo de origem do lembrete — 'tasks' para tarefas (padrão: calendário). */
  module?: string;
}

const SNOOZE_MINUTES = [5, 10, 15, 30, 60];

/**
 * Janela de lembretes estilo Outlook: lista os compromissos disparados com
 * ações por item (Abrir / Dispensar) e "Adiar" na barra inferior.
 */
export function ReminderDialog({
  reminders,
  onDismiss,
  onDismissAll,
  onSnooze,
  onOpen,
}: {
  reminders: ReminderEvent[];
  onDismiss: (key: string) => void;
  onDismissAll: () => void;
  onSnooze: (minutes: number) => void;
  onOpen: (r: ReminderEvent) => void;
}) {
  useLang();
  const [snoozeMin, setSnoozeMin] = useState(5);
  const [snoozeOpen, setSnoozeOpen] = useState(false);

  if (!reminders.length) return null;

  return (
    <div className="modal-overlay" style={{ zIndex: 300 }}>
      <div className="reminder-window" onClick={(e) => e.stopPropagation()}>
        <div className="reminder-head">
          <strong>
            <Icon name="calendar" size={15} />{' '}
            {reminders.length === 1
              ? t('1 Lembrete')
              : t('{n} Lembretes', { n: reminders.length })}
          </strong>
          <span className="close-x danger" title={t('Fechar')} onClick={onDismissAll}>
            <Icon name="close" size={15} />
          </span>
        </div>

        <div className="reminder-list">
          {reminders.map((r) => (
            <div className="reminder-row" key={r.key}>
              <Icon name="calendar" size={18} color="#0173c7" />
              <div className="reminder-info">
                <strong>{r.subject || t('Compromisso')}</strong>
                <span>
                  {r.startTime ? new Date(r.startTime).toLocaleString(locale()) : ''}
                  {r.location ? ` — ${r.location}` : ''}
                </span>
              </div>
              <div className="reminder-row-actions">
                <button className="btn" onClick={() => onOpen(r)} title={t('Abrir item')}>
                  {t('Abrir item')}
                </button>
                <button className="btn" onClick={() => onDismiss(r.key)} title={t('Dispensar')}>
                  {t('Dispensar')}
                </button>
              </div>
            </div>
          ))}
        </div>

        <div className="reminder-foot">
          <div className="reminder-snooze">
            <button className="btn btn-primary" onClick={() => onSnooze(snoozeMin)}>
              {t('Adiar')}
            </button>
            <span className="rp-more" style={{ position: 'relative' }}>
              <button className="btn" onClick={() => setSnoozeOpen((v) => !v)}>
                {snoozeMin === 60
                  ? t('1 hora')
                  : t('{n} minutos', { n: snoozeMin })}{' '}
                ▾
              </button>
              {snoozeOpen && (
                <>
                  <div className="menu-backdrop" onClick={() => setSnoozeOpen(false)} />
                  <div className="rp-more-menu" style={{ bottom: '100%', top: 'auto' }}>
                    {SNOOZE_MINUTES.map((m) => (
                      <button
                        key={m}
                        onClick={() => {
                          setSnoozeMin(m);
                          setSnoozeOpen(false);
                        }}
                      >
                        {m === 60 ? t('1 hora') : t('{n} minutos', { n: m })}
                      </button>
                    ))}
                  </div>
                </>
              )}
            </span>
          </div>
          <button className="btn" onClick={onDismissAll}>
            {t('Dispensar todos')}
          </button>
        </div>
      </div>
    </div>
  );
}
