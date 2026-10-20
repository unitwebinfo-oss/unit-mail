import React, { useState } from 'react';
import { Icon } from '../icons';
import { t, useLang } from '../i18n';
import type { MailSource, NewSpecialItem } from '../../../shared/types';

const TITLE_KEYS: Record<NewSpecialItem['kind'], string> = {
  appt: 'Novo Compromisso',
  contact: 'Novo Contato',
  task: 'Nova Tarefa',
};

const EDIT_TITLE_KEYS: Record<NewSpecialItem['kind'], string> = {
  appt: 'Editar Compromisso',
  contact: 'Editar Contato',
  task: 'Editar Tarefa',
};

function F({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="fld">
      <span>{label}</span>
      {children}
    </div>
  );
}

/** Diálogo "Novos Itens" — cria compromisso, contato ou tarefa numa base de dados. */
export function NewItemDialog({
  kind,
  sources,
  primarySourceId,
  initial,
  onClose,
  onSave,
}: {
  kind: NewSpecialItem['kind'];
  /** Bases que aceitam itens (local ou PST com sombra). */
  sources: MailSource[];
  /** Arquivo de dados principal (sourceId aberto) — destino fixo dos itens especiais. */
  primarySourceId?: string;
  /** Valores atuais do item em edição — edição quando presente. */
  initial?: NewSpecialItem;
  onClose: () => void;
  onSave: (sourceId: string, item: NewSpecialItem) => Promise<void> | void;
}) {
  useLang();
  const editing = !!initial;
  const [item, setItem] = useState<NewSpecialItem>(initial ?? { kind, subject: '' });
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState('');

  const up = (patch: Partial<NewSpecialItem>) => setItem((i) => ({ ...i, ...patch }));

  // Compromissos, contatos e tarefas sempre vão para o arquivo de dados principal.
  const target = sources.find((s) => s.id === primarySourceId);
  const targetId = target?.id || '';

  const save = async () => {
    if (!item.subject.trim() && !(kind === 'contact' && item.contactName?.trim())) {
      setErr(kind === 'contact' ? t('Informe o nome do contato.') : t('Informe o assunto.'));
      return;
    }
    if (!targetId) {
      setErr(
        t('O arquivo de dados principal não está aberto. Abra-o ou defina outro em Configurações → Arquivos de Dados.'),
      );
      return;
    }
    setBusy(true);
    setErr('');
    try {
      await onSave(targetId, { ...item, subject: item.subject || item.contactName || '' });
    } catch (e) {
      setErr((e as Error).message);
      setBusy(false);
    }
  };

  return (
    <div className="modal-overlay" style={{ zIndex: 160 }} onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div
        className="set-card acct-modal"
        style={{ display: 'flex', flexDirection: 'column', gap: 12, width: 560 }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <strong style={{ fontSize: 14 }}>
            <Icon
              name={kind === 'appt' ? 'calendar' : kind === 'contact' ? 'contact' : 'task'}
              size={15}
            />{' '}
            {t((editing ? EDIT_TITLE_KEYS : TITLE_KEYS)[kind])}
          </strong>
          <button className="set-close" onClick={onClose} title={t('Fechar')}>
            <Icon name="close" size={15} />
          </button>
        </div>

        <F label={t('Arquivo de dados')}>
          <span style={{ fontSize: 12, color: target ? 'var(--text-dim)' : '#a4262c' }}>
            {target
              ? `${target.label}${editing ? '' : ` (${t('principal')})`}`
              : editing
                ? t('Arquivo de dados do item não está aberto')
                : t('Arquivo de dados principal não está aberto')}
          </span>
        </F>

        {kind === 'contact' ? (
          <>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 10 }}>
              <F label={t('Nome completo')}>
                <input
                  className="inp"
                  autoFocus
                  value={item.contactName || ''}
                  onChange={(e) => up({ contactName: e.target.value })}
                />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <F label={t('Nome')}>
                <input className="inp" value={item.givenName || ''} onChange={(e) => up({ givenName: e.target.value })} />
              </F>
              <F label={t('Sobrenome')}>
                <input className="inp" value={item.surname || ''} onChange={(e) => up({ surname: e.target.value })} />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <F label={t('Email')}>
                <input className="inp" type="email" value={item.email || ''} onChange={(e) => up({ email: e.target.value })} />
              </F>
              <F label={t('Email 2')}>
                <input className="inp" type="email" value={item.email2 || ''} onChange={(e) => up({ email2: e.target.value })} />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label={t('Empresa')}>
                <input className="inp" value={item.company || ''} onChange={(e) => up({ company: e.target.value })} />
              </F>
              <F label={t('Cargo')}>
                <input className="inp" value={item.jobTitle || ''} onChange={(e) => up({ jobTitle: e.target.value })} />
              </F>
              <F label={t('Departamento')}>
                <input className="inp" value={item.dept || ''} onChange={(e) => up({ dept: e.target.value })} />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label={t('Telefone comercial')}>
                <input className="inp" value={item.phone || ''} onChange={(e) => up({ phone: e.target.value })} />
              </F>
              <F label={t('Celular')}>
                <input className="inp" value={item.mobile || ''} onChange={(e) => up({ mobile: e.target.value })} />
              </F>
              <F label={t('Aniversário')}>
                <input className="inp" type="date" value={item.birthday || ''} onChange={(e) => up({ birthday: e.target.value })} />
              </F>
            </div>
          </>
        ) : kind === 'appt' ? (
          <>
            <F label={t('Assunto')}>
              <input className="inp" autoFocus value={item.subject} onChange={(e) => up({ subject: e.target.value })} />
            </F>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <F label={t('Local')}>
                <input className="inp" value={item.location || ''} onChange={(e) => up({ location: e.target.value })} />
              </F>
              <F label={t('Participantes (emails, separados por ;)')}>
                <input className="inp" value={item.attendees || ''} onChange={(e) => up({ attendees: e.target.value })} />
                {editing && !!item.attendees?.trim() && (
                  <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>
                    {t('As alterações serão enviadas aos participantes.')}
                  </span>
                )}
              </F>
            </div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={!!item.allDay}
                onChange={(e) => up({ allDay: e.target.checked })}
              />
              {t('Evento de dia inteiro')}
            </label>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label={t('Início')}>
                <input
                  className="inp"
                  type={item.allDay ? 'date' : 'datetime-local'}
                  value={item.start || ''}
                  onChange={(e) => up({ start: e.target.value })}
                />
              </F>
              <F label={t('Término')}>
                <input
                  className="inp"
                  type={item.allDay ? 'date' : 'datetime-local'}
                  value={item.end || ''}
                  onChange={(e) => up({ end: e.target.value })}
                />
              </F>
              <F label={t('Lembrete')}>
                <select
                  className="inp"
                  value={item.reminderMin == null ? '' : String(item.reminderMin)}
                  onChange={(e) => up({ reminderMin: e.target.value === '' ? null : Number(e.target.value) })}
                >
                  <option value="">{t('Sem lembrete')}</option>
                  <option value="0">{t('Na hora')}</option>
                  <option value="5">{t('5 minutos antes')}</option>
                  <option value="15">{t('15 minutos antes')}</option>
                  <option value="30">{t('30 minutos antes')}</option>
                  <option value="60">{t('1 hora antes')}</option>
                  <option value="1440">{t('1 dia antes')}</option>
                </select>
              </F>
            </div>
          </>
        ) : (
          <>
            <F label={t('Assunto')}>
              <input className="inp" autoFocus value={item.subject} onChange={(e) => up({ subject: e.target.value })} />
            </F>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label={t('Início')}>
                <input className="inp" type="date" value={item.startDate || ''} onChange={(e) => up({ startDate: e.target.value })} />
              </F>
              <F label={t('Vencimento')}>
                <input className="inp" type="date" value={item.dueDate || ''} onChange={(e) => up({ dueDate: e.target.value })} />
              </F>
              <F label={t('% concluída')}>
                <input
                  className="inp"
                  type="number"
                  min={0}
                  max={100}
                  value={item.percentComplete ?? 0}
                  onChange={(e) => up({ percentComplete: Math.max(0, Math.min(100, Number(e.target.value) || 0)) })}
                />
              </F>
            </div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={!!item.isComplete}
                onChange={(e) => up({ isComplete: e.target.checked, percentComplete: e.target.checked ? 100 : item.percentComplete })}
              />
              {t('Tarefa concluída')}
            </label>
          </>
        )}

        <F label={kind === 'contact' ? t('Anotações') : t('Descrição')}>
          <textarea
            className="inp"
            rows={4}
            value={item.body || ''}
            onChange={(e) => up({ body: e.target.value })}
          />
        </F>

        {err && <div style={{ fontSize: 12, color: '#a4262c' }}>{err}</div>}
        <div style={{ display: 'flex', gap: 8 }}>
          <button className="btn btn-primary" disabled={busy} onClick={() => void save()}>
            {busy ? t('Salvando…') : t('Salvar')}
          </button>
          <button className="btn" disabled={busy} onClick={onClose}>
            {t('Cancelar')}
          </button>
        </div>
      </div>
    </div>
  );
}
