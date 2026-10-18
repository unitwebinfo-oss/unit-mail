import React, { useState } from 'react';
import { Icon } from '../icons';
import type { MailSource, NewSpecialItem } from '../../../shared/types';

const TITLES: Record<NewSpecialItem['kind'], string> = {
  appt: 'Novo Compromisso',
  contact: 'Novo Contato',
  task: 'Nova Tarefa',
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
  onClose,
  onSave,
}: {
  kind: NewSpecialItem['kind'];
  /** Bases que aceitam itens (local ou PST com sombra). */
  sources: MailSource[];
  /** Arquivo de dados principal (sourceId aberto) — destino fixo dos itens especiais. */
  primarySourceId?: string;
  onClose: () => void;
  onSave: (sourceId: string, item: NewSpecialItem) => Promise<void> | void;
}) {
  const [item, setItem] = useState<NewSpecialItem>({ kind, subject: '' });
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState('');

  const up = (patch: Partial<NewSpecialItem>) => setItem((i) => ({ ...i, ...patch }));

  // Compromissos, contatos e tarefas sempre vão para o arquivo de dados principal.
  const target = sources.find((s) => s.id === primarySourceId);
  const targetId = target?.id || '';

  const save = async () => {
    if (!item.subject.trim() && !(kind === 'contact' && item.contactName?.trim())) {
      setErr(kind === 'contact' ? 'Informe o nome do contato.' : 'Informe o assunto.');
      return;
    }
    if (!targetId) {
      setErr(
        'O arquivo de dados principal não está aberto. Abra-o ou defina outro em Configurações → Arquivos de Dados.',
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
            {TITLES[kind]}
          </strong>
          <button className="set-close" onClick={onClose} title="Fechar">
            <Icon name="close" size={15} />
          </button>
        </div>

        <F label="Arquivo de dados">
          <span style={{ fontSize: 12, color: target ? 'var(--text-dim)' : '#a4262c' }}>
            {target
              ? `${target.label} (principal)`
              : 'Arquivo de dados principal não está aberto'}
          </span>
        </F>

        {kind === 'contact' ? (
          <>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 10 }}>
              <F label="Nome completo">
                <input
                  className="inp"
                  autoFocus
                  value={item.contactName || ''}
                  onChange={(e) => up({ contactName: e.target.value })}
                />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <F label="Nome">
                <input className="inp" value={item.givenName || ''} onChange={(e) => up({ givenName: e.target.value })} />
              </F>
              <F label="Sobrenome">
                <input className="inp" value={item.surname || ''} onChange={(e) => up({ surname: e.target.value })} />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <F label="Email">
                <input className="inp" type="email" value={item.email || ''} onChange={(e) => up({ email: e.target.value })} />
              </F>
              <F label="Email 2">
                <input className="inp" type="email" value={item.email2 || ''} onChange={(e) => up({ email2: e.target.value })} />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label="Empresa">
                <input className="inp" value={item.company || ''} onChange={(e) => up({ company: e.target.value })} />
              </F>
              <F label="Cargo">
                <input className="inp" value={item.jobTitle || ''} onChange={(e) => up({ jobTitle: e.target.value })} />
              </F>
              <F label="Departamento">
                <input className="inp" value={item.dept || ''} onChange={(e) => up({ dept: e.target.value })} />
              </F>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label="Telefone comercial">
                <input className="inp" value={item.phone || ''} onChange={(e) => up({ phone: e.target.value })} />
              </F>
              <F label="Celular">
                <input className="inp" value={item.mobile || ''} onChange={(e) => up({ mobile: e.target.value })} />
              </F>
              <F label="Aniversário">
                <input className="inp" type="date" value={item.birthday || ''} onChange={(e) => up({ birthday: e.target.value })} />
              </F>
            </div>
          </>
        ) : kind === 'appt' ? (
          <>
            <F label="Assunto">
              <input className="inp" autoFocus value={item.subject} onChange={(e) => up({ subject: e.target.value })} />
            </F>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
              <F label="Local">
                <input className="inp" value={item.location || ''} onChange={(e) => up({ location: e.target.value })} />
              </F>
              <F label="Participantes (emails, separados por ;)">
                <input className="inp" value={item.attendees || ''} onChange={(e) => up({ attendees: e.target.value })} />
              </F>
            </div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={!!item.allDay}
                onChange={(e) => up({ allDay: e.target.checked })}
              />
              Evento de dia inteiro
            </label>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label="Início">
                <input
                  className="inp"
                  type={item.allDay ? 'date' : 'datetime-local'}
                  value={item.start || ''}
                  onChange={(e) => up({ start: e.target.value })}
                />
              </F>
              <F label="Término">
                <input
                  className="inp"
                  type={item.allDay ? 'date' : 'datetime-local'}
                  value={item.end || ''}
                  onChange={(e) => up({ end: e.target.value })}
                />
              </F>
              <F label="Lembrete">
                <select
                  className="inp"
                  value={item.reminderMin == null ? '' : String(item.reminderMin)}
                  onChange={(e) => up({ reminderMin: e.target.value === '' ? null : Number(e.target.value) })}
                >
                  <option value="">Sem lembrete</option>
                  <option value="0">Na hora</option>
                  <option value="5">5 minutos antes</option>
                  <option value="15">15 minutos antes</option>
                  <option value="30">30 minutos antes</option>
                  <option value="60">1 hora antes</option>
                  <option value="1440">1 dia antes</option>
                </select>
              </F>
            </div>
          </>
        ) : (
          <>
            <F label="Assunto">
              <input className="inp" autoFocus value={item.subject} onChange={(e) => up({ subject: e.target.value })} />
            </F>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 10 }}>
              <F label="Início">
                <input className="inp" type="date" value={item.startDate || ''} onChange={(e) => up({ startDate: e.target.value })} />
              </F>
              <F label="Vencimento">
                <input className="inp" type="date" value={item.dueDate || ''} onChange={(e) => up({ dueDate: e.target.value })} />
              </F>
              <F label="% concluída">
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
              Tarefa concluída
            </label>
          </>
        )}

        <F label={kind === 'contact' ? 'Anotações' : 'Descrição'}>
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
            {busy ? 'Salvando…' : 'Salvar'}
          </button>
          <button className="btn" disabled={busy} onClick={onClose}>
            Cancelar
          </button>
        </div>
      </div>
    </div>
  );
}
