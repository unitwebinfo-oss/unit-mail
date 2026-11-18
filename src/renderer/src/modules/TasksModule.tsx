import React, { useEffect, useMemo, useState } from 'react';
import { fmtDateTime, fmtSmartDate, fmtTime } from '../utils';
import { locale, t, useLang } from '../i18n';
import { MiniMonth } from './CalendarModule';
import type { FolderNode, ItemSummary, ItemDetail, NewSpecialItem } from '../../../shared/types';

type TaskRef = { sourceId: string; folderId: string; index: number; stableId?: string };

interface TasksModuleProps {
  folder: FolderNode | null;
  sourceId: string | null;
  items: ItemSummary[];
  selectedIndex: number | null;
  onSelect: (item: ItemSummary) => void;
  onEditItem?: (ref: TaskRef, detail: ItemDetail) => void;
  onChanged?: () => void;
  /** Tarefa a abrir no inspetor (ex.: clique num lembrete) — consumida uma vez. */
  openDetail?: ItemDetail | null;
  onOpenDetailHandled?: () => void;
}

interface TaskGroup {
  id: 'overdue' | 'today' | 'upcoming' | 'nodate' | 'done';
  title: string;
  items: ItemSummary[];
}

// Mesma base do CalendarModule — semana no domingo, nomes via Intl no
// idioma da interface.
const DOW_BASE = new Date(2024, 0, 7);
function taskDowNames(): string[] {
  const f = new Intl.DateTimeFormat(locale(), { weekday: 'long' });
  return Array.from({ length: 7 }, (_, i) =>
    f.format(new Date(DOW_BASE.getFullYear(), DOW_BASE.getMonth(), DOW_BASE.getDate() + i)),
  ).map((n) => n.toUpperCase());
}
function taskMonthNames(): string[] {
  const f = new Intl.DateTimeFormat(locale(), { month: 'long' });
  return Array.from({ length: 12 }, (_, m) => f.format(new Date(2024, m, 1)));
}
function taskDayKey(d: Date): string {
  return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
}
function startOfMonth(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

/** Inspetor de tarefa — duplo clique abre, como o inspetor de compromisso. */
function TaskInspector({
  detail,
  busy,
  onClose,
  onToggle,
  onEdit,
}: {
  detail: ItemDetail;
  busy: boolean;
  onClose: () => void;
  onToggle: () => void;
  onEdit?: () => void;
}) {
  useLang();
  const td = detail.task;
  const done = !!detail.isTaskComplete;
  const pct = done ? 100 : td?.percentComplete ?? detail.percentComplete ?? 0;
  return (
    <div className="modal-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="appt-window" onClick={(e) => e.stopPropagation()}>
        <div className="compose-title">
          <h2>{detail.subject || t('Tarefa')}</h2>
          <span className="close-x danger" title={t('Fechar')} onClick={onClose}>
            ✕
          </span>
        </div>
        <div className="appt-body">
          <div className="appt-row">
            <span className="appt-k">{t('Início')}</span>
            <span>{td?.startDate ? fmtDateTime(td.startDate) : t('Sem data de início')}</span>
          </div>
          <div className="appt-row">
            <span className="appt-k">{t('Vencimento')}</span>
            <span>{td?.dueDate ? fmtDateTime(td.dueDate) : t('Sem vencimento')}</span>
          </div>
          <div className="appt-row">
            <span className="appt-k">{t('Situação')}</span>
            <span>{done ? t('Concluída') : t('{n}% concluída', { n: pct })}</span>
          </div>
          {td?.owner && (
            <div className="appt-row">
              <span className="appt-k">{t('Responsável')}</span>
              <span>{td.owner}</span>
            </div>
          )}
          <div className="task-progress wide">
            <i style={{ width: `${pct}%` }} />
          </div>
          {(td?.body || detail.body) && <pre className="task-inspect-body">{td?.body || detail.body}</pre>}
        </div>
        <div className="acct-foot">
          <button className="btn btn-primary" disabled={busy} onClick={onToggle}>
            {done ? t('Reabrir tarefa') : t('Marcar como concluída')}
          </button>
          {onEdit && (
            <button className="btn" disabled={busy} onClick={onEdit}>
              {t('Editar')}
            </button>
          )}
          <button className="btn" onClick={onClose}>
            {t('Fechar')}
          </button>
        </div>
      </div>
    </div>
  );
}

export function TasksModule({
  folder,
  sourceId,
  items,
  selectedIndex,
  onSelect,
  onEditItem,
  onChanged,
  openDetail,
  onOpenDetailHandled,
}: TasksModuleProps) {
  useLang();
  const [newSubject, setNewSubject] = useState('');
  const [busy, setBusy] = useState<number | null>(null);
  const [showDone, setShowDone] = useState(true);
  const [view, setView] = useState<'list' | 'calendar'>('list');
  const [viewDate, setViewDate] = useState(() => startOfMonth(new Date()));
  const [selectedDay, setSelectedDay] = useState(() => taskDayKey(new Date()));
  const [inspect, setInspect] = useState<{ ref: TaskRef; detail: ItemDetail } | null>(null);

  useEffect(() => {
    if (!openDetail || !sourceId || !folder) return;
    setInspect({
      ref: { sourceId, folderId: folder.id, index: openDetail.index, stableId: openDetail.stableId },
      detail: openDetail,
    });
    onOpenDetailHandled?.();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [openDetail]);

  const groups = useMemo<TaskGroup[]>(() => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const g: TaskGroup[] = [
      { id: 'overdue', title: t('Atrasadas'), items: [] },
      { id: 'today', title: t('Hoje'), items: [] },
      { id: 'upcoming', title: t('Próximas'), items: [] },
      { id: 'nodate', title: t('Sem vencimento'), items: [] },
      { id: 'done', title: t('Concluídas'), items: [] },
    ];
    const byDue = (a: ItemSummary, b: ItemSummary) =>
      (a.dueDate ? new Date(a.dueDate).getTime() : Infinity) -
      (b.dueDate ? new Date(b.dueDate).getTime() : Infinity);
    for (const task of items) {
      if (task.isTaskComplete) g[4].items.push(task);
      else if (!task.dueDate) g[3].items.push(task);
      else {
        const d = new Date(task.dueDate);
        if (d < today) g[0].items.push(task);
        else if (d < tomorrow) g[1].items.push(task);
        else g[2].items.push(task);
      }
    }
    for (const gr of g) gr.items.sort(byDue);
    return g;
  }, [items]);

  const pending = items.filter((i) => !i.isTaskComplete).length;
  const overdue = groups[0].items.length;

  // Tarefas com vencimento por dia — alimenta a grade mensal.
  const byDay = useMemo(() => {
    const map = new Map<string, ItemSummary[]>();
    for (const task of items) {
      if (!task.dueDate) continue;
      const k = taskDayKey(new Date(task.dueDate));
      const arr = map.get(k) || [];
      arr.push(task);
      map.set(k, arr);
    }
    return map;
  }, [items]);

  const calCells = useMemo(() => {
    const y = viewDate.getFullYear();
    const m = viewDate.getMonth();
    const offset = new Date(y, m, 1).getDay(); // domingo = 0
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    const cells: Array<{ date: Date; inMonth: boolean }> = [];
    for (let i = offset - 1; i >= 0; i--) cells.push({ date: new Date(y, m, -i), inMonth: false });
    for (let d = 1; d <= daysInMonth; d++) cells.push({ date: new Date(y, m, d), inMonth: true });
    let next = 1;
    while (cells.length % 7 !== 0) cells.push({ date: new Date(y, m + 1, next++), inMonth: false });
    return cells;
  }, [viewDate]);

  const navMonth = (dir: number) => setViewDate((d) => new Date(d.getFullYear(), d.getMonth() + dir, 1));
  const goToday = () => {
    const d = new Date();
    setSelectedDay(taskDayKey(d));
    setViewDate(startOfMonth(d));
  };

  const taskDayClass = (task: ItemSummary): string => {
    if (task.isTaskComplete) return 'done';
    if (!task.dueDate) return '';
    const d = new Date(task.dueDate);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (d < today) return 'overdue';
    if (taskDayKey(d) === taskDayKey(today)) return 'today';
    return '';
  };

  const refOf = (task: ItemSummary): TaskRef | null =>
    sourceId && folder ? { sourceId, folderId: folder.id, index: task.index, stableId: task.stableId } : null;

  const createTask = async () => {
    const subject = newSubject.trim();
    if (!subject || !sourceId) return;
    try {
      await window.outlook.createSpecialItem(sourceId, { kind: 'task', subject });
      setNewSubject('');
      onChanged?.();
    } catch (e) {
      console.error('createSpecialItem', e);
    }
  };

  /** Marca/desmarca conclusão preservando os demais campos da tarefa. */
  const setDone = async (task: ItemSummary, ref: TaskRef | null, done: boolean) => {
    if (!ref) return;
    setBusy(task.index);
    try {
      const detail = await window.outlook.getItem(ref.sourceId, ref.folderId, ref.index, ref.stableId);
      const td = detail.task;
      const item: NewSpecialItem = {
        kind: 'task',
        subject: detail.subject,
        body: td?.body ?? detail.body ?? '',
        startDate: td?.startDate ?? undefined,
        dueDate: td?.dueDate ?? undefined,
        percentComplete: done ? 100 : td?.percentComplete === 100 ? 0 : td?.percentComplete ?? 0,
        isComplete: done,
        owner: td?.owner ?? '',
      };
      await window.outlook.updateSpecialItem(ref.sourceId, ref.folderId, ref.index, item, ref.stableId);
      onChanged?.();
    } catch (err) {
      console.error('updateSpecialItem', err);
    } finally {
      setBusy(null);
    }
  };

  const toggleComplete = (task: ItemSummary, e: React.MouseEvent) => {
    e.stopPropagation();
    void setDone(task, refOf(task), !task.isTaskComplete);
  };

  /** Duplo clique: carrega a tarefa e abre o inspetor. */
  const openInspect = async (task: ItemSummary) => {
    const ref = refOf(task);
    if (!ref) return;
    try {
      const detail = await window.outlook.getItem(ref.sourceId, ref.folderId, ref.index, ref.stableId);
      setInspect({ ref, detail });
    } catch (err) {
      console.error('getItem', err);
    }
  };

  /** Após concluir pelo inspetor, recarrega a tarefa para o modal refletir o estado. */
  const toggleInspect = async () => {
    if (!inspect) return;
    const { ref, detail } = inspect;
    await setDone(detail, ref, !detail.isTaskComplete);
    try {
      const fresh = await window.outlook.getItem(ref.sourceId, ref.folderId, ref.index, ref.stableId);
      setInspect({ ref, detail: fresh });
    } catch {
      setInspect(null);
    }
  };

  const renderRow = (task: ItemSummary, group: TaskGroup['id']) => {
    const pct = task.percentComplete ?? 0;
    return (
      <div
        key={task.stableId ?? task.index}
        className={`task-row ${task.isTaskComplete ? 'done' : ''} ${selectedIndex === task.index ? 'selected' : ''}`}
        onClick={() => onSelect(task)}
        onDoubleClick={() => void openInspect(task)}
      >
        <button
          className="check"
          title={task.isTaskComplete ? t('Marcar como não concluída') : t('Marcar como concluída')}
          disabled={busy === task.index || !sourceId}
          onClick={(e) => toggleComplete(task, e)}
        >
          {task.isTaskComplete ? '✓' : ''}
        </button>
        <div className="t-body">
          <div className="t-subject">{task.subject}</div>
          <div className="t-due-line">
            <span className={`t-due ${group === 'overdue' ? 'overdue' : ''}`}>
              {task.dueDate ? t('Vence {date}', { date: fmtSmartDate(task.dueDate) }) : t('Sem vencimento')}
            </span>
            {group === 'overdue' && <span className="t-tag overdue">{t('atrasada')}</span>}
            {group === 'today' && <span className="t-tag today">{t('hoje')}</span>}
          </div>
          {pct > 0 && pct < 100 && (
            <div className="task-progress">
              <i style={{ width: `${pct}%` }} />
            </div>
          )}
        </div>
        <span className="t-pct">{pct}%</span>
      </div>
    );
  };

  const summary: Array<{ id: TaskGroup['id']; count: number }> = groups.map((g) => ({ id: g.id, count: g.items.length }));

  return (
    <div className="calendar-wrap tasks-full">
      <div className="cal-side left">
        {sourceId && (
          <div className="task-side-new">
            <input
              className="inp"
              placeholder={t('Nova tarefa (Enter)')}
              value={newSubject}
              onChange={(e) => setNewSubject(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && void createTask()}
            />
          </div>
        )}
        <div className="mini-nav">
          <button className="cal-nav" title={t('Mês anterior')} onClick={() => navMonth(-1)}>‹</button>
          <button className="cal-nav" title={t('Ir para hoje')} onClick={goToday}>{t('Hoje')}</button>
          <button className="cal-nav" title={t('Próximo mês')} onClick={() => navMonth(1)}>›</button>
        </div>
        <MiniMonth
          viewDate={viewDate}
          selectedDay={selectedDay}
          onPick={(d) => {
            setSelectedDay(taskDayKey(d));
            setViewDate(startOfMonth(d));
          }}
        />
        <div className="my-cals">
          <h4>{t('Resumo')}</h4>
          {summary.map((s) => {
            const g = groups.find((x) => x.id === s.id)!;
            return (
              <div key={s.id} className={`task-sum ${s.id}`}>
                <span>{g.title}</span>
                <b>{s.count}</b>
              </div>
            );
          })}
        </div>
      </div>
      <div className="cal-grid-area task-main">
        <div className="cal-header">
          {view === 'calendar' ? (
            <>
              <button className="cal-nav" onClick={() => navMonth(-1)}>‹</button>
              <button className="cal-nav" onClick={goToday} title={t('Hoje')}>●</button>
              <button className="cal-nav" onClick={() => navMonth(1)}>›</button>
              <h2>
                {taskMonthNames()[viewDate.getMonth()]} {viewDate.getFullYear()}
              </h2>
            </>
          ) : (
            <h2>
              {t('Tarefas')}
              <span className="task-head-count">
                {overdue > 0 && <span className="t-tag overdue">{t('{n} atrasada(s)', { n: overdue })}</span>}
                {items.length > 0 && <span>{t('{n} pendente(s)', { n: pending })}</span>}
              </span>
            </h2>
          )}
          <div className="task-view-toggle">
            <button className={view === 'list' ? 'active' : ''} onClick={() => setView('list')}>
              {t('Lista')}
            </button>
            <button className={view === 'calendar' ? 'active' : ''} onClick={() => setView('calendar')}>
              {t('Calendário')}
            </button>
          </div>
        </div>

        {view === 'calendar' ? (
          <div className="cal-grid">
            {taskDowNames().map((d) => (
              <div key={d} className="cal-dow">
                {d}
              </div>
            ))}
            {calCells.map((c, i) => {
              const k = taskDayKey(c.date);
              const dayTasks = byDay.get(k) || [];
              const isToday = k === taskDayKey(new Date());
              return (
                <div
                  key={i}
                  className={`cal-day ${c.inMonth ? '' : 'other-month'} ${isToday ? 'today' : ''} ${k === selectedDay ? 'selected' : ''}`}
                  onClick={() => setSelectedDay(k)}
                >
                  <span className="day-num">{c.date.getDate()}</span>
                  {dayTasks.slice(0, 4).map((task) => (
                    <div
                      key={task.stableId ?? task.index}
                      className={`cal-event task ${taskDayClass(task)}`}
                      title={task.subject}
                      onClick={(e) => {
                        e.stopPropagation();
                        setSelectedDay(k);
                        onSelect(task);
                      }}
                      onDoubleClick={(e) => {
                        e.stopPropagation();
                        void openInspect(task);
                      }}
                    >
                      {task.dueDate ? `${fmtTime(task.dueDate)} ` : ''}
                      {task.subject}
                    </div>
                  ))}
                  {dayTasks.length > 4 && (
                    <div style={{ color: 'var(--text-dim)', fontSize: 11 }}>
                      {t('+{n} mais', { n: dayTasks.length - 4 })}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        ) : (
          <div className="task-list">
            {items.length === 0 && (
              <div className="empty-list">
                {folder ? t('Nenhuma tarefa nesta pasta') : t('Selecione uma pasta de tarefas')}
              </div>
            )}
            {groups.map((gr) =>
              gr.items.length === 0 || (gr.id === 'done' && !showDone) ? null : (
                <div key={gr.id} className={`task-group ${gr.id}`}>
                  <div
                    className={`task-group-title ${gr.id === 'done' ? 'collapsible' : ''}`}
                    onClick={gr.id === 'done' ? () => setShowDone((v) => !v) : undefined}
                  >
                    {gr.id === 'done' && <span className="chevron">{showDone ? '▾' : '▸'}</span>}
                    {gr.title}
                    <span className="task-group-count">{gr.items.length}</span>
                  </div>
                  {gr.items.map((task) => renderRow(task, gr.id))}
                </div>
              ),
            )}
          </div>
        )}
      </div>
      {inspect && (
        <TaskInspector
          detail={inspect.detail}
          busy={busy !== null}
          onClose={() => setInspect(null)}
          onToggle={() => void toggleInspect()}
          onEdit={
            onEditItem
              ? () => {
                  const { ref, detail } = inspect;
                  setInspect(null);
                  onEditItem(ref, detail);
                }
              : undefined
          }
        />
      )}
    </div>
  );
}
