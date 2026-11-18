import React, { useMemo, useState } from 'react';
import { fmtSmartDate, fmtTime } from '../utils';
import { locale, t, useLang } from '../i18n';
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

export function TasksModule({
  folder,
  sourceId,
  items,
  selectedIndex,
  onSelect,
  onEditItem,
  onChanged,
}: TasksModuleProps) {
  useLang();
  const [newSubject, setNewSubject] = useState('');
  const [busy, setBusy] = useState<number | null>(null);
  const [showDone, setShowDone] = useState(true);
  const [view, setView] = useState<'list' | 'calendar'>('list');
  const [viewDate, setViewDate] = useState(() => {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth(), 1);
  });

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

  // Tarefas com vencimento por dia — alimenta a visão Calendário.
  const byDay = useMemo(() => {
    const map = new Map<string, ItemSummary[]>();
    for (const task of items) {
      if (!task.dueDate) continue;
      const d = new Date(task.dueDate);
      const k = taskDayKey(d);
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

  const navMonth = (dir: number) =>
    setViewDate((d) => new Date(d.getFullYear(), d.getMonth() + dir, 1));
  const goToday = () => {
    const d = new Date();
    setViewDate(new Date(d.getFullYear(), d.getMonth(), 1));
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

  const toggleComplete = async (task: ItemSummary, e: React.MouseEvent) => {
    e.stopPropagation();
    const ref = refOf(task);
    if (!ref) return;
    setBusy(task.index);
    try {
      const detail = await window.outlook.getItem(ref.sourceId, ref.folderId, ref.index, ref.stableId);
      const td = detail.task;
      const done = !task.isTaskComplete;
      const item: NewSpecialItem = {
        kind: 'task',
        subject: detail.subject,
        body: td?.body ?? detail.body ?? '',
        startDate: td?.startDate ?? undefined,
        dueDate: td?.dueDate ?? undefined,
        percentComplete: done ? 100 : (td?.percentComplete === 100 ? 0 : td?.percentComplete ?? 0),
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

  const openEdit = async (task: ItemSummary, e: React.MouseEvent) => {
    e.stopPropagation();
    const ref = refOf(task);
    if (!ref || !onEditItem) return;
    setBusy(task.index);
    try {
      const detail = await window.outlook.getItem(ref.sourceId, ref.folderId, ref.index, ref.stableId);
      onEditItem(ref, detail);
    } catch (err) {
      console.error('getItem', err);
    } finally {
      setBusy(null);
    }
  };

  const renderRow = (task: ItemSummary, group: TaskGroup['id']) => {
    const pct = task.percentComplete ?? 0;
    return (
      <div
        key={task.index}
        className={`task-row ${task.isTaskComplete ? 'done' : ''} ${selectedIndex === task.index ? 'selected' : ''}`}
        onClick={() => onSelect(task)}
        onDoubleClick={(e) => void openEdit(task, e)}
      >
        <button
          className="check"
          title={task.isTaskComplete ? t('Marcar como não concluída') : t('Marcar como concluída')}
          disabled={busy === task.index || !sourceId}
          onClick={(e) => void toggleComplete(task, e)}
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

  return (
    <div className="tasks-wrap">
      <div className="task-toolbar">
        {sourceId && (
          <input
            className="inp task-new"
            placeholder={t('Clique aqui para adicionar uma tarefa')}
            value={newSubject}
            onChange={(e) => setNewSubject(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && void createTask()}
          />
        )}
        <span className="task-count">
          {overdue > 0 && (
            <span className="t-tag overdue" style={{ marginRight: 8 }}>
              {t('{n} atrasada(s)', { n: overdue })}
            </span>
          )}
          {items.length > 0 && t('{n} pendente(s)', { n: pending })}
        </span>
        <div className="task-view-toggle">
          <button
            className={view === 'list' ? 'active' : ''}
            onClick={() => setView('list')}
            title={t('Lista')}
          >
            {t('Lista')}
          </button>
          <button
            className={view === 'calendar' ? 'active' : ''}
            onClick={() => setView('calendar')}
            title={t('Calendário')}
          >
            {t('Calendário')}
          </button>
        </div>
      </div>
      {view === 'calendar' ? (
        <div className="cal-grid-area task-cal-area">
          <div className="cal-header">
            <button className="cal-nav" onClick={() => navMonth(-1)}>
              ‹
            </button>
            <button className="cal-nav" onClick={goToday} title={t('Hoje')}>
              ●
            </button>
            <button className="cal-nav" onClick={() => navMonth(1)}>
              ›
            </button>
            <h2>
              {taskMonthNames()[viewDate.getMonth()]} {viewDate.getFullYear()}
            </h2>
          </div>
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
                  className={`cal-day ${c.inMonth ? '' : 'other-month'} ${isToday ? 'today' : ''}`}
                >
                  <span className="day-num">{c.date.getDate()}</span>
                  {dayTasks.slice(0, 4).map((task) => (
                    <div
                      key={task.stableId ?? task.index}
                      className={`cal-event task ${taskDayClass(task)}`}
                      title={task.subject}
                      onClick={(e) => {
                        e.stopPropagation();
                        onSelect(task);
                      }}
                      onDoubleClick={(e) => void openEdit(task, e)}
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
  );
}
