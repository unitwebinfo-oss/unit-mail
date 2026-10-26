import React, { useState } from 'react';
import { fmtSmartDate } from '../utils';
import { t, useLang } from '../i18n';
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
  const sorted = [...items].sort((a, b) => {
    if (!!a.isTaskComplete !== !!b.isTaskComplete) return a.isTaskComplete ? 1 : -1;
    const da = a.dueDate ? new Date(a.dueDate).getTime() : Infinity;
    const db = b.dueDate ? new Date(b.dueDate).getTime() : Infinity;
    return da - db;
  });
  const pending = items.filter((i) => !i.isTaskComplete).length;

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
          {items.length > 0 && t('{n} pendente(s)', { n: pending })}
        </span>
      </div>
      <div className="task-list">
        {sorted.length === 0 && (
          <div className="empty-list">
            {folder ? t('Nenhuma tarefa nesta pasta') : t('Selecione uma pasta de tarefas')}
          </div>
        )}
        {sorted.map((task) => (
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
              <div className="t-due">
                {task.dueDate ? t('Vence {date}', { date: fmtSmartDate(task.dueDate) }) : t('Sem vencimento')}
              </div>
            </div>
            <span className="t-pct">{task.percentComplete ?? 0}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}
