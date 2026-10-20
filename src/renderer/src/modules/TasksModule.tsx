import React from 'react';
import { fmtSmartDate } from '../utils';
import { t, useLang } from '../i18n';
import type { FolderNode, ItemSummary } from '../../../shared/types';

interface TasksModuleProps {
  folder: FolderNode | null;
  items: ItemSummary[];
  selectedIndex: number | null;
  onSelect: (item: ItemSummary) => void;
}

export function TasksModule({ folder, items, selectedIndex, onSelect }: TasksModuleProps) {
  useLang();
  const sorted = [...items].sort((a, b) => {
    if (!!a.isTaskComplete !== !!b.isTaskComplete) return a.isTaskComplete ? 1 : -1;
    const da = a.dueDate ? new Date(a.dueDate).getTime() : Infinity;
    const db = b.dueDate ? new Date(b.dueDate).getTime() : Infinity;
    return da - db;
  });

  return (
    <div className="tasks-wrap">
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
          >
            <span className="check">{task.isTaskComplete ? '✓' : ''}</span>
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
