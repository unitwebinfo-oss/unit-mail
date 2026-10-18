import React from 'react';
import { fmtSmartDate } from '../utils';
import type { FolderNode, ItemSummary } from '../../../shared/types';

interface TasksModuleProps {
  folder: FolderNode | null;
  items: ItemSummary[];
  selectedIndex: number | null;
  onSelect: (item: ItemSummary) => void;
}

export function TasksModule({ folder, items, selectedIndex, onSelect }: TasksModuleProps) {
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
            {folder ? 'Nenhuma tarefa nesta pasta' : 'Selecione uma pasta de tarefas'}
          </div>
        )}
        {sorted.map((t) => (
          <div
            key={t.index}
            className={`task-row ${t.isTaskComplete ? 'done' : ''} ${selectedIndex === t.index ? 'selected' : ''}`}
            onClick={() => onSelect(t)}
          >
            <span className="check">{t.isTaskComplete ? '✓' : ''}</span>
            <div className="t-body">
              <div className="t-subject">{t.subject}</div>
              <div className="t-due">
                {t.dueDate ? `Vence ${fmtSmartDate(t.dueDate)}` : 'Sem vencimento'}
              </div>
            </div>
            <span className="t-pct">{t.percentComplete ?? 0}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}
