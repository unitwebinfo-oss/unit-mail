import React from 'react';
import { Icon } from '../icons';
import type { ModuleId } from '../../../shared/types';

interface ModuleBarProps {
  module: ModuleId;
  onChange: (m: ModuleId) => void;
  /** Indicador Enviar/Receber (clicável → abre o diálogo de progresso). */
  syncStatus?: {
    active: boolean;
    doneTasks: number;
    totalTasks: number;
    onClick: () => void;
  };
}

const MODULES: Array<{ id: ModuleId; label: string; icon: string }> = [
  { id: 'mail', label: 'Email', icon: 'mail' },
  { id: 'calendar', label: 'Calendário', icon: 'calendar' },
  { id: 'people', label: 'Pessoas', icon: 'contact' },
  { id: 'tasks', label: 'Tarefas', icon: 'task' },
];

export function ModuleBar({ module, onChange, syncStatus }: ModuleBarProps) {
  return (
    <div className="module-bar">
      {MODULES.map((m) => (
        <button
          key={m.id}
          className={`module-btn ${module === m.id ? 'active' : ''}`}
          onClick={() => onChange(m.id)}
        >
          <Icon name={m.icon} />
          {m.label}
        </button>
      ))}
      {syncStatus && (
        <button
          className={`sync-status ${syncStatus.active ? 'running' : ''}`}
          onClick={syncStatus.onClick}
          title="Ver progresso de envio/recebimento"
        >
          <Icon name="sendReceive" size={13} />
          <span>Enviar/receber</span>
          <span className="sync-mini-bar">
            <span
              className="sync-mini-fill"
              style={{
                width: syncStatus.totalTasks
                  ? `${(syncStatus.doneTasks / syncStatus.totalTasks) * 100}%`
                  : '0%',
              }}
            />
          </span>
          <span className="sync-mini-count">
            {syncStatus.doneTasks}/{syncStatus.totalTasks}
          </span>
        </button>
      )}
    </div>
  );
}
