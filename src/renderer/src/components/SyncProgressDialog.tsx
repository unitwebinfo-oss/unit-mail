import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from '../icons';
import type { SyncProgress } from '../../../shared/types';

interface Props {
  tasks: SyncProgress[];
  onClose: () => void;
  onCancelAll: () => void;
  /** Cancela só a tarefa selecionada (botão "Cancelar tarefa"). */
  onCancelTask: (taskId: string) => void;
}

/** Diálogo "Progresso de envio/recebimento" — estilo Outlook 2016: uma
 * linha por tarefa (conta), sem listar mensagem a mensagem. */
export function SyncProgressDialog({ tasks, onClose, onCancelAll, onCancelTask }: Props) {
  const [details, setDetails] = useState(true);
  const done = tasks.filter((t) => t.status !== 'processing' && t.status !== 'pending').length;
  const running = tasks.some((t) => t.status === 'processing' || t.status === 'pending');
  const errors = useMemo(() => tasks.filter((t) => t.status === 'error'), [tasks]);
  const [tab, setTab] = useState<'tasks' | 'errors'>('tasks');
  const [sel, setSel] = useState<string | null>(null);

  const selTask = tasks.find((t) => t.taskId === sel);
  const canCancelSel = selTask?.status === 'processing';

  // Fecha sozinho quando o envio/recebimento termina — mas só se não houve
  // erro (com erro o usuário precisa ver a aba Erros). Pequeno atraso para
  // mostrar o estado final "Concluído" antes de sumir.
  const wasRunning = useRef(running);
  useEffect(() => {
    if (running) {
      wasRunning.current = true;
      return;
    }
    if (!wasRunning.current || errors.length > 0) return;
    const t = setTimeout(onClose, 1500);
    return () => clearTimeout(t);
  }, [running, errors.length, onClose]);

  const statusLabel = (t: SyncProgress) =>
    t.status === 'done'
      ? 'Concluído'
      : t.status === 'cancelled'
        ? 'Cancelado'
        : t.status === 'error'
          ? 'Erro'
          : t.status === 'pending'
            ? 'Aguardando'
            : 'Processando';
  const remaining = (t: SyncProgress) =>
    t.status === 'processing' && t.total !== undefined
      ? `${Math.max(0, t.total - (t.done ?? 0))} item(ns)`
      : '';

  return (
    <div className="modal-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="sync-dialog" onClick={(e) => e.stopPropagation()}>
        <div className="sync-dlg-title">
          <Icon name="sendReceive" size={16} />
          <span>Progresso de envio/recebimento no Unit Mail</span>
          <button className="close-x danger" onClick={onClose} title="Fechar">
            <Icon name="close" size={14} />
          </button>
        </div>

        <div className="sync-dlg-count">
          {done} de {tasks.length || 0} tarefas foram concluídas
          {errors.length > 0 ? ` (${errors.length} com erro)` : ''}
        </div>
        <div className="sync-dlg-bar">
          <div
            className={`sync-dlg-fill ${errors.length ? 'has-error' : ''}`}
            style={{ width: tasks.length ? `${(done / tasks.length) * 100}%` : '0%' }}
          />
        </div>

        {details && (
          <>
            <div className="sync-dlg-tabs">
              <button
                className={tab === 'tasks' ? 'active' : ''}
                onClick={() => setTab('tasks')}
              >
                Tarefas
              </button>
              <button
                className={tab === 'errors' ? 'active' : ''}
                onClick={() => setTab('errors')}
              >
                Erros{errors.length ? ` (${errors.length})` : ''}
              </button>
            </div>

            {tab === 'tasks' && (
              <div className="sync-dlg-list">
                <div className="sync-dlg-head">
                  <span>Nome</span>
                  <span>Progresso</span>
                  <span>Restante</span>
                </div>
                {tasks.length === 0 && (
                  <div className="sync-dlg-empty">Nenhuma tarefa em execução.</div>
                )}
                {tasks.map((t) => (
                  <button
                    key={t.taskId}
                    className={`sync-dlg-row ${sel === t.taskId ? 'sel' : ''}`}
                    onClick={() => setSel(t.taskId)}
                    title={t.error || ''}
                  >
                    <span>
                      <Icon
                        name={t.status === 'done' ? 'check' : t.status === 'error' ? 'close' : 'sendReceive'}
                        size={13}
                      />{' '}
                      {t.label}
                    </span>
                    <span className={t.status === 'error' ? 'err' : ''}>{statusLabel(t)}</span>
                    <span>{remaining(t)}</span>
                  </button>
                ))}
              </div>
            )}
            {tab === 'errors' && (
              <div className="sync-dlg-list">
                {errors.length === 0 && (
                  <div className="sync-dlg-empty">Nenhum erro nesta execução.</div>
                )}
                {errors.map((t) => (
                  <div key={t.taskId} className="sync-dlg-error">
                    <b>{t.label}</b>
                    <div>{t.error}</div>
                  </div>
                ))}
              </div>
            )}
          </>
        )}

        <div className="sync-dlg-footer">
          <button
            className="ghost-btn"
            disabled={!canCancelSel}
            title={
              canCancelSel
                ? `Cancelar "${selTask?.label}"`
                : 'Selecione uma tarefa em execução na lista'
            }
            onClick={() => sel && onCancelTask(sel)}
          >
            Cancelar tarefa
          </button>
          <button className="ghost-btn" disabled={!running} onClick={onCancelAll}>
            Cancelar tudo
          </button>
          <button className="ghost-btn" onClick={() => setDetails((d) => !d)}>
            {details ? '>> Resumir' : '<< Detalhes'}
          </button>
          <button className="primary-btn" onClick={onClose}>
            {running ? 'Ocultar' : 'Fechar'}
          </button>
        </div>
      </div>
    </div>
  );
}
