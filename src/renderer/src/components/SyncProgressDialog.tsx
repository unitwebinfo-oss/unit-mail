import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from '../icons';
import { t, useLang } from '../i18n';
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
  useLang();
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

  const statusLabel = (tk: SyncProgress) =>
    tk.status === 'done'
      ? t('Concluído')
      : tk.status === 'cancelled'
        ? t('Cancelado')
        : tk.status === 'error'
          ? t('Erro')
          : tk.status === 'pending'
            ? t('Aguardando')
            : t('Processando');
  const remaining = (tk: SyncProgress) =>
    tk.status === 'processing' && tk.total !== undefined
      ? t('{n} item(ns)', { n: Math.max(0, tk.total - (tk.done ?? 0)) })
      : '';

  return (
    <div className="modal-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="sync-dialog" onClick={(e) => e.stopPropagation()}>
        <div className="sync-dlg-title">
          <Icon name="sendReceive" size={16} />
          <span>{t('Progresso de envio/recebimento no Unit Mail')}</span>
          <button className="close-x danger" onClick={onClose} title={t('Fechar')}>
            <Icon name="close" size={14} />
          </button>
        </div>

        <div className="sync-dlg-count">
          {t('{done} de {total} tarefas foram concluídas', { done, total: tasks.length || 0 })}
          {errors.length > 0 ? ` (${t('{n} com erro', { n: errors.length })})` : ''}
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
                {t('Tarefas')}
              </button>
              <button
                className={tab === 'errors' ? 'active' : ''}
                onClick={() => setTab('errors')}
              >
                {t('Erros')}{errors.length ? ` (${errors.length})` : ''}
              </button>
            </div>

            {tab === 'tasks' && (
              <div className="sync-dlg-list">
                <div className="sync-dlg-head">
                  <span>{t('Nome')}</span>
                  <span>{t('Progresso')}</span>
                  <span>{t('Restante')}</span>
                </div>
                {tasks.length === 0 && (
                  <div className="sync-dlg-empty">{t('Nenhuma tarefa em execução.')}</div>
                )}
                {tasks.map((tk) => (
                  <button
                    key={tk.taskId}
                    className={`sync-dlg-row ${sel === tk.taskId ? 'sel' : ''}`}
                    onClick={() => setSel(tk.taskId)}
                    title={tk.error || ''}
                  >
                    <span>
                      <Icon
                        name={tk.status === 'done' ? 'check' : tk.status === 'error' ? 'close' : 'sendReceive'}
                        size={13}
                      />{' '}
                      {tk.label}
                    </span>
                    <span className={tk.status === 'error' ? 'err' : ''}>{statusLabel(tk)}</span>
                    <span>{remaining(tk)}</span>
                  </button>
                ))}
              </div>
            )}
            {tab === 'errors' && (
              <div className="sync-dlg-list">
                {errors.length === 0 && (
                  <div className="sync-dlg-empty">{t('Nenhum erro nesta execução.')}</div>
                )}
                {errors.map((tk) => (
                  <div key={tk.taskId} className="sync-dlg-error">
                    <b>{tk.label}</b>
                    <div>{tk.error}</div>
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
                ? t('Cancelar "{label}"', { label: selTask?.label || '' })
                : t('Selecione uma tarefa em execução na lista')
            }
            onClick={() => sel && onCancelTask(sel)}
          >
            {t('Cancelar tarefa')}
          </button>
          <button className="ghost-btn" disabled={!running} onClick={onCancelAll}>
            {t('Cancelar tudo')}
          </button>
          <button className="ghost-btn" onClick={() => setDetails((d) => !d)}>
            {details ? `>> ${t('Resumir')}` : `<< ${t('Detalhes')}`}
          </button>
          <button className="primary-btn" onClick={onClose}>
            {running ? t('Ocultar') : t('Fechar')}
          </button>
        </div>
      </div>
    </div>
  );
}
