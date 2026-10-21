import React from 'react';
import { Icon } from '../icons';
import { t, useLang } from '../i18n';

interface ConfirmDialogProps {
  title: string;
  message: string;
  /** Rótulo do botão de confirmação (padrão "Excluir"). */
  confirmLabel?: string;
  /** false = ação não destrutiva (botão azul em vez de vermelho). */
  danger?: boolean;
  /** Conteúdo extra entre a mensagem e os botões (ex.: checkbox de opção). */
  extra?: React.ReactNode;
  /** Ação alternativa não-destrutiva à esquerda (ex.: "Manter mensagens"). */
  secondaryAction?: { label: string; danger?: boolean; onConfirm: () => void };
  onConfirm: () => void;
  onCancel: () => void;
}

/** Diálogo "Tem certeza?" — confirmação antes de ações destrutivas. */
export function ConfirmDialog({
  title,
  message,
  confirmLabel = 'Excluir',
  danger = true,
  extra,
  secondaryAction,
  onConfirm,
  onCancel,
}: ConfirmDialogProps) {
  useLang();
  React.useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCancel();
      if (e.key === 'Enter') onConfirm();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onConfirm, onCancel]);

  return (
    <div className="modal-overlay" style={{ zIndex: 400 }} onMouseDown={(e) => e.target === e.currentTarget && onCancel()}>
      <div className="confirm-dialog" onClick={(e) => e.stopPropagation()}>
        <div className="confirm-title">
          <Icon name="junk" size={18} />
          <span>{title}</span>
        </div>
        <div className="confirm-msg">{message}</div>
        {extra}
        <div className="confirm-actions">
          {secondaryAction && (
            <button
              className={secondaryAction.danger ? 'danger-btn' : 'btn btn-primary'}
              style={{ marginRight: 'auto' }}
              onClick={secondaryAction.onConfirm}
            >
              {secondaryAction.label}
            </button>
          )}
          <button className="ghost-btn" onClick={onCancel} autoFocus>
            {t('Cancelar')}
          </button>
          <button className={danger ? 'danger-btn' : 'btn btn-primary'} onClick={onConfirm}>
            {confirmLabel === 'Excluir' ? t('Excluir') : confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
