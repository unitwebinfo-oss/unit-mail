import React from 'react';
import { Icon } from '../icons';

interface ConfirmDialogProps {
  title: string;
  message: string;
  /** Rótulo do botão de confirmação (padrão "Excluir"). */
  confirmLabel?: string;
  onConfirm: () => void;
  onCancel: () => void;
}

/** Diálogo "Tem certeza?" — confirmação antes de ações destrutivas. */
export function ConfirmDialog({
  title,
  message,
  confirmLabel = 'Excluir',
  onConfirm,
  onCancel,
}: ConfirmDialogProps) {
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
        <div className="confirm-actions">
          <button className="ghost-btn" onClick={onCancel} autoFocus>
            Cancelar
          </button>
          <button className="danger-btn" onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
