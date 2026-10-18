import React from 'react';

interface SplitterProps {
  /** true = divisor entre painéis lado a lado (arraste em X); false = empilhados (arraste em Y) */
  vertical?: boolean;
  /** Chamado no início do arraste — capture aqui o valor base */
  onDragStart?: () => void;
  /** Chamado a cada movimento com o delta em px desde o início do arraste */
  onDelta: (delta: number) => void;
}

/** Divisor arrastável entre painéis (estilo Outlook). */
export function Splitter({ vertical = true, onDragStart, onDelta }: SplitterProps) {
  const onMouseDown = (e: React.MouseEvent) => {
    e.preventDefault();
    onDragStart?.();
    const start = vertical ? e.clientX : e.clientY;
    const move = (ev: MouseEvent) => onDelta((vertical ? ev.clientX : ev.clientY) - start);
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  return <div className={vertical ? 'splitter-v' : 'splitter-h'} onMouseDown={onMouseDown} />;
}
