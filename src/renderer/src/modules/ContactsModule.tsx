import React, { useState } from 'react';
import type { FolderNode, ItemSummary } from '../../../shared/types';

interface ContactsModuleProps {
  folder: FolderNode | null;
  items: ItemSummary[];
  selectedIndex: number | null;
  onSelect: (item: ItemSummary) => void;
}

function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  return (parts[0][0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}

export function ContactsModule({ folder, items, selectedIndex, onSelect }: ContactsModuleProps) {
  const [filter, setFilter] = useState('');
  const filtered = filter
    ? items.filter(
        (i) =>
          (i.contactDisplayName || i.subject || '').toLowerCase().includes(filter.toLowerCase()) ||
          (i.contactEmail || '').toLowerCase().includes(filter.toLowerCase()) ||
          (i.contactCompany || '').toLowerCase().includes(filter.toLowerCase()),
      )
    : items;

  const sorted = [...filtered].sort((a, b) =>
    (a.contactDisplayName || a.subject).localeCompare(b.contactDisplayName || b.subject, 'pt-BR'),
  );

  return (
    <div className="people-wrap">
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 }}>
        <div className="search-bar">
          <input
            type="search"
            placeholder="Pesquisar contatos…"
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
          />
          <span className="scope">
            {sorted.length} contato(s){folder ? ` em ${folder.name}` : ''}
          </span>
        </div>
        <div className="contact-grid">
          {sorted.length === 0 && <div className="empty-list">Nenhum contato nesta pasta</div>}
          {sorted.map((c) => (
            <div
              key={c.index}
              className={`contact-card ${selectedIndex === c.index ? 'selected' : ''}`}
              onClick={() => onSelect(c)}
            >
              <div className="contact-avatar">{initials(c.contactDisplayName || c.subject)}</div>
              <div style={{ minWidth: 0 }}>
                <div className="c-name">{c.contactDisplayName || c.subject}</div>
                {c.contactCompany && <div className="c-line">{c.contactCompany}</div>}
                {c.contactEmail && <div className="c-line">{c.contactEmail}</div>}
                {c.contactPhone && <div className="c-line">{c.contactPhone}</div>}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
