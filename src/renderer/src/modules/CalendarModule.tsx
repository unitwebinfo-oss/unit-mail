import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Icon } from '../icons';
import { fmtTime, fmtDateTime } from '../utils';
import { locale, t, useLang } from '../i18n';
import { Splitter } from '../components/Splitter';
import { externalizeLinks } from '../components/ReadingPane';
import type { FolderNode, ItemDetail, ItemSummary } from '../../../shared/types';

export interface CalendarRef {
  /** "sourceId|folderId" */
  key: string;
  sourceId: string;
  folderId: string;
  name: string;
  sourceLabel: string;
}

interface CalendarModuleProps {
  sourceId: string | null;
  folder: FolderNode | null;
  items: ItemSummary[];
  selectedIndex: number | null;
  onSelect: (item: ItemSummary) => void;
  /** Todas as pastas de calendário das bases abertas (para "Meus Calendários"). */
  calendars: CalendarRef[];
  /** Editar compromisso criado nesta base — abre o diálogo com os dados. */
  onEditItem?: (
    ref: { sourceId: string; folderId: string; index: number; stableId?: string },
    detail: ItemDetail,
  ) => void;
}

// Semana começa no domingo, como no Outlook. Nomes de dias/meses vêm do
// Intl no idioma da interface — segue a troca de idioma sem tabela fixa.
// 07/01/2024 era domingo — base para gerar os nomes na ordem certa.
const DOW_BASE = new Date(2024, 0, 7);
function dowNames(fmt: 'long' | 'narrow'): string[] {
  const f = new Intl.DateTimeFormat(locale(), { weekday: fmt });
  const names = Array.from({ length: 7 }, (_, i) =>
    f.format(new Date(DOW_BASE.getFullYear(), DOW_BASE.getMonth(), DOW_BASE.getDate() + i)),
  );
  return fmt === 'long' ? names.map((n) => n.toUpperCase()) : names;
}
function monthNames(): string[] {
  const f = new Intl.DateTimeFormat(locale(), { month: 'long' });
  return Array.from({ length: 12 }, (_, m) => f.format(new Date(2024, m, 1)));
}

function dayKey(d: Date): string {
  return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;
}

interface CalEvent extends ItemSummary {
  calKey: string;
  calSourceId: string;
  calFolderId: string;
}

/** Mini-calendário do painel esquerdo (mês inteiro, clique navega). */
function MiniMonth({
  viewDate,
  selectedDay,
  onPick,
}: {
  viewDate: Date;
  selectedDay: string;
  onPick: (d: Date) => void;
}) {
  useLang();
  const months = monthNames();
  const dowMini = dowNames('narrow');
  const y = viewDate.getFullYear();
  const m = viewDate.getMonth();
  const first = new Date(y, m, 1);
  const offset = first.getDay(); // domingo = 0
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const todayKey = dayKey(new Date());
  const cells: Array<Date | null> = [];
  for (let i = 0; i < offset; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(new Date(y, m, d));
  while (cells.length % 7 !== 0) cells.push(null);

  return (
    <div className="mini-month">
      <div className="mini-title">
        {months[m]} {y}
      </div>
      <div className="mini-grid">
        {dowMini.map((d, i) => (
          <span key={i} className="mini-dow">
            {d}
          </span>
        ))}
        {cells.map((c, i) => {
          if (!c) return <span key={i} />;
          const k = dayKey(c);
          return (
            <span
              key={i}
              className={`mini-day ${k === todayKey ? 'today' : ''} ${k === selectedDay ? 'sel' : ''}`}
              onClick={() => onPick(c)}
            >
              {c.getDate()}
            </span>
          );
        })}
      </div>
    </div>
  );
}

/** Inspector de compromisso — duplo clique abre a janela do item, como no Outlook. */
function AppointmentInspector({
  item,
  onClose,
  onEdit,
}: {
  item: ItemDetail;
  onClose: () => void;
  /** Presente quando o item é editável (criado nesta base — X-UM-*). */
  onEdit?: () => void;
}) {
  useLang();
  return (
    <div className="modal-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="appt-window" onClick={(e) => e.stopPropagation()}>
        <div className="compose-title">
          <h2>
            <Icon name="calendar" size={16} /> {item.subject || t('Compromisso')}
          </h2>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            {onEdit && (
              <button className="btn" onClick={onEdit} title={t('Editar')}>
                <Icon name="pen" size={13} /> {t('Editar')}
              </button>
            )}
            <span className="close-x danger" title={t('Fechar')} onClick={onClose}>
              <Icon name="close" size={16} />
            </span>
          </div>
        </div>
        <div className="appt-body">
          <div className="appt-row">
            <span className="appt-k">{t('Início')}</span>
            <span>{fmtDateTime(item.startTime)}</span>
          </div>
          <div className="appt-row">
            <span className="appt-k">{t('Término')}</span>
            <span>{fmtDateTime(item.endTime)}</span>
          </div>
          {item.location && (
            <div className="appt-row">
              <span className="appt-k">{t('Local')}</span>
              <span>{item.location}</span>
            </div>
          )}
          {item.organizer && (
            <div className="appt-row">
              <span className="appt-k">{t('Organizador')}</span>
              <span>{item.organizer}</span>
            </div>
          )}
          {item.requiredAttendees && (
            <div className="appt-row">
              <span className="appt-k">{t('Obrigatório')}</span>
              <span>{item.requiredAttendees}</span>
            </div>
          )}
          {item.allAttendees && (
            <div className="appt-row">
              <span className="appt-k">{t('Participantes')}</span>
              <span>{item.allAttendees}</span>
            </div>
          )}
          {item.isRecurring && (
            <div className="appt-row">
              <span className="appt-k">{t('Recorrência')}</span>
              <span>{item.recurrencePattern || t('Sim')}</span>
            </div>
          )}
          {(item.bodyHTML || item.body) && (
            <div className="appt-notes">
              {item.bodyHTML ? (
                <iframe
                  sandbox="allow-popups allow-popups-to-escape-sandbox"
                  title="notas"
                  srcDoc={externalizeLinks(item.bodyHTML)}
                  style={{ width: '100%', height: '100%', border: 0 }}
                />
              ) : (
                <pre>{item.body}</pre>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export function CalendarModule({
  sourceId,
  folder,
  items,
  selectedIndex,
  onSelect,
  calendars,
  onEditItem,
}: CalendarModuleProps) {
  useLang();
  const today = new Date();
  const [viewDate, setViewDate] = useState(new Date(today.getFullYear(), today.getMonth(), 1));
  const [selectedDay, setSelectedDay] = useState<string>(dayKey(today));
  const [sideW, setSideW] = useState(() => Number(localStorage.getItem('pane.calSideW')) || 240);
  const [openDetail, setOpenDetail] = useState<ItemDetail | null>(null);
  const [openRef, setOpenRef] = useState<{ sourceId: string; folderId: string; index: number; stableId?: string } | null>(null);
  const [search, setSearch] = useState('');
  const dragBase = useRef(0);

  // "Meus Calendários" — set de calendários habilitados (persistido)
  const [enabled, setEnabled] = useState<Set<string>>(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('cal.enabled') || 'null');
      if (Array.isArray(saved)) return new Set(saved);
    } catch {
      // ignora
    }
    return new Set<string>();
  });
  // Pasta selecionada entra habilitada por padrão
  useEffect(() => {
    if (!sourceId || !folder) return;
    const key = `${sourceId}|${folder.id}`;
    setEnabled((prev) => {
      if (prev.size > 0 || prev.has(key)) return prev;
      return new Set([key]);
    });
  }, [sourceId, folder]);

  const toggleCal = (key: string) => {
    setEnabled((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      localStorage.setItem('cal.enabled', JSON.stringify([...next]));
      return next;
    });
  };

  // Itens dos calendários habilitados (a pasta aberta usa os `items` já carregados)
  const [extraItems, setExtraItems] = useState<Record<string, CalEvent[]>>({});
  const enabledKey = [...enabled].sort().join(',');
  const selKey = sourceId && folder ? `${sourceId}|${folder.id}` : null;
  useEffect(() => {
    let alive = true;
    const load = async () => {
      const out: Record<string, CalEvent[]> = {};
      for (const cal of calendars) {
        if (!enabled.has(cal.key) || cal.key === selKey) continue;
        try {
          const its = await window.outlook.getItems(cal.sourceId, cal.folderId, 0, 500);
          out[cal.key] = its.map((it) => ({
            ...it,
            calKey: cal.key,
            calSourceId: cal.sourceId,
            calFolderId: cal.folderId,
          }));
        } catch {
          // calendário indisponível
        }
      }
      if (alive) setExtraItems(out);
    };
    void load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [enabledKey, selKey, calendars.length]);

  // Evento da pasta aberta (se habilitada) + dos demais calendários marcados
  const allEvents = useMemo<CalEvent[]>(() => {
    const evs: CalEvent[] = [];
    if (selKey && enabled.has(selKey) && sourceId && folder) {
      for (const it of items) {
        evs.push({ ...it, calKey: selKey, calSourceId: sourceId, calFolderId: folder.id });
      }
    }
    for (const k of enabled) {
      if (k === selKey) continue;
      evs.push(...(extraItems[k] || []));
    }
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      return evs.filter(
        (e) =>
          e.subject.toLowerCase().includes(q) ||
          (e.location || '').toLowerCase().includes(q),
      );
    }
    return evs;
  }, [items, extraItems, enabled, selKey, sourceId, folder, search]);

  const openItem = async (ev: CalEvent) => {
    onSelect(ev);
    try {
      const detail = await window.outlook.getItem(
        ev.calSourceId,
        ev.calFolderId,
        ev.index,
        ev.stableId,
      );
      setOpenRef({
        sourceId: ev.calSourceId,
        folderId: ev.calFolderId,
        index: ev.index,
        stableId: ev.stableId,
      });
      setOpenDetail(detail);
    } catch {
      // detalhe indisponível — a seleção no painel de leitura ainda funciona
    }
  };

  // Indexa compromissos por dia (eventos multi-dia aparecem em cada dia)
  const byDay = useMemo(() => {
    const map = new Map<string, CalEvent[]>();
    for (const it of allEvents) {
      if (!it.startTime) continue;
      const start = new Date(it.startTime);
      const end = it.endTime ? new Date(it.endTime) : start;
      const cur = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const last = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      let guard = 0;
      while (cur <= last && guard++ < 400) {
        const k = dayKey(cur);
        if (!map.has(k)) map.set(k, []);
        map.get(k)!.push(it);
        cur.setDate(cur.getDate() + 1);
      }
    }
    return map;
  }, [allEvents]);

  // Células do mês — semana começando no domingo
  const cells = useMemo(() => {
    const y = viewDate.getFullYear();
    const m = viewDate.getMonth();
    const first = new Date(y, m, 1);
    const offset = first.getDay(); // domingo = 0
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    const prevDays = new Date(y, m, 0).getDate();
    const total = Math.ceil((offset + daysInMonth) / 7) * 7;
    const arr: Array<{ date: Date; inMonth: boolean }> = [];
    for (let i = 0; i < total; i++) {
      const dayNum = i - offset + 1;
      if (dayNum < 1) {
        arr.push({ date: new Date(y, m - 1, prevDays + dayNum), inMonth: false });
      } else if (dayNum > daysInMonth) {
        arr.push({ date: new Date(y, m + 1, dayNum - daysInMonth), inMonth: false });
      } else {
        arr.push({ date: new Date(y, m, dayNum), inMonth: true });
      }
    }
    return arr;
  }, [viewDate]);

  const nav = (delta: number) => {
    setViewDate(new Date(viewDate.getFullYear(), viewDate.getMonth() + delta, 1));
  };

  const pickDay = (d: Date) => {
    setSelectedDay(dayKey(d));
    setViewDate(new Date(d.getFullYear(), d.getMonth(), 1));
  };

  const nextMonth = new Date(viewDate.getFullYear(), viewDate.getMonth() + 1, 1);

  return (
    <div className="calendar-wrap">
      {/* Painel esquerdo: mini-meses + Meus Calendários */}
      <div className="cal-side left" style={{ width: sideW }}>
        <div className="mini-nav">
          <button className="cal-nav" title={t('Mês anterior')} onClick={() => nav(-1)}>‹</button>
          <button
            className="cal-nav"
            title={t('Ir para hoje')}
            onClick={() => {
              setSelectedDay(dayKey(today));
              setViewDate(new Date(today.getFullYear(), today.getMonth(), 1));
            }}
          >
            {t('Hoje')}
          </button>
          <button className="cal-nav" title={t('Próximo mês')} onClick={() => nav(1)}>›</button>
        </div>
        <MiniMonth viewDate={viewDate} selectedDay={selectedDay} onPick={pickDay} />
        <MiniMonth viewDate={nextMonth} selectedDay={selectedDay} onPick={pickDay} />
        <div className="my-cals">
          <h4>{t('Meus Calendários')}</h4>
          {calendars.length === 0 && (
            <div className="empty-list" style={{ padding: 8 }}>
              {t('Nenhum calendário aberto')}
            </div>
          )}
          {calendars.map((c) => (
            <label key={c.key} className="my-cal" title={`${c.sourceLabel} — ${c.name}`}>
              <input
                type="checkbox"
                checked={enabled.has(c.key)}
                onChange={() => toggleCal(c.key)}
              />
              <span>
                {c.name}
                {calendars.filter((x) => x.name === c.name).length > 1 && (
                  <em> — {c.sourceLabel}</em>
                )}
              </span>
            </label>
          ))}
        </div>
      </div>
      <Splitter
        onDragStart={() => {
          dragBase.current = sideW;
        }}
        onDelta={(d) => {
          const v = Math.min(400, Math.max(170, dragBase.current + d));
          setSideW(v);
          localStorage.setItem('pane.calSideW', String(v));
        }}
      />

      {/* Grade mensal */}
      <div className="cal-grid-area">
        <div className="cal-header">
          <button className="cal-nav" onClick={() => nav(-1)}>
            ‹
          </button>
          <button className="cal-nav" onClick={() => nav(1)}>
            ›
          </button>
          <h2>
            {monthNames()[viewDate.getMonth()]} {viewDate.getFullYear()}
          </h2>
          <div className="cal-search">
            <Icon name="search" size={14} />
            <input
              placeholder={t('Pesquisar Calendário')}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
        </div>
        <div className="cal-grid">
          {dowNames('long').map((d) => (
            <div key={d} className="cal-dow">
              {d}
            </div>
          ))}
          {cells.map((c, i) => {
            const k = dayKey(c.date);
            const evs = byDay.get(k) || [];
            const isToday = k === dayKey(today);
            return (
              <div
                key={i}
                className={`cal-day ${c.inMonth ? '' : 'other-month'} ${isToday ? 'today' : ''} ${k === selectedDay ? 'selected' : ''}`}
                onClick={() => setSelectedDay(k)}
                onDoubleClick={() => {
                  /* novo compromisso — em breve */
                }}
              >
                <span className="day-num">{c.date.getDate()}</span>
                {evs.slice(0, 4).map((ev, j) => (
                  <div
                    key={`${ev.calKey}-${ev.index}-${j}`}
                    className="cal-event"
                    title={t('{subject} — duplo clique abre os detalhes', { subject: ev.subject })}
                    onClick={(e) => {
                      e.stopPropagation();
                      setSelectedDay(k);
                      onSelect(ev);
                    }}
                    onDoubleClick={(e) => {
                      e.stopPropagation();
                      void openItem(ev);
                    }}
                  >
                    {ev.startTime ? `${fmtTime(ev.startTime)} ` : ''}
                    {ev.subject}
                  </div>
                ))}
                {evs.length > 4 && (
                  <div style={{ color: 'var(--text-dim)', fontSize: 11 }}>
                    {t('+{n} mais', { n: evs.length - 4 })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
      {openDetail && (
        <AppointmentInspector
          item={openDetail}
          onClose={() => {
            setOpenDetail(null);
            setOpenRef(null);
          }}
          onEdit={
            openDetail.local && openRef && onEditItem
              ? () => {
                  const ref = openRef;
                  const d = openDetail;
                  setOpenDetail(null);
                  setOpenRef(null);
                  onEditItem(ref, d);
                }
              : undefined
          }
        />
      )}
    </div>
  );
}
