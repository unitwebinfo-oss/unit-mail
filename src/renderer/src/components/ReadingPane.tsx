import React, { useMemo } from 'react';
import { Icon } from '../icons';
import { fmtFullDate, fmtDateTime, fmtSize } from '../utils';
import type { ItemDetail, MailCategory } from '../../../shared/types';

interface ReadingPaneProps {
  item: ItemDetail | null;
  loading: boolean;
  onSaveAttachment: (attachIndex: number) => void;
  /** Abrir anexo direto (cópia temporária, aplicativo padrão do SO). */
  onOpenAttachment?: (attachIndex: number) => void;
  /** "Salvar todos": grava todos os anexos numa pasta escolhida. */
  onSaveAllAttachments?: () => void;
  /** Baixar imagens externas no HTML dos emails (configuração de privacidade) */
  loadRemoteImages: boolean;
  /** Responder convite de reunião: aceitar, recusar ou propor novo horário. */
  onRespondInvite?: (
    response: 'accept' | 'decline' | 'counter',
    newStart?: string,
    newEnd?: string,
  ) => void;
  /** Ações do cabeçalho da mensagem (Responder/Encaminhar/…). */
  onReply?: () => void;
  onReplyAll?: () => void;
  onForward?: () => void;
  onDelete?: () => void;
  onPrint?: () => void;
  onSaveAs?: () => void;
  /** Lista-mestre de categorias — selos coloridos no cabeçalho. */
  categories?: MailCategory[];
}

/** Reescreve links do HTML para abrir fora do app: target=_blank + noopener.
 *  Dentro do iframe sandboxed o clique gera um popup que o main intercepta
 *  (setWindowOpenHandler) e manda para o navegador do SO / mailto no app.
 *  O primeiro atributo target vence — duplicatas originais são ignoradas. */
export function externalizeLinks(html: string): string {
  return html.replace(/<a(?=[\s>])/gi, '<a target="_blank" rel="noopener noreferrer"');
}

/** Esconde imagens remotas (src http/https e url()) — privacidade, como o Outlook bloqueia por padrão. */
function stripRemoteImages(html: string): { html: string; blocked: boolean } {
  let blocked = false;
  const out = html
    .replace(/(<img[^>]+src\s*=\s*["'])https?:\/\/[^"']*["']/gi, (m) => {
      blocked = true;
      return m.replace(/src\s*=\s*["']https?:\/\/[^"']*["']/i, 'src="about:blank" style="display:none"');
    })
    .replace(/url\(\s*['"]?https?:\/\/[^)'"]*['"]?\s*\)/gi, (m) => {
      blocked = true;
      return 'url(about:blank)';
    });
  return { html: out, blocked };
}

function MailBody({ item, loadRemoteImages }: { item: ItemDetail; loadRemoteImages: boolean }) {
  if (item.bodyHTML) {
    const { html, blocked } = loadRemoteImages
      ? { html: item.bodyHTML, blocked: false }
      : stripRemoteImages(item.bodyHTML);
    return (
      <>
        {blocked && (
          <div className="img-blocked-bar">
            Para proteger sua privacidade, as imagens externas desta mensagem foram bloqueadas.
            Ative em Configurações de Conta → Geral.
          </div>
        )}
        {/* iframe sandboxed: sem scripts; allow-popups permite links abrirem
            fora (o main nega a janela e abre no navegador do SO). */}
        <iframe
          sandbox="allow-popups allow-popups-to-escape-sandbox"
          title="email"
          srcDoc={externalizeLinks(html)}
        />
      </>
    );
  }
  return <pre>{item.body || '(sem conteúdo)'}</pre>;
}

/** Ícone + cor do anexo por extensão (estilo dos ícones de arquivo do SO). */
export function attachIcon(filename: string): { icon: string; color: string } {
  const ext = (filename.split('.').pop() || '').toLowerCase();
  const map: Record<string, { icon: string; color: string }> = {
    pdf: { icon: 'fileTxt', color: '#c9382c' },
    doc: { icon: 'fileTxt', color: '#2b579a' },
    docx: { icon: 'fileTxt', color: '#2b579a' },
    odt: { icon: 'fileTxt', color: '#2b579a' },
    rtf: { icon: 'fileTxt', color: '#2b579a' },
    xls: { icon: 'fileTxt', color: '#217346' },
    xlsx: { icon: 'fileTxt', color: '#217346' },
    csv: { icon: 'fileTxt', color: '#217346' },
    ods: { icon: 'fileTxt', color: '#217346' },
    ppt: { icon: 'fileTxt', color: '#d24726' },
    pptx: { icon: 'fileTxt', color: '#d24726' },
    odp: { icon: 'fileTxt', color: '#d24726' },
    txt: { icon: 'fileTxt', color: '#6b6b6b' },
    md: { icon: 'fileTxt', color: '#6b6b6b' },
    log: { icon: 'fileTxt', color: '#6b6b6b' },
    msg: { icon: 'fileMail', color: '#0078d4' },
    eml: { icon: 'fileMail', color: '#0078d4' },
    zip: { icon: 'fileZip', color: '#87643a' },
    rar: { icon: 'fileZip', color: '#87643a' },
    '7z': { icon: 'fileZip', color: '#87643a' },
    tar: { icon: 'fileZip', color: '#87643a' },
    gz: { icon: 'fileZip', color: '#87643a' },
    png: { icon: 'fileImg', color: '#7a5ea6' },
    jpg: { icon: 'fileImg', color: '#7a5ea6' },
    jpeg: { icon: 'fileImg', color: '#7a5ea6' },
    gif: { icon: 'fileImg', color: '#7a5ea6' },
    bmp: { icon: 'fileImg', color: '#7a5ea6' },
    svg: { icon: 'fileImg', color: '#7a5ea6' },
    webp: { icon: 'fileImg', color: '#7a5ea6' },
    ico: { icon: 'fileImg', color: '#7a5ea6' },
    mp3: { icon: 'fileAudio', color: '#d62598' },
    wav: { icon: 'fileAudio', color: '#d62598' },
    ogg: { icon: 'fileAudio', color: '#d62598' },
    flac: { icon: 'fileAudio', color: '#d62598' },
    mp4: { icon: 'fileVideo', color: '#1073c7' },
    mkv: { icon: 'fileVideo', color: '#1073c7' },
    avi: { icon: 'fileVideo', color: '#1073c7' },
    mov: { icon: 'fileVideo', color: '#1073c7' },
    webm: { icon: 'fileVideo', color: '#1073c7' },
  };
  return map[ext] || { icon: 'file', color: '#8a8a8a' };
}

function AppointmentDetail({ item, loadRemoteImages }: { item: ItemDetail; loadRemoteImages: boolean }) {
  return (
    <div className="detail-card">
      <h2>{item.subject}</h2>
      <div className="field">
        <span className="k">Início</span>
        <span>{fmtDateTime(item.startTime)}</span>
      </div>
      <div className="field">
        <span className="k">Término</span>
        <span>{fmtDateTime(item.endTime)}</span>
      </div>
      {item.location && (
        <div className="field">
          <span className="k">Local</span>
          <span>{item.location}</span>
        </div>
      )}
      {item.organizer && (
        <div className="field">
          <span className="k">Organizador</span>
          <span>{item.organizer}</span>
        </div>
      )}
      {item.requiredAttendees && (
        <div className="field">
          <span className="k">Participantes</span>
          <span>{item.requiredAttendees}</span>
        </div>
      )}
      {item.isRecurring && (
        <div className="field">
          <span className="k">Recorrência</span>
          <span>{item.recurrencePattern || 'Sim'}</span>
        </div>
      )}
      {(item.body || item.bodyHTML) && (
        <div style={{ marginTop: 14 }}>
          <MailBody item={item} loadRemoteImages={loadRemoteImages} />
        </div>
      )}
    </div>
  );
}

function ContactDetailView({ item }: { item: ItemDetail }) {
  const c = item.contact;
  if (!c) return <pre>{item.body}</pre>;
  const rows: Array<[string, string]> = [
    ['Nome', [c.displayName, `${c.givenName} ${c.middleName} ${c.surname}`.trim()].filter(Boolean)[0]],
    ['Empresa', c.company],
    ['Cargo', c.jobTitle],
    ['Departamento', c.department],
    ['Email', c.email1Address || c.email1],
    ['Email 2', c.email2Address],
    ['Email 3', c.email3Address],
    ['Tel. Comercial', c.businessPhone],
    ['Tel. Residencial', c.homePhone],
    ['Celular', c.mobilePhone],
    ['Endereço Comercial', c.businessAddress],
    ['Endereço Residencial', c.homeAddress],
    ['Aniversário', c.birthday ? fmtDateTime(c.birthday).split(' ')[0] : ''],
  ];
  return (
    <div className="detail-card">
      <h2>{c.displayName || item.subject}</h2>
      {rows
        .filter(([, v]) => v)
        .map(([k, v]) => (
          <div className="field" key={k}>
            <span className="k">{k}</span>
            <span>{v}</span>
          </div>
        ))}
      {c.notes && <div className="notes">{c.notes}</div>}
    </div>
  );
}

const TASK_STATUS = ['Não iniciada', 'Em andamento', 'Concluída', 'Aguardando', 'Adiada'];

function TaskDetailView({ item }: { item: ItemDetail }) {
  const t = item.task;
  if (!t) return <pre>{item.body}</pre>;
  return (
    <div className="detail-card">
      <h2>{item.subject}</h2>
      <div className="field">
        <span className="k">Status</span>
        <span>{TASK_STATUS[t.status] ?? t.status}</span>
      </div>
      <div className="field">
        <span className="k">% concluída</span>
        <span>{t.percentComplete}%</span>
      </div>
      <div className="field">
        <span className="k">Início</span>
        <span>{fmtDateTime(t.startDate)}</span>
      </div>
      <div className="field">
        <span className="k">Vencimento</span>
        <span>{fmtDateTime(t.dueDate)}</span>
      </div>
      {t.dateCompleted && (
        <div className="field">
          <span className="k">Concluída em</span>
          <span>{fmtDateTime(t.dateCompleted)}</span>
        </div>
      )}
      {t.owner && (
        <div className="field">
          <span className="k">Responsável</span>
          <span>{t.owner}</span>
        </div>
      )}
      {t.body && <div className="notes">{t.body}</div>}
    </div>
  );
}

/** Converte ISO → valor de input datetime-local. */
function toLocalInput(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
}

/** Barra de ações de convite de reunião (estilo Outlook). */
function InviteBar({
  item,
  onRespond,
}: {
  item: ItemDetail;
  onRespond?: ReadingPaneProps['onRespondInvite'];
}) {
  const [counter, setCounter] = React.useState(false);
  const [ns, setNs] = React.useState(toLocalInput(item.invite?.start));
  const [ne, setNe] = React.useState(toLocalInput(item.invite?.end));
  const [done, setDone] = React.useState('');
  const inv = item.invite!;

  if (done) {
    return <div className="invite-bar responded">✔ {done}</div>;
  }
  if (inv.method === 'CANCEL') {
    return (
      <div className="invite-bar cancelled">
        Esta reunião foi cancelada pelo organizador.
      </div>
    );
  }
  if (inv.method === 'REPLY') {
    return <div className="invite-bar">Esta mensagem é uma resposta a um convite de reunião.</div>;
  }
  return (
    <div className="invite-bar">
      <div className="invite-info">
        <Icon name="calendar" size={18} />
        <div>
          <strong>{inv.summary || item.subject}</strong>
          <span>
            {fmtDateTime(inv.start)}
            {inv.end ? ` – ${fmtDateTime(inv.end)}` : ''}
            {inv.location ? ` · ${inv.location}` : ''}
            {inv.organizer ? ` · Organizador: ${inv.organizer}` : ''}
          </span>
        </div>
      </div>
      <div className="invite-actions">
        <button
          className="accept"
          onClick={() => {
            onRespond?.('accept');
            setDone('Convite aceito');
          }}
        >
          Aceitar
        </button>
        <button
          onClick={() => {
            onRespond?.('decline');
            setDone('Convite recusado');
          }}
        >
          Recusar
        </button>
        <button onClick={() => setCounter((v) => !v)}>
          {counter ? 'Cancelar proposta' : 'Propor novo horário…'}
        </button>
      </div>
      {counter && (
        <div className="invite-counter">
          <label>
            Início
            <input
              type="datetime-local"
              value={ns}
              onChange={(e) => setNs(e.target.value)}
            />
          </label>
          <label>
            Término
            <input
              type="datetime-local"
              value={ne}
              onChange={(e) => setNe(e.target.value)}
            />
          </label>
          <button
            className="accept"
            disabled={!ns}
            onClick={() => {
              onRespond?.('counter', new Date(ns).toISOString(), ne ? new Date(ne).toISOString() : undefined);
              setDone('Nova proposta de horário enviada');
            }}
          >
            Enviar proposta
          </button>
        </div>
      )}
    </div>
  );
}

export function ReadingPane({
  item,
  loading,
  onSaveAttachment,
  onOpenAttachment,
  onSaveAllAttachments,
  loadRemoteImages,
  onRespondInvite,
  onReply,
  onReplyAll,
  onForward,
  onDelete,
  onPrint,
  onSaveAs,
  categories,
}: ReadingPaneProps) {
  const [attachMenu, setAttachMenu] = React.useState<number | null>(null);
  const [moreOpen, setMoreOpen] = React.useState(false);
  const catById = useMemo(
    () => new Map((categories ?? []).map((c) => [c.id, c])),
    [categories],
  );

  React.useEffect(() => {
    setAttachMenu(null);
    setMoreOpen(false);
  }, [item]);

  const isAppointment =
    item?.messageClass.startsWith('IPM.Appointment') ||
    item?.messageClass.startsWith('IPM.Schedule.Meeting');
  const isContact = item?.messageClass === 'IPM.Contact';
  const isTask = item?.messageClass === 'IPM.Task';
  const showActions = !!onReply && !isContact && !isTask;

  const openAttach = (attachIndex: number) => {
    setAttachMenu(null);
    if (onOpenAttachment) onOpenAttachment(attachIndex);
    else onSaveAttachment(attachIndex);
  };
  const saveAttach = (attachIndex: number) => {
    setAttachMenu(null);
    onSaveAttachment(attachIndex);
  };

  const header = useMemo(() => {
    if (!item) return null;
    return (
      <div className="rp-head-main">
        <div className="rp-subject">{item.subject}</div>
        {item.categories?.length ? (
          <div className="rp-cats">
            {item.categories.map((id) => {
              const c = catById.get(id);
              return c ? (
                <span key={id} className="cat-chip" style={{ borderColor: c.color, color: c.color }}>
                  <span className="cat-dot" style={{ background: c.color }} />
                  {c.name}
                </span>
              ) : null;
            })}
          </div>
        ) : null}
        {!isContact && !isTask && (
          <div className="rp-meta">
            <span className="lbl">De:</span>
            <span>
              {item.senderName}
              {item.senderEmailAddress ? ` <${item.senderEmailAddress}>` : ''}
            </span>
            {item.displayTo && (
              <>
                <span className="lbl">Para:</span>
                <span>{item.displayTo}</span>
              </>
            )}
            {item.displayCC && (
              <>
                <span className="lbl">Cc:</span>
                <span>{item.displayCC}</span>
              </>
            )}
            {isAppointment && (
              <>
                <span className="lbl">Início:</span>
                <span>
                  {fmtDateTime(item.startTime)} – {fmtDateTime(item.endTime)}
                </span>
              </>
            )}
            {!showActions && !isAppointment && (
              <>
                <span className="lbl">Data:</span>
                <span>{fmtFullDate(item.deliveryTime || item.sentTime)}</span>
              </>
            )}
          </div>
        )}
      </div>
    );
  }, [item, isContact, isTask, isAppointment, showActions, catById]);

  if (loading) {
    return (
      <div className="reading-pane">
        <div className="rp-empty">
          <div className="spinner" />
        </div>
      </div>
    );
  }

  if (!item) {
    return (
      <div className="reading-pane">
        <div className="rp-empty">
          <Icon name="envelope" />
          <div>Selecione um item para ler</div>
        </div>
      </div>
    );
  }

  return (
    <div className="reading-pane">
      <div className="rp-header">
        {header}
        {showActions && (
          <div className="rp-head-actions">
            <div className="rp-action-row">
              <button onClick={onReply} title="Responder">
                <Icon name="reply" size={14} />
                Responder
              </button>
              <button onClick={onReplyAll} title="Responder a todos">
                <Icon name="replyAll" size={14} />
                Responder a Todos
              </button>
              <button onClick={onForward} title="Encaminhar">
                <Icon name="forward" size={14} />
                Encaminhar
              </button>
              <span className="rp-more">
                <button
                  className="rp-more-btn"
                  title="Mais ações"
                  onClick={() => setMoreOpen((v) => !v)}
                >
                  …
                </button>
                {moreOpen && (
                  <>
                    <div className="menu-backdrop" onClick={() => setMoreOpen(false)} />
                    <div className="rp-more-menu">
                      {onSaveAs && (
                        <button
                          onClick={() => {
                            setMoreOpen(false);
                            onSaveAs();
                          }}
                        >
                          <Icon name="fileTxt" size={14} />
                          Salvar como…
                        </button>
                      )}
                      {onPrint && (
                        <button
                          onClick={() => {
                            setMoreOpen(false);
                            onPrint();
                          }}
                        >
                          <Icon name="printer" size={14} />
                          Imprimir…
                        </button>
                      )}
                      {onDelete && (
                        <button
                          className="danger"
                          onClick={() => {
                            setMoreOpen(false);
                            onDelete();
                          }}
                        >
                          <Icon name="trash" size={14} />
                          Excluir
                        </button>
                      )}
                    </div>
                  </>
                )}
              </span>
            </div>
            {!isAppointment && (
              <div className="rp-date">{fmtFullDate(item.deliveryTime || item.sentTime)}</div>
            )}
          </div>
        )}
      </div>
      {item.invite && <InviteBar item={item} onRespond={onRespondInvite} />}
      {item.attachments.length > 0 && (
        <div className="rp-attachments">
          {item.attachments.map((a) => {
            const badge = attachIcon(a.filename);
            return (
              <span key={a.index} className="attach-chip">
                <button
                  className="attach-main"
                  title="Abrir"
                  onClick={() => openAttach(a.index)}
                >
                  <span className="attach-badge" style={{ color: badge.color }}>
                    <Icon name={badge.icon} size={18} />
                  </span>
                  {a.filename}
                  {a.size > 0 && <span className="size">({fmtSize(a.size)})</span>}
                </button>
                <button
                  className="attach-caret"
                  title="Opções do anexo"
                  onClick={() => setAttachMenu(attachMenu === a.index ? null : a.index)}
                >
                  <Icon name="chevronD" size={12} />
                </button>
                {attachMenu === a.index && (
                  <>
                    <div className="menu-backdrop" onClick={() => setAttachMenu(null)} />
                    <span className="attach-menu">
                      <button onClick={() => openAttach(a.index)}>
                        <Icon name="open" size={14} />
                        Abrir
                      </button>
                      <button onClick={() => saveAttach(a.index)}>
                        <Icon name="download" size={14} />
                        Salvar como…
                      </button>
                    </span>
                  </>
                )}
              </span>
            );
          })}
          {item.attachments.length > 1 && onSaveAllAttachments && (
            <button className="attach-saveall" onClick={onSaveAllAttachments}>
              <Icon name="download" size={14} />
              Salvar todos
            </button>
          )}
        </div>
      )}
      {/* mail-html: o iframe do email ocupa o corpo todo e é o único que
          rola — evita um segundo scrollbar externo no painel/janela. */}
      <div className={`rp-body${!isAppointment && !isContact && !isTask && item.bodyHTML ? ' mail-html' : ''}`}>
        {isAppointment ? (
          <AppointmentDetail item={item} loadRemoteImages={loadRemoteImages} />
        ) : isContact ? (
          <ContactDetailView item={item} />
        ) : isTask ? (
          <TaskDetailView item={item} />
        ) : (
          <MailBody item={item} loadRemoteImages={loadRemoteImages} />
        )}
      </div>
    </div>
  );
}
