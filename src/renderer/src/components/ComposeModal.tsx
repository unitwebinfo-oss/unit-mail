import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { Icon } from '../icons';
import { attachIcon } from './ReadingPane';
import type { AccountConfig, ComposeInput } from '../../../shared/types';

export interface ComposeDraft {
  accountId: string;
  /** Tipo — decide qual assinatura nomeada será usada */
  kind?: 'new' | 'reply' | 'forward';
  to: string;
  cc?: string;
  bcc?: string;
  subject: string;
  bodyText: string;
  /** Se já veio HTML (ex: resposta futura), usa no lugar do bodyText */
  bodyHtml?: string;
  /** Mensagem original citada (resposta/encaminhamento) — fica abaixo da assinatura */
  quoteHtml?: string;
  inReplyTo?: string;
  references?: string;
}

interface ComposeModalProps {
  draft: ComposeDraft;
  accounts: AccountConfig[];
  onClose: () => void;
  onSent: () => void;
}

const inputStyle: React.CSSProperties = {
  padding: '7px 10px',
  border: '1px solid var(--border)',
  borderRadius: 6,
  fontFamily: 'inherit',
  fontSize: 13,
};

function fileName(p: string): string {
  return p.split(/[\\/]/).pop() || p;
}

const EMAIL_RE = /^[^\s@;,<>]+@[^\s@;,<>]+\.[^\s@;,<>]+$/;

/** Fontes oferecidas na barra de formatação — conjunto seguro p/ email. */
const FONTS = [
  'Calibri',
  'Arial',
  'Segoe UI',
  'Tahoma',
  'Verdana',
  'Georgia',
  'Times New Roman',
  'Courier New',
];

/** Extrai o endereço de "Nome <email>" ou email puro. */
function addrOf(token: string): string {
  const m = token.match(/<([^>]+)>/);
  return (m ? m[1] : token).trim();
}

/** Divide a lista em endereços — aceita `;` e `,` como separador. */
function splitAddrs(raw: string): string[] {
  return raw
    .split(/[;,]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

/**
 * Normaliza a lista de destinatários: separadores viram `;`, remove vazios e
 * duplicados, e devolve os endereços com formato inválido para avisar.
 */
function normalizeAddrs(raw: string): { ok: string[]; bad: string[] } {
  const seen = new Set<string>();
  const ok: string[] = [];
  const bad: string[] = [];
  for (const t of splitAddrs(raw)) {
    const addr = addrOf(t);
    const norm = addr.replace(/\s+/g, '').toLowerCase();
    if (EMAIL_RE.test(addr)) {
      if (!seen.has(norm)) {
        seen.add(norm);
        ok.push(t);
      }
    } else {
      bad.push(t);
    }
  }
  return { ok, bad };
}

interface ContactSuggestion {
  name: string;
  email: string;
  count: number;
}

/**
 * Campo de destinatários com autocompletar do histórico de envios
 * (como o cache NK2 do Outlook) — separa vários endereços com `;` ou `,`.
 */
function AddrField({
  value,
  onChange,
  onBlurNormalize,
  placeholder,
}: {
  value: string;
  onChange: (v: string) => void;
  /** Chamado ao sair do campo com a lista de endereços inválidos. */
  onBlurNormalize: (bad: string[]) => void;
  placeholder?: string;
}) {
  const [sugs, setSugs] = useState<ContactSuggestion[]>([]);
  const [open, setOpen] = useState(false);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const wrapRef = useRef<HTMLDivElement>(null);

  // Último token depois do último `;`/`,` — o que está sendo digitado
  const currentToken = (v: string) => (v.split(/[;,]/).pop() || '').trim();

  const handleChange = (v: string) => {
    onChange(v);
    const tok = currentToken(v);
    if (timer.current) clearTimeout(timer.current);
    if (tok.length < 1) {
      setOpen(false);
      return;
    }
    timer.current = setTimeout(async () => {
      try {
        const list = (await window.outlook.suggestContacts?.(tok)) || [];
        setSugs(list);
        setOpen(list.length > 0);
      } catch {
        setOpen(false);
      }
    }, 180);
  };

  /** Troca o token em digitação pelo contato escolhido, já pronto p/ o próximo. */
  const pick = (c: ContactSuggestion) => {
    const idx = Math.max(value.lastIndexOf(';'), value.lastIndexOf(','));
    const base = idx >= 0 ? value.slice(0, idx + 1) : '';
    const display = c.name ? `"${c.name.replace(/"/g, '')}" <${c.email}>` : c.email;
    onChange(`${base}${base ? ' ' : ''}${display}; `);
    setOpen(false);
    setSugs([]);
  };

  const blur = () => {
    setOpen(false);
    if (!value.trim()) return;
    const { ok, bad } = normalizeAddrs(value);
    // Reescreve a lista normalizada — inválidos voltam no fim p/ correção
    onChange([...ok, ...bad].join('; '));
    onBlurNormalize(bad);
  };

  // Fecha a lista ao clicar fora
  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setOpen(false);
    };
    window.addEventListener('mousedown', close);
    return () => window.removeEventListener('mousedown', close);
  }, [open]);

  return (
    <div className="addr-wrap" ref={wrapRef}>
      <input
        style={inputStyle}
        value={value}
        placeholder={placeholder}
        onChange={(e) => handleChange(e.target.value)}
        onBlur={blur}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === 'Tab') {
            if (open && sugs[0]) {
              e.preventDefault();
              pick(sugs[0]);
            }
          } else if (e.key === 'Escape') {
            setOpen(false);
          }
        }}
      />
      {open && (
        <div className="addr-sug">
          {sugs.map((c) => (
            <div
              key={c.email}
              className="addr-sug-item"
              onMouseDown={(e) => {
                e.preventDefault(); // não dispara blur antes do pick
                pick(c);
              }}
            >
              <span className="sug-name">{c.name || c.email}</span>
              {c.name && <span className="sug-mail">{c.email}</span>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function textToHtml(text: string): string {
  return text
    .split('\n')
    .map((l) => l.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') || '&nbsp;')
    .join('<br>');
}

/** Botão da barra de formatação. */
function Tb({
  label,
  title,
  onClick,
  children,
}: {
  label?: string;
  title: string;
  onClick: () => void;
  children?: React.ReactNode;
}) {
  return (
    <button
      type="button"
      className="tb-btn"
      title={title}
      onMouseDown={(e) => {
        e.preventDefault(); // não perde a seleção do editor
        onClick();
      }}
    >
      {children ?? label}
    </button>
  );
}

export function ComposeModal({ draft, accounts, onClose, onSent }: ComposeModalProps) {
  const [accountId, setAccountId] = useState(draft.accountId || accounts[0]?.id || '');
  const [to, setTo] = useState(draft.to);
  const [cc, setCc] = useState(draft.cc || '');
  const [bcc, setBcc] = useState(draft.bcc || '');
  const [showCc, setShowCc] = useState(!!draft.cc);
  const [showBcc, setShowBcc] = useState(!!draft.bcc);
  const [full, setFull] = useState(false);
  const [subject, setSubject] = useState(draft.subject);
  const [attachPaths, setAttachPaths] = useState<string[]>([]);
  const [sending, setSending] = useState(false);
  const [error, setError] = useState('');
  const [badAddrs, setBadAddrs] = useState<string[]>([]);
  const editorRef = useRef<HTMLDivElement>(null);
  const sigReady = useRef(false);
  const kind = draft.kind || 'new';

  const exec = (cmd: string, value?: string) => {
    document.execCommand(cmd, false, value);
    editorRef.current?.focus();
  };

  // Conteúdo inicial + bloco reservado para a assinatura
  useLayoutEffect(() => {
    const ed = editorRef.current;
    if (!ed || sigReady.current) return;
    const bodyHtml = draft.bodyHtml ?? textToHtml(draft.bodyText);
    // Ordem estilo Outlook: texto novo → assinatura → mensagem original citada
    ed.innerHTML = `${bodyHtml}<div data-sig="1" class="sig-block"><br></div>${draft.quoteHtml || ''}`;
    sigReady.current = true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Assinatura da conta selecionada — aparece no corpo, como no Outlook
  useEffect(() => {
    if (!accountId || !sigReady.current) return;
    window.outlook.signatureFor?.(accountId, kind).then((sigHtml) => {
      const slot = editorRef.current?.querySelector('[data-sig]');
      if (slot) slot.innerHTML = sigHtml ? `<br>${sigHtml}` : '<br>';
    });
  }, [accountId, kind]);

  const pickFiles = async () => {
    const paths = await window.outlook.pickAttachments();
    if (paths.length) setAttachPaths((prev) => [...prev, ...paths]);
  };

  const insertImage = async () => {
    const dataUri = await window.outlook.pickImage?.();
    if (dataUri) exec('insertImage', dataUri);
  };

  // window.prompt() não existe no Electron — campo inline para a URL do link.
  const [linkMode, setLinkMode] = useState(false);
  const [linkUrl, setLinkUrl] = useState('https://');
  const savedRange = useRef<Range | null>(null);

  const insertLink = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount) savedRange.current = sel.getRangeAt(0).cloneRange();
    setLinkUrl('https://');
    setLinkMode(true);
  };

  const applyLink = () => {
    if (linkUrl && linkUrl !== 'https://') {
      const sel = window.getSelection();
      if (sel && savedRange.current) {
        sel.removeAllRanges();
        sel.addRange(savedRange.current);
      }
      exec('createLink', linkUrl);
    }
    setLinkMode(false);
  };

  const send = async () => {
    if (!accountId) {
      setError('Selecione uma conta de envio.');
      return;
    }
    if (!to.trim()) {
      setError('Informe ao menos um destinatário.');
      return;
    }
    // Valida as listas — inválidos são apontados antes de tentar enviar
    const all = [...normalizeAddrs(to).bad, ...normalizeAddrs(cc).bad, ...normalizeAddrs(bcc).bad];
    if (all.length) {
      setError(`Endereço(s) com formato inválido: ${all.join(', ')} — corrija antes de enviar.`);
      setBadAddrs(all);
      return;
    }
    setSending(true);
    setError('');
    try {
      const input: ComposeInput = {
        accountId,
        kind,
        to: normalizeAddrs(to).ok.join('; '),
        cc: normalizeAddrs(cc).ok.join('; ') || undefined,
        bcc: normalizeAddrs(bcc).ok.join('; ') || undefined,
        subject,
        html: editorRef.current?.innerHTML || '',
        inReplyTo: draft.inReplyTo,
        references: draft.references,
        attachmentPaths: attachPaths,
      };
      await window.outlook.sendMail(input);
      onSent();
    } catch (e) {
      setError(`Erro ao enviar: ${(e as Error).message}`);
    } finally {
      setSending(false);
    }
  };

  return (
    <div className="modal-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div
        className={`compose-window ${full ? 'full' : ''}`}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="compose-title">
          <h2>{kind === 'reply' ? 'Responder' : kind === 'forward' ? 'Encaminhar' : 'Novo Email'}</h2>
          <span style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
            <span
              className="close-x"
              title={full ? 'Restaurar tamanho' : 'Maximizar janela'}
              onClick={() => setFull((f) => !f)}
            >
              {full ? '🗗' : '🗖'}
            </span>
            <span className="close-x danger" title="Fechar" onClick={onClose}>
              <Icon name="close" size={16} />
            </span>
          </span>
        </div>
        <div className="compose-fields">
          <div className="field-row">
            <label>De</label>
            <select style={inputStyle} value={accountId} onChange={(e) => setAccountId(e.target.value)}>
              {accounts.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.label || a.email} ({a.email})
                </option>
              ))}
            </select>
          </div>
          <div className="field-row addr">
            <label>Para</label>
            <AddrField
              value={to}
              onChange={(v) => {
                setTo(v);
                setBadAddrs([]);
              }}
              onBlurNormalize={setBadAddrs}
              placeholder="email1@ex.com; email2@ex.com"
            />
            <span className="cc-toggles">
              {!showCc && (
                <button type="button" onClick={() => setShowCc(true)}>
                  Cc
                </button>
              )}
              {!showBcc && (
                <button type="button" onClick={() => setShowBcc(true)}>
                  Cco
                </button>
              )}
            </span>
          </div>
          {showCc && (
            <div className="field-row">
              <label>Cc</label>
              <AddrField
                value={cc}
                onChange={(v) => {
                  setCc(v);
                  setBadAddrs([]);
                }}
                onBlurNormalize={setBadAddrs}
              />
            </div>
          )}
          {showBcc && (
            <div className="field-row">
              <label>Cco</label>
              <AddrField
                value={bcc}
                onChange={(v) => {
                  setBcc(v);
                  setBadAddrs([]);
                }}
                onBlurNormalize={setBadAddrs}
              />
            </div>
          )}
          <div className="field-row">
            <label>Assunto</label>
            <input style={inputStyle} value={subject} onChange={(e) => setSubject(e.target.value)} />
          </div>
          {badAddrs.length > 0 && (
            <div className="addr-warn">
              Endereço(s) com formato inválido: <b>{badAddrs.join(', ')}</b> — corrija antes de enviar.
            </div>
          )}
          <div className="compose-toolbar">
            <Tb title="Negrito" onClick={() => exec('bold')}><b>B</b></Tb>
            <Tb title="Itálico" onClick={() => exec('italic')}><i>I</i></Tb>
            <Tb title="Sublinhado" onClick={() => exec('underline')}><u>U</u></Tb>
            <Tb title="Tachado" onClick={() => exec('strikeThrough')}><s>S</s></Tb>
            <span className="tb-sep" />
            <select
              className="tb-fontsize tb-fontname"
              title="Fonte"
              onMouseDown={(e) => e.stopPropagation()}
              onChange={(e) => {
                if (e.target.value) exec('fontName', e.target.value);
                e.target.value = '';
              }}
              defaultValue=""
            >
              <option value="" disabled>Fonte</option>
              {FONTS.map((f) => (
                <option key={f} value={f} style={{ fontFamily: f }}>
                  {f}
                </option>
              ))}
            </select>
            <select
              className="tb-fontsize"
              title="Tamanho da fonte"
              onMouseDown={(e) => e.stopPropagation()}
              onChange={(e) => {
                if (e.target.value) exec('fontSize', e.target.value);
                e.target.value = '';
              }}
              defaultValue=""
            >
              <option value="" disabled>Tamanho</option>
              <option value="2">Pequena</option>
              <option value="3">Normal</option>
              <option value="4">Média</option>
              <option value="5">Grande</option>
              <option value="6">Muito grande</option>
            </select>
            <input
              type="color"
              className="tb-color"
              title="Cor do texto"
              defaultValue="#000000"
              onChange={(e) => exec('foreColor', e.target.value)}
            />
            <Tb title="Marcador de cor do texto (realce)" onClick={() => exec('hiliteColor', '#ffff00')}>🖊</Tb>
            <span className="tb-sep" />
            <Tb title="Alinhar à esquerda" onClick={() => exec('justifyLeft')}>⇤</Tb>
            <Tb title="Centralizar" onClick={() => exec('justifyCenter')}>≡</Tb>
            <Tb title="Alinhar à direita" onClick={() => exec('justifyRight')}>⇥</Tb>
            <span className="tb-sep" />
            <Tb title="Lista com marcadores" onClick={() => exec('insertUnorderedList')}>•</Tb>
            <Tb title="Lista numerada" onClick={() => exec('insertOrderedList')}>1.</Tb>
            <Tb title="Diminuir recuo" onClick={() => exec('outdent')}>⇐</Tb>
            <Tb title="Aumentar recuo" onClick={() => exec('indent')}>⇒</Tb>
            <span className="tb-sep" />
            <Tb title="Inserir link" onClick={insertLink}>🔗</Tb>
            <Tb title="Inserir imagem" onClick={insertImage}>🖼</Tb>
            <Tb title="Limpar formatação" onClick={() => exec('removeFormat')}>⌫</Tb>
          </div>
          {linkMode && (
            <div className="img-toolbar">
              <span style={{ fontSize: 11, color: 'var(--text-dim)' }}>Link:</span>
              <input
                className="inp"
                style={{ flex: 1, minWidth: 200 }}
                autoFocus
                value={linkUrl}
                onChange={(e) => setLinkUrl(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') applyLink();
                  if (e.key === 'Escape') setLinkMode(false);
                }}
                placeholder="https://…"
              />
              <button className="btn btn-primary" onClick={applyLink}>Aplicar</button>
              <button className="btn btn-ghost" onClick={() => setLinkMode(false)}>Cancelar</button>
            </div>
          )}
          <div ref={editorRef} className="compose-editor" contentEditable suppressContentEditableWarning />
          <div className="compose-attach">
            <button onClick={pickFiles} className="attach-btn">
              <Icon name="clip" size={14} /> Anexar
            </button>
            {attachPaths.map((p, i) => (
              <span key={i} className="attach-chip" title={p}>
                <span className="attach-badge" style={{ color: attachIcon(fileName(p)).color }}>
                  <Icon name={attachIcon(fileName(p)).icon} size={16} />
                </span>
                {fileName(p)}
                <span
                  className="attach-x"
                  onClick={() => setAttachPaths((prev) => prev.filter((_, j) => j !== i))}
                >
                  ×
                </span>
              </span>
            ))}
          </div>
          {error && <div style={{ color: '#c00', fontSize: 12 }}>{error}</div>}
        </div>
        <div className="compose-footer">
          <button disabled={sending} onClick={send} className="primary-btn">
            {sending ? 'Enviando…' : 'Enviar'}
          </button>
          <button disabled={sending} onClick={onClose} style={inputStyle}>
            Cancelar
          </button>
        </div>
      </div>
    </div>
  );
}
