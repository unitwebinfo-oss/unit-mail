export function decodeMimeWords(str: string): string {
  return str.replace(/=\?([^?\s]+)\?([bBqQ])\?([^?]+)\?=/g, (_m, cs, enc, txt) => {
    let buf: Buffer;
    if (String(enc).toUpperCase() === 'B') {
      buf = Buffer.from(txt, 'base64');
    } else {
      const bytes: number[] = [];
      for (let i = 0; i < txt.length; i++) {
        if (txt[i] === '=' && i + 2 < txt.length) {
          bytes.push(parseInt(txt.slice(i + 1, i + 3), 16));
          i += 2;
        } else if (txt[i] === '_') {
          bytes.push(0x20);
        } else {
          bytes.push(txt.charCodeAt(i) & 0xff);
        }
      }
      buf = Buffer.from(bytes);
    }
    const charset = String(cs).toLowerCase();
    try {
      return new TextDecoder(charset === 'utf8' ? 'utf-8' : charset).decode(buf);
    } catch {
      return buf.toString('latin1');
    }
  });
}

/** Lê um cabeçalho do bloco de headers, desdobrando as linhas de continuação (RFC 5322). */
export function getHeader(raw: string, name: string): string {
  const head = raw.split(/\r?\n\r?\n/)[0].split(/\r?\n/);
  const key = `${name.toLowerCase()}:`;
  const parts: string[] = [];
  let on = false;
  for (const line of head) {
    if (on && /^[ \t]/.test(line)) {
      parts.push(line.trim());
      continue;
    }
    on = line.toLowerCase().startsWith(key);
    if (on) parts.push(line.slice(key.length).trim());
  }
  return parts.length ? decodeMimeWords(parts.join(' ')) : '';
}

export function parseFrom(raw: string): { name: string; addr: string } {
  const m = raw.match(/^(.*?)\s*<([^>]+)>\s*$/s);
  if (m) return { name: m[1].replace(/^"|"$/g, '').trim(), addr: m[2].trim() };
  return { name: raw.trim(), addr: raw.trim() };
}

/** Fusos textuais obsoletos do RFC 822 (ainda comuns em mensagens antigas/relé). */
const OBSOLETE_ZONE_MS: Record<string, number> = {
  UT: 0, UTC: 0, GMT: 0, Z: 0,
  EST: -5 * 3600e3, EDT: -4 * 3600e3,
  CST: -6 * 3600e3, CDT: -5 * 3600e3,
  MST: -7 * 3600e3, MDT: -6 * 3600e3,
  PST: -8 * 3600e3, PDT: -7 * 3600e3,
};
const MONTH_IDX: Record<string, number> = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
};

/**
 * Date: robusto — Date.parse recusa vários formatos reais (zona textual
 * obsoleta "EST", comentários "(GMT-03:00)", "Dec 02, 2024 10:00", data ISO
 * "2024-12-02 10:00:00", ano de 2 dígitos). Retorna null se nada servir.
 */
export function parseMailDate(raw: string | null | undefined): Date | null {
  if (!raw) return null;
  const s = String(raw)
    .replace(/\([^)]*\)/g, ' ') // comentários RFC "(BRT)", "(UTC-03:00)"
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return null;
  let t = Date.parse(s);
  if (!isNaN(t)) return new Date(t);
  // Zona obsoleta no fim ("-0300" vindo de relé antigo já é coberto pelo parse)
  const zm = /(\s+|^)(UT|UTC|GMT|Z|EST|EDT|CST|CDT|MST|MDT|PST|PDT)$/i.exec(s);
  if (zm) {
    const base = s.slice(0, zm.index).trim();
    t = Date.parse(base);
    if (!isNaN(t)) return new Date(t + (OBSOLETE_ZONE_MS[zm[2].toUpperCase()] ?? 0));
  }
  // Manual: "2 Dec 2024 10:00:00" / "Dec 2, 24 10:00" — separadores soltos,
  // ano de 2 dígitos, segundos e offset opcionais.
  const m =
    /(?:(\d{1,2})[\s.\-/]+([A-Za-z]{3,}|\d{1,2})[\s.\-/]+(\d{2,4})|([A-Za-z]{3,})[\s.\-/]+(\d{1,2}),?[\s]+(\d{2,4}))[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([+-]\d{2}:?\d{2})?/.exec(
      s,
    );
  if (m) {
    const monthTok = (m[2] ?? m[4] ?? '').toLowerCase().slice(0, 3);
    const mon = MONTH_IDX[monthTok] ?? (/^\d{1,2}$/.test(m[2] ?? '') ? +m[2] - 1 : NaN);
    const day = +(m[1] ?? m[5]);
    let yr = +(m[3] ?? m[6]);
    if (yr < 100) yr += yr < 70 ? 2000 : 1900;
    let offMin = 0;
    if (m[10]) {
      const sign = m[10].startsWith('-') ? -1 : 1;
      const digs = m[10].replace(/\D/g, '').padStart(4, '0');
      offMin = sign * (parseInt(digs.slice(0, 2), 10) * 60 + parseInt(digs.slice(2), 10));
    }
    const ms = Date.UTC(yr, mon, day, +m[7], +m[8], +(m[9] ?? 0));
    if (!isNaN(ms)) return new Date(ms - offMin * 60000);
  }
  // ISO sem 'T': "2024-12-02 10:00:00"
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(s);
  if (iso) {
    const ms = Date.UTC(+iso[1], +iso[2] - 1, +iso[3], +iso[4], +iso[5], +(iso[6] ?? 0));
    if (!isNaN(ms)) return new Date(ms);
  }
  return null;
}

/** Extrai a data de um cabeçalho Received: ("... ; Tue, 2 Dec 2024 10:00:00 -0300"). */
export function parseReceivedDate(received: string | null | undefined): Date | null {
  if (!received) return null;
  const tail = String(received).split(';').pop() ?? '';
  return parseMailDate(tail.trim());
}
