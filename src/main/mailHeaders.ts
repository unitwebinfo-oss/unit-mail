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
