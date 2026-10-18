/**
 * Nome seguro para arquivo/pasta tanto no Windows quanto no Linux:
 * remove os caracteres proibidos e de controle, corta ponto/espaco no fim
 * ou no inicio (o Win32 nao aceita — nomes assim viram fantasmas no NTFS
 * e fazem a copia falhar) e evita nomes reservados do DOS.
 */
export function safeFsName(name: string, fallback = 'Pasta'): string {
  // eslint-disable-next-line no-control-regex
  let s = name.replace(/[\\/:*?"<>|\x00-\x1f]/g, '');
  s = s.replace(/^[. ]+/, '').replace(/[. ]+$/, '');
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(s)) s = `_${s}`;
  return s || fallback;
}
