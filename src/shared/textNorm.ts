/**
 * Normalização de texto para pesquisa — minúsculas e sem acentos/diacríticos
 * (NFD + strip de marcas), então "cafe" encontra "café". Usado pelos backends
 * de busca (arquivos de dados, POP3, IMAP) e pelo realce de termos na lista.
 */
export function normSearch(s: string): string {
  return s
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{M}/gu, '');
}
