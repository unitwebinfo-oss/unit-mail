import { app } from 'electron';
import {
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from 'fs';
import { join } from 'path';

/**
 * Diretório de dados portátil — modelo "perfil" do Thunderbird.
 *
 * Tudo que é DADO do usuário (contas, regras, assinaturas, estado de sync,
 * registro de bases, configurações) fica numa pasta raiz que ele pode escolher
 * e copiar para outro computador — apontar o app para ela restaura tudo.
 *
 * Estrutura:
 *   <raiz>/
 *     unitmail-profile.json   marcador do perfil
 *     accounts.json, rules.json, settings.json, …
 *     Mail/<nome da base>/    bases locais (uma por conta)
 *
 * O userData do Electron guarda apenas o PONTEIRO (profile.json) e os caches
 * do Chromium — não fazem parte do backup.
 */

/** Arquivos de configuração que vivem na raiz de dados (migram juntos). */
const DATA_FILES = [
  'accounts.json',
  'rules.json',
  'signatures.json',
  'imap-state.json',
  'contact-history.json',
  'localstores.json',
  'settings.json',
  'pst-index.json',
  'categories.json', // lista-mestre de categorias coloridas
  'secrets.key', // chave AES das senhas — sem ela as contas perdem a senha na outra máquina
];

/** Padrões de arquivo copiados junto (caches por conta — sem eles o POP3
 *  baixaria tudo de novo na máquina nova). */
const DATA_FILE_GLOBS = [/^pop3-cache-.*\.json$/];

const DATA_DIRS = [
  'archives', // sidecars de reserva (PST em pasta não-gravável)
  'Mail', // bases locais do perfil — conteúdo de email propriamente dito
];

function pointerPath(): string {
  return join(app.getPath('userData'), 'profile.json');
}

let cached: string | null = null;

/**
 * Raiz de dados atual. Padrão: Documentos/Unit Mail — visível ao usuário e
 * fácil de fazer backup (diferente do %APPDATA% escondido).
 */
export function dataDir(): string {
  if (cached) return cached;
  try {
    const p = JSON.parse(readFileSync(pointerPath(), 'utf-8'));
    if (p?.dataDir && existsSync(p.dataDir)) {
      cached = p.dataDir as string;
      return cached;
    }
  } catch {
    // sem ponteiro ou corrompido — usa o padrão
  }
  cached = join(app.getPath('documents'), 'Unit Mail');
  return cached;
}

/** Pasta onde as bases locais ficam — <raiz>/Mail, como o Mail/ do Thunderbird. */
export function mailRoot(): string {
  const r = mailRootPath();
  mkdirSync(r, { recursive: true });
  return r;
}

/** Caminho de <raiz>/Mail SEM criar a pasta — para consultas que não podem
 *  materializar o perfil (ex.: app:info no primeiro boot, antes do wizard). */
export function mailRootPath(): string {
  return join(dataDir(), 'Mail');
}

/** O perfil já existe? (ponteiro gravado ou pasta de dados presente).
 *  No primeiro boot — antes do wizard escolher a pasta — é false e nada
 *  deve criar a raiz de dados "por acidente" (logs, app:info, caches). */
export function profileEstablished(): boolean {
  return existsSync(pointerPath()) || existsSync(dataDir());
}

/** O ponteiro do perfil já foi gravado? Significa que a pasta de dados já
 *  foi escolhida alguma vez — o setup inicial já passou por aqui, mesmo que
 *  a pasta esteja temporariamente indisponível (drive desmontado). */
export function profilePointerExists(): boolean {
  return existsSync(pointerPath());
}

/** Esta pasta já é um perfil Unit Mail? (tem marcador ou dados nossos) */
export function isUnitMailProfile(dir: string): boolean {
  return (
    existsSync(join(dir, 'unitmail-profile.json')) ||
    existsSync(join(dir, 'accounts.json')) ||
    existsSync(join(dir, 'settings.json')) ||
    existsSync(join(dir, 'Mail'))
  );
}

/**
 * Aponta o app para uma nova raiz de dados.
 * - Se a pasta já é um perfil (backup/outro PC), NADA é sobrescrito — o app
 *   adota o que existe lá.
 * - Se é uma pasta nova/vazia, os dados atuais são copiados para ela.
 * Requer reiniciar o app para todas as camadas relerem os arquivos.
 */
export function setDataDir(dir: string): { adopted: boolean } {
  mkdirSync(dir, { recursive: true });
  const adopted = isUnitMailProfile(dir);
  if (!adopted) {
    const cur = dataDir();
    for (const f of DATA_FILES) {
      const src = join(cur, f);
      if (existsSync(src) && !existsSync(join(dir, f))) {
        try {
          cpSync(src, join(dir, f));
        } catch {
          // segue — arquivo individual falhou
        }
      }
    }
    try {
      for (const f of readdirSync(cur)) {
        if (!DATA_FILE_GLOBS.some((re) => re.test(f))) continue;
        const src = join(cur, f);
        if (!existsSync(join(dir, f)) && statSync(src).isFile()) {
          cpSync(src, join(dir, f));
        }
      }
    } catch {
      // listagem falhou — caches serão reconstruídos no uso
    }
    for (const d of DATA_DIRS) {
      const src = join(cur, d);
      if (existsSync(src) && statSync(src).isDirectory() && !existsSync(join(dir, d))) {
        try {
          cpSync(src, join(dir, d), { recursive: true });
        } catch {
          // migração best-effort
        }
      }
    }
  }
  writeFileSync(
    join(dir, 'unitmail-profile.json'),
    JSON.stringify({ format: 'unitmail-profile@1', updated: new Date().toISOString() }, null, 2),
  );
  mkdirSync(app.getPath('userData'), { recursive: true });
  writeFileSync(pointerPath(), JSON.stringify({ dataDir: dir }, null, 2));
  cached = dir;
  return { adopted };
}

/**
 * Varre <raiz>/Mail procurando bases locais (subpastas com store.json).
 * Usado no boot para reconhecer bases trazidas de outro computador sem que o
 * usuário precise reimportar uma a uma.
 */
export function scanMailDir(): Array<{ name: string; path: string }> {
  const root = join(dataDir(), 'Mail');
  if (!existsSync(root)) return [];
  const found: Array<{ name: string; path: string }> = [];
  for (const entry of readdirSync(root)) {
    const p = join(root, entry);
    try {
      if (!statSync(p).isDirectory()) continue;
      const meta = join(p, 'store.json');
      let name = entry;
      if (existsSync(meta)) {
        name = JSON.parse(readFileSync(meta, 'utf-8')).name || entry;
      }
      found.push({ name, path: p });
    } catch {
      // pasta ilegível — pula
    }
  }
  return found;
}
