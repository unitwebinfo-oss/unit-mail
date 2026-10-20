import { app } from 'electron';
import { existsSync, readFileSync } from 'fs';
import { join } from 'path';
import { dataDir } from './dataDir';
import type { Lang } from '../shared/folderNames';

/** Idioma do SO → idioma do app: pt-BR/pt-PT → 'pt'; qualquer outro → 'en'
 *  (os dois únicos idiomas suportados hoje). */
function systemLanguage(): Lang {
  try {
    return app.getLocale().toLowerCase().startsWith('pt') ? 'pt' : 'en';
  } catch {
    return 'pt';
  }
}

/**
 * Idioma da interface. Se o usuário já escolheu em Configurações, vale a
 * escolha salva em settings.json; caso contrário (primeiro uso), segue o
 * idioma do sistema operacional automaticamente.
 */
export function appLanguage(): Lang {
  try {
    const p = join(dataDir(), 'settings.json');
    if (existsSync(p)) {
      const s = JSON.parse(readFileSync(p, 'utf-8')) as { language?: string };
      if (s.language === 'en' || s.language === 'pt') return s.language;
    }
  } catch {
    // ignora
  }
  return systemLanguage();
}
