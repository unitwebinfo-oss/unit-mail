# Contribuindo com o Unit Mail

Obrigado pelo interesse! O projeto está em estágio inicial e toda ajuda é
bem-vinda.

## Como começar

```bash
git clone <url-do-repo>
cd unit-mail
npm install
npm run dev
```

O app abre em Electron com hot-reload. Se você não tem um arquivo `.pst`,
qualquer email exportado como `.eml` pode ser colocado numa "base de dados
local" criada pelo app para testes.

## Antes de enviar um PR

1. `npm run typecheck` — deve passar sem erros.
2. `npm run build` — deve produzir `out/` sem falhas.
3. Teste o fluxo afetado manualmente (abrir pasta, mover mensagem, sync).

## Convenções

- **Idioma**: código e comentários em português (pt-BR); UI em pt-BR.
- **Estilo**: React funcional + hooks; CSS por classes em `styles.css` com
  variáveis de tema (`var(--surface)`, `var(--text)`…) — nunca cores fixas em
  componentes novos, para o tema escuro funcionar.
- **Tipos** compartilhados entre main e renderer ficam em `src/shared/types.ts`.
- Novo IPC = handler em `src/main/index.ts` + método tipado em
  `src/preload/index.ts` (`OutlookApi`).

## O que NÃO fazer

- Não copiar código, ícones ou assets do Microsoft Outlook — a UI é inspirada
  no layout, mas implementada de forma independente.
- Não versionar arquivos `.pst`, `.eml` de usuário, senhas ou conteúdo de
  `userData` — já estão no `.gitignore`.
- Não contornar a natureza somente leitura do PST fingindo escrita — usar
  sempre o modelo pasta-sombra (ver `Docs/ARCHITECTURE.md`).

## Licença

Ao contribuir, você concorda que suas alterações são licenciadas sob a
**GPL-3.0**, mesma licença do projeto. Se criar um arquivo de código novo,
adicione o cabeçalho:

```ts
/*
 * Unit Mail — cliente de email para desktop com leitura de PST
 * Copyright (C) 2025  <seu nome>
 * SPDX-License-Identifier: GPL-3.0-only
 */
```

## Reportando bugs

Abra uma issue descrevendo: o que você esperava, o que aconteceu, passos para
reproduzir, e o sistema (Windows/Fedora). Não anexe `.pst` com dados
pessoais — se precisar, crie um arquivo de teste só com mensagens fictícias.
