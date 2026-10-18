# Arquitetura do Unit Mail

Este documento descreve como o código está organizado e onde ficam os dados.
Para uso e recursos, veja o [README](../README.md).

## Processos do Electron

| Camada | Pasta | Responsabilidade |
|---|---|---|
| Main (Node.js) | `src/main/` | Janela, IPC, leitura de PST, protocolos de email, regras, sync, arquivos de dados |
| Preload | `src/preload/` | `contextBridge`: expõe `window.outlook` tipado (`OutlookApi`) ao renderer |
| Renderer (React) | `src/renderer/src/` | Interface: módulos, ribbon, listas, painel de leitura, diálogos |
| Compartilhado | `src/shared/types.ts` | Tipos usados pelos dois lados (`FolderNode`, `ItemSummary`, `MailSource`...) |

O renderer não tem acesso a Node.js (`contextIsolation: true`, `nodeIntegration: false`).
Toda operação com disco, rede ou criptografia passa por um handler IPC em
`src/main/index.ts`, com método correspondente em `src/preload/index.ts`.

## Fontes de email (`sourceId`)

Cada base aberta recebe um identificador com prefixo, que define a implementação:

| Prefixo | Tipo | Implementação |
|---|---|---|
| `pst:` | Arquivo PST (somente leitura) | `pstService.ts` + sombras `.eml` graváveis |
| `local:` | Base nativa (diretório de `.eml`) | `localStore.ts` / `pstService.ts` |
| `acct:` | Conta POP3/IMAP | `pop3Source.ts` / `imapSource.ts` |

`sourceManager.ts` cria as instâncias e roteia cada chamada pelo prefixo.
Todas implementam a interface `MailBackend` (`src/main/mailBackend.ts`):
árvore de pastas, contagem, leitura paginada, item, anexo, busca e fechamento.

## Modelo de dados

- **Mensagens nativas** são arquivos `.eml` dentro da pasta da base. Itens especiais
  (compromissos, contatos, tarefas) são `.eml` com headers `X-UM-*`; compromissos
  também levam um `.ics` real.
- **PST é somente leitura.** Alterações do usuário (lida/não lida, pastas ocultas,
  itens copiados ou movidos) ficam em sidecars dentro da pasta de dados.
- **Pastas-sombra:** pastas locais que espelham pastas de um PST aberto, fundidas na
  árvore pelo nome.

## Pasta de dados (perfil portátil)

A pasta de dados é escolhida pelo usuário na primeira execução. O `userData` do
Electron guarda só `profile.json`, o ponteiro para essa pasta
(`src/main/dataDir.ts`).

Arquivos principais dentro da pasta de dados:

| Arquivo | Conteúdo |
|---|---|
| `unitmail-profile.json` | Marca a pasta como perfil Unit Mail |
| `accounts.json` | Contas (servidores, portas, vínculos de entrega). Senhas cifradas via `safeStorage` |
| `settings.json` | Preferências, ordem das pastas, base principal |
| `rules.json` | Regras de recebimento |
| `signatures.json` | Assinaturas nomeadas |
| `contact-history.json` | Histórico de destinatários (autocompletar) |
| `localstores.json` / `pst-index.json` | Bases e PSTs registrados |
| `imap-state.json` | UIDs/UIDLs já baixados por conta |
| `pop3-cache-<conta>.json` | Cabeçalhos POP3 em cache |
| `Mail/` | Uma base por conta |

Dentro de cada base ou sidecar de PST:

| Arquivo | Conteúdo |
|---|---|
| `store.json` | Metadados da base |
| `readstate.json` | Estado de leitura |
| `hiddenfolders.json` | Pastas ocultadas pelo usuário |
| `sumindex.json` | Resumos já calculados (evita reparsear `.eml`/PST) |
| `searchindex.json` | Texto indexado para busca |

## Sincronização (Enviar/Receber)

`src/main/ruleEngine.ts`:

1. Para cada conta, as mensagens novas são lidas: POP3 pela sessão única UIDL/RETR,
   IMAP pelos UIDs acima do último processado.
2. Cada mensagem passa pelas regras do usuário; se nenhuma casar, vai para o destino
   padrão da conta.
3. Só depois de gravada a mensagem é marcada como processada (e, no POP3, só então
   pode ser removida do servidor). Falhas ficam para a próxima sincronização.
4. Gravações são deduplicadas por hash SHA-1 do `.eml`.

## Renderer

- `App.tsx` concentra o estado de navegação, a seleção, o cache de pastas por
  `sourceId|folderId` e os listeners de eventos do main.
- Listas usam virtualização (`MessageList.tsx`): só as linhas visíveis entram no DOM.
- Cores vêm de variáveis CSS em `styles.css` (tema claro/escuro). Componentes novos
  não devem usar cores fixas.

## Testes e verificação

```bash
npm run typecheck   # checagem TypeScript
npm test            # testes unitários (funções puras de regras e cabeçalhos)
npm run build       # gera out/
```
