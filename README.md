# Unit Mail

Cliente de email para desktop no estilo do Outlook 2016, feito para quem está
migrando do Windows para o Linux e **não quer perder o histórico de mensagens**:
o Unit Mail lê arquivos **PST** existentes diretamente e permite continuar
recebendo email por **POP3 e IMAP**, com envio via **SMTP**.

Interface em português (pt-BR), visual inspirado no Outlook 2016, código 100%
independente — nenhum arquivo ou asset da Microsoft é usado ou distribuído.

- **Código-fonte**: https://github.com/unitwebinfo-oss/unit-mail
- **Licença**: GPL-3.0 — livre para uso pessoal e comercial
- **Plataformas**: Windows 10/11 e Linux (Fedora/RPM, AppImage, tar.gz)

## Instalação

### Windows

Baixe o instalador **`Unit Mail Setup x.y.z.exe`** na página de
[Releases](https://github.com/unitwebinfo-oss/unit-mail/releases) e execute.
O assistente permite escolher a pasta de instalação e cria atalhos no menu
Iniciar e na área de trabalho.

Também existe a versão **portable** (`Unit Mail x.y.z.exe`) — roda direto,
sem instalar.

> O Windows SmartScreen pode exibir "O Windows protegeu seu computador" —
> clique em **"Mais informações" → "Executar mesmo assim"**. O aviso aparece
> porque o executável ainda não tem certificado de assinatura de código.

### Linux via GitHub Releases

Baixe o pacote da sua preferência em
[Releases](https://github.com/unitwebinfo-oss/unit-mail/releases):

```bash
# Ubuntu / Debian / Mint — pacote .deb:
sudo apt install ./unit-mail_0.1.0_amd64.deb     # ou: sudo dpkg -i ...

# Fedora / RHEL / openSUSE — pacote RPM:
sudo dnf install unit-mail-0.1.0.x86_64.rpm      # ou: sudo rpm -i ...

# Arch / Manjaro — pacote .pacman:
sudo pacman -U unit-mail-0.1.0.pacman

# AppImage — qualquer distribuição, sem instalar:
chmod +x "Unit Mail-0.1.0.AppImage"
./"Unit Mail-0.1.0.AppImage"

# Qualquer distro — tar.gz (extrai e roda):
tar -xzf unit-mail-0.1.0.tar.gz
./unit-mail-0.1.0/unit-mail
```

Os pacotes são gerados automaticamente pelo GitHub Actions a cada tag —
não precisa compilar na sua máquina.

### Linux compilando a partir do código

```bash
git clone https://github.com/unitwebinfo-oss/unit-mail.git
cd unit-mail
npm install
npm run dist:linux   # gera AppImage, .deb, .rpm, .pacman e tar.gz em release/
```

Para rodar direto sem empacotar: `npm start` (produção) ou `npm run dev`
(desenvolvimento com hot-reload).

### Primeira execução

Na primeira abertura, o assistente pergunta **onde guardar seus dados**
(padrão: `Documentos/Unit Mail`) e ajuda a configurar a primeira conta ou
abrir/importar um arquivo `.pst`. Todos os seus emails, contas e
configurações ficam nessa única pasta — backup é copiar ela inteira.

## Principais recursos

### Arquivos de dados

- Abertura de **múltiplos arquivos PST** simultâneos, com árvore de pastas
  igual ao Outlook (pastas internas/de sistema ficam ocultas).
- **Migração PST → base própria**: copia todas as pastas e mensagens de um
  `.pst` para uma base nativa do app (formato `.eml`) — depois disso, excluir
  apaga de verdade e libera espaço. O `.pst` original nunca é alterado.
- **Bases de dados locais** (diretórios `.eml`): equivalente gravável ao PST,
  uma base por conta — modelo do Thunderbird (`Mail/<conta>/`).
- **Mesclagem transparente**: mensagens entregues pela conta aparecem *dentro*
  das pastas reais do PST (ex.: "Caixa de Entrada"), sem pastas duplicadas.
- Arquivo de dados **Principal**: destino padrão de mensagens quando nenhuma
  regra casa — configurável em *Configurações → Arquivos de Dados*.
- Pastas: criar, renomear, mover (arrastar), copiar, excluir, esvaziar,
  reordenar, classificar A–Z, favoritos — tudo persistido entre sessões.
- Estado de leitura (lida/não lida) persistido para mensagens locais.

### Email

- Lista de mensagens compacta estilo Outlook: colunas De/Assunto/Recebido
  em/Tamanho, agrupamento por data (Hoje, Ontem, Esta semana…).
- **Multi-seleção** (Ctrl/Shift-clique, Ctrl+A), tecla Delete e
  **arrastar-e-soltar** de mensagens entre pastas (várias de uma vez).
- Painel de leitura com HTML renderizado, bloqueio de imagens externas
  (privacidade, como o Outlook), anexos com tipo e salvamento.
- Composição HTML: resposta cita a mensagem original abaixo do corpo,
  encaminhamento, anexos, imagem inline, assinatura por conta
  (nova mensagem vs resposta/encaminhamento).
- Envio SMTP com autenticação; senhas armazenadas de forma segura no
  `userData` (nunca saem do processo principal).

### Contas e regras

- Contas **POP3 e IMAP** (teste de conexão incluído), importação de contas do
  Outlook instalado no Windows — servidores, portas **e senhas salvas** são
  recuperados do Registro (DPAPI, mesma sessão de usuário).
- **Regras e alertas**: condições por remetente/destinatário/assunto, entrega
  em pasta de qualquer base, "parar de processar", "executar regras agora"
  sobre a pasta aberta.
- **Enviar/Receber** manual e automático com intervalo configurável;
  notificação do SO + aviso dentro do app quando chegam mensagens.

### Calendário

- Grade mensal (semana começando no domingo), mini-meses de navegação,
  "Meus Calendários" com checkbox por base, eventos como chips azuis.
- **Convites de reunião**: parse de iCalendar (.ics), barra de ação com
  **Aceitar / Recusar / Propor novo horário** (envia resposta `METHOD:REPLY` /
  `METHOD:COUNTER` ao organizador); aceitar grava o compromisso no calendário.
- **Lembretes** de compromissos com notificação do sistema.
- Inspector de compromisso: início/término, local, organizador, participantes,
  recorrência e notas.

### Interface

- Ribbon estilo Outlook 2016 com abas Página Inicial, Enviar/Receber, Pasta e
  backstage **Arquivo** (barra lateral azul).
- Painéis redimensionáveis (pastas, lista, leitura, calendário) e layout de
  leitura direita/inferior/desligado.
- **Tema claro, escuro ou "igual ao sistema"** (*Configurações → Geral*).
- Instância única: relançar o app foca a janela existente.

## Requisitos

- **Node.js 20+** e npm
- Windows 10/11 ou Linux (testado no Fedora; alvo de empacotamento: RPM/AppImage)

## Rodando em desenvolvimento

```bash
npm install
npm run dev        # abre o Electron com hot-reload (Vite)
```

## Build

```bash
npm run build      # compila main + preload + renderer para out/
npm start          # executa o build de produção
npm run typecheck  # checagem TypeScript sem emitir
npm test           # testes unitários (regras e cabeçalhos de email)
```

## Estrutura do projeto

```
src/
  main/        Processo principal do Electron (Node.js)
    index.ts         Janela, IPC, boot, sync automático, lembretes
    pstService.ts    Leitura de PST + pastas-sombra .eml graváveis
    sourceManager.ts Roteamento por sourceId (PST, local, conta)
    ruleEngine.ts    Regras de entrega + sincronização POP3/IMAP
    pop3Source.ts / pop3Client.ts / imapSource.ts   Protocolos
    smtpService.ts   Envio, testes de conexão, assinaturas, iCalendar
    accountStore.ts  Contas + senhas (persistência segura)
    dataDir.ts       Pasta de dados portátil (perfil estilo Thunderbird)
    pstImport.ts     Migração PST → base nativa do app
    signatureStore.ts / localStore.ts / paths.ts / outlookImportWin.ts
  preload/     contextBridge — superfície `window.outlook` tipada
  renderer/    React + TypeScript (Vite)
    src/App.tsx               Shell: módulos, seleção, drag&drop, toasts
    src/components/           Ribbon, FolderPane, MessageList, ReadingPane,
                              ComposeModal, AccountsSettings, OnboardingWizard…
    src/modules/              CalendarModule, ContactsModule…
    src/styles.css            Tema claro/escuro por variáveis CSS
  shared/      Tipos compartilhados entre main e renderer
scripts/       launch.js — orquestra dev/build/start
Docs/          Documentação de arquitetura (ARCHITECTURE.md)
```

## Onde ficam os dados — pasta portátil

Todos os dados ficam numa **única pasta escolhida por você** — na primeira
execução o assistente pergunta onde guardar (padrão: `Documentos/Unit Mail`;
alterável depois em *Configurações → Arquivos de Dados → Pasta de dados*).

```
Unit Mail/               ← a pasta que você escolheu (o "perfil")
├─ unitmail-profile.json ← marca que a pasta é um perfil Unit Mail
├─ accounts.json         ← contas (servidores, portas, vínculos de entrega)
├─ settings.json         ← preferências, ordem de pastas, base principal
├─ rules.json            ← regras de recebimento
├─ signatures.json       ← assinaturas nomeadas
├─ imap-state.json       ← mensagens já baixadas por conta (UIDLs POP3 / UIDs IMAP)
├─ pst-index.json        ← índice de PSTs já abertos (fingerprint)
└─ Mail/                 ← uma base por conta, como o Mail/ do Thunderbird
   └─ Emails — voce@…/
       ├─ store.json
       ├─ Caixa de Entrada/   ← mensagens .eml
       └─ Itens Enviados/
```

**Backup = copiar a pasta.** Para mudar de computador (Windows → Linux, por
exemplo): copie a pasta inteira e, na primeira execução do app na máquina nova,
aponte para ela em *"Onde guardar seus dados?"*. O app reconhece o perfil e
restaura contas, regras, bases e todo o email — sem recadastrar nada.

> **Senhas**: ficam criptografadas pelo sistema operacional (DPAPI no Windows,
> keyring no Linux) e **não são portáveis** — na máquina nova cada conta pede a
> senha uma vez. É o mesmo comportamento do Thunderbird.

O `userData` do Electron (`%APPDATA%/Unit Mail`, `~/.config/Unit Mail`) guarda
apenas o ponteiro para a pasta de dados e caches internos — não faz parte do
backup.

## Migrando do Outlook (Windows → Linux)

Fluxo recomendado para quem vem do Outlook instalado no Windows:

1. **No Windows**, rode o Unit Mail e use *Configurações → Email →
   "Procurar contas neste Windows"* — as contas do Outlook são lidas do
   Registro, incluindo servidores, portas e a senha salva. Opcionalmente
   exporte o arquivo `outlook-export.json` (sem senhas) para levar ao Linux.
2. **Copie os `.pst`** e a pasta de dados do Unit Mail para o Linux.
3. **No Linux**, na primeira execução aponte a pasta de dados copiada — as
   contas voltam prontas (só a senha é pedida de novo).
4. Para cada `.pst`, use *Configurações → Arquivos de Dados → "Migrar"* —
   todo o conteúdo vira uma base nativa gravável do app, e o `.pst` original
   permanece intacto como backup.

## Roadmap

- [x] Empacotamento para Linux (RPM, AppImage, tar.gz) e Windows (NSIS, portable)
- [ ] Novo compromisso / edição de compromissos pelo inspector
- [ ] Contatos e Tarefas com CRUD completo
- [ ] Pastas de pesquisa (search folders)
- [ ] Assinatura de código (remove o aviso do SmartScreen)

## Contribuindo

Veja [CONTRIBUTING.md](CONTRIBUTING.md) e a arquitetura interna em
[Docs/ARCHITECTURE.md](Docs/ARCHITECTURE.md).

## Licença

```
Unit Mail — cliente de email para desktop com leitura de PST
Copyright (C) 2025 Roger Oliveira

Este programa é software livre: você pode redistribuí-lo e/ou modificá-lo sob
os termos da Licença Pública Geral GNU conforme publicada pela Free Software
Foundation, na versão 3 da Licença.

Este programa é distribuído na esperança de ser útil, mas SEM NENHUMA GARANTIA;
sem mesmo a garantia implícita de COMERCIALIZAÇÃO ou ADEQUAÇÃO A UM PROPÓSITO
ESPECÍFICO. Veja a Licença Pública Geral GNU para mais detalhes.
```

O texto completo está em [LICENSE](LICENSE).

## Aviso legal

Unit Mail é um projeto independente, **não afiliado à Microsoft**. "Outlook" e
o formato PST são marcas/formatos da Microsoft Corporation; o acesso a arquivos
PST é feito por implementação de terceiros (`pst-extractor`), com fins de
interoperabilidade e migração de dados do próprio usuário. Nenhum código-fonte
ou asset da Microsoft é usado ou redistribuído.
