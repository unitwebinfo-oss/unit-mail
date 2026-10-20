<div align="center">

# Unit Mail

**Cliente de email para desktop no estilo do Outlook 2016 —<br/>
para quem migra do Windows para o Linux sem perder o histórico de emails em PST.**

[![Release](https://img.shields.io/github/v/release/unitwebinfo-oss/unit-mail?color=0173c7&label=release&logo=github)](https://github.com/unitwebinfo-oss/unit-mail/releases)
[![Status](https://img.shields.io/badge/status-beta-eaa300?labelColor=555)](https://github.com/unitwebinfo-oss/unit-mail/releases)
[![Licença](https://img.shields.io/badge/licen%C3%A7a-GPL--3.0-0173c7)](LICENSE)
[![Electron](https://img.shields.io/badge/Electron-47848f?logo=electron&logoColor=white)](https://www.electronjs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-3178c6?logo=typescript&logoColor=white)](https://www.typescriptlang.org)
[![Plataforma](https://img.shields.io/badge/plataforma-Linux%20%C2%B7%20Windows-2ea44f)](#-instalação)

[Download](#-instalação) ·
[Recursos](#-principais-recursos) ·
[Migração do Outlook](#-migrando-do-outlook-windows--linux) ·
[Desenvolvimento](#-rodando-em-desenvolvimento) ·
[Documentação](Docs/ARCHITECTURE.md)

</div>

---

> [!NOTE]
> **Versão Beta** — em fase de testes ativos. O app já verifica atualizações
> automaticamente e avisa quando uma release nova sair. Encontrou um problema?
> [Abra uma issue](https://github.com/unitwebinfo-oss/unit-mail/issues).

O Unit Mail lê arquivos **PST** existentes diretamente, continua recebendo
email por **POP3/IMAP** e envia via **SMTP** — com interface em português
(pt-BR) inspirada no Outlook 2016. Código 100% independente: nenhum arquivo
ou asset da Microsoft é usado ou distribuído.

---

## 📦 Instalação

### Linux — pacotes prontos (Releases)

Baixe em [**Releases**](https://github.com/unitwebinfo-oss/unit-mail/releases)
o formato da sua distribuição:

| Pacote | Distribuições | Comando |
|--------|---------------|---------|
| `.AppImage` | Qualquer distro (executável direto) | `chmod +x "Unit Mail-*.AppImage" && ./"Unit Mail-*.AppImage"` |
| `.deb` | Debian, Ubuntu, Mint e derivados | `sudo apt install ./unit-mail_*_amd64.deb` |
| `.rpm` | Fedora, RHEL, openSUSE | `sudo dnf install ./unit-mail-*.x86_64.rpm` |
| `.pacman` | Arch, Manjaro | `sudo pacman -U unit-mail-*.pacman` |
| `.tar.gz` | Qualquer distro (portátil) | `tar -xzf unit-mail-*.tar.gz` |

Os pacotes são gerados automaticamente pelo GitHub Actions a cada tag —
não precisa compilar na sua máquina.

### Windows — compilando a partir do código

As releases publicadas são Linux-only. Para Windows, compile localmente:

```bash
git clone https://github.com/unitwebinfo-oss/unit-mail.git
cd unit-mail
npm install
npm start          # roda o build de produção
```

### Compilando os pacotes Linux

```bash
npm install
npm run dist:linux   # gera AppImage, .deb, .rpm, .pacman e tar.gz em release/
```

### Primeira execução

Na primeira abertura, o assistente pergunta **onde guardar seus dados**
(padrão: `Documentos/Unit Mail`) e ajuda a configurar a primeira conta ou
abrir/importar um `.pst`. Todos os emails, contas e configurações ficam
nessa única pasta — **backup é copiar ela inteira**.

---

## ✨ Principais recursos

<table>
<tr>
<td width="50%" valign="top">

### 📁 Arquivos de dados

- Múltiplos **PST** abertos ao mesmo tempo, com árvore igual ao Outlook
- **Migração PST → base nativa** com deduplicação — reimportar não duplica
- Bases locais `.eml` por conta (modelo Thunderbird `Mail/`)
- Entrega **mesclada** nas pastas reais do PST — sem duplicatas
- Arquivo de dados **Principal** configurável
- Criar, renomear, mover, copiar, excluir, favoritos — tudo persistido

### ✉️ Email

- Lista estilo Outlook + agrupamento por data
- Multi-seleção, Delete e **arrastar-e-soltar** entre pastas
- Bloqueio de imagens externas (privacidade)
- Composição em **janela separada** — minimiza e volta quando quiser
- Resposta cita o original; anexos, imagem inline, assinatura por conta

</td>
<td width="50%" valign="top">

### 🔀 Contas, regras e proteção

- Contas **POP3 e IMAP** + importação do Outlook (servidores **e senhas**)
- **Regras e alertas**: condições, destino em qualquer base, execução manual
- **Filtro de spam**: marcações do servidor (SpamAssassin/cPanel) vão direto
  para **Lixo eletrônico**
- **Antivírus** (Windows): mensagens infectadas isoladas em
  **"Itens infectados"** — pasta protegida contra exclusão
- Enviar/Receber automático + notificações do SO

### 📅 Calendário e interface

- Grade mensal, mini-meses, múltiplos calendários por base
- **Convites iCalendar**: aceitar / recusar / propor horário
- Lembretes com notificação do sistema
- Ribbon minimalista com ícones coloridos, tema claro/escuro/sistema
- **Verificação de atualização** automática na abertura

</td>
</tr>
</table>

---

## 📂 Onde ficam os dados — pasta portátil

Todos os dados ficam numa **única pasta escolhida por você**:

```
Unit Mail/                ← a pasta que você escolheu (o "perfil")
├─ unitmail-profile.json  ← marca que a pasta é um perfil Unit Mail
├─ accounts.json          ← contas (servidores, portas, vínculos de entrega)
├─ settings.json          ← preferências, tema, janela, base principal
├─ rules.json             ← regras de recebimento
├─ signatures.json        ← assinaturas nomeadas
├─ imap-state.json        ← mensagens já baixadas (UIDLs POP3 / UIDs IMAP)
├─ pst-index.json         ← índice de PSTs já abertos (fingerprint)
└─ Mail/                  ← uma base por conta, como o Mail/ do Thunderbird
   └─ Emails — voce@…/
       ├─ Caixa de Entrada/   ← mensagens .eml
       ├─ Lixo eletrônico/    ← spam detectado pelo servidor
       ├─ Itens infectados/   ← quarentena do antivírus (pasta protegida)
       └─ Itens Enviados/
```

**Backup = copiar a pasta.** Para mudar de máquina (Windows → Linux): copie
a pasta inteira e aponte para ela na primeira execução — o app reconhece o
perfil e restaura contas, regras, bases e todo o email.

> [!IMPORTANT]
> **Senhas** ficam criptografadas pelo SO (DPAPI no Windows, keyring no
> Linux) e **não são portáveis** — na máquina nova cada conta pede a senha
> uma vez. Mesmo comportamento do Thunderbird.

---

## 🔄 Migrando do Outlook (Windows → Linux)

1. **No Windows** — rode o Unit Mail e use *Configurações → Email →
   "Procurar contas neste Windows"*: contas do Outlook são lidas do
   Registro, incluindo servidores, portas e senha salva. Opcionalmente
   exporte `outlook-export.json` (sem senhas) para levar ao Linux.
2. **Copie** os `.pst` e a pasta de dados do Unit Mail para o Linux.
3. **No Linux** — na primeira execução aponte a pasta copiada: as contas
   voltam prontas (só a senha é pedida de novo).
4. **Migre** cada `.pst` em *Configurações → Arquivos de Dados → "Migrar"* —
   vira base nativa gravável; o `.pst` original fica intacto como backup.

---

## 🛠️ Rodando em desenvolvimento

**Requisitos**: Node.js 20+ e npm

```bash
npm install
npm run dev        # Electron + Vite com hot-reload
```

```bash
npm run build      # compila main + preload + renderer para out/
npm start          # executa o build de produção
npm run typecheck  # checagem TypeScript sem emitir
npm test           # testes unitários
```

<details>
<summary><b>Estrutura do projeto</b></summary>

```
src/
  main/        Processo principal do Electron (Node.js)
    index.ts         Janela, IPC, boot, sync automático, lembretes
    pstService.ts    Leitura de PST + pastas-sombra .eml graváveis
    sourceManager.ts Roteamento por sourceId (PST, local, conta)
    ruleEngine.ts    Regras de entrega + sync POP3/IMAP + spam/AV
    pop3Source.ts / pop3Client.ts / imapSource.ts   Protocolos
    smtpService.ts   Envio, testes de conexão, assinaturas, iCalendar
    accountStore.ts  Contas + senhas (persistência segura)
    dataDir.ts       Pasta de dados portátil (perfil estilo Thunderbird)
    pstImport.ts     Migração PST → base nativa (com deduplicação)
  preload/     contextBridge — superfície `window.outlook` tipada
  renderer/    React + TypeScript (Vite)
    src/App.tsx               Shell: módulos, seleção, drag&drop, toasts
    src/components/           Ribbon, FolderPane, MessageList, ReadingPane…
    src/modules/              CalendarModule, ContactsModule…
    src/styles.css            Tema claro/escuro por variáveis CSS
  shared/      Tipos compartilhados entre main e renderer
scripts/       launch.js — orquestra dev/build/start
Docs/          Documentação de arquitetura (ARCHITECTURE.md)
```

</details>

---

## 🗺️ Roadmap

- [x] Empacotamento Linux (AppImage, deb, rpm, pacman, tar.gz)
- [x] Verificação de atualização + identificação Beta
- [x] Filtro de spam do servidor → Lixo eletrônico
- [x] Quarentena de antivírus ("Itens infectados", pasta protegida)
- [ ] Novo compromisso / edição de compromissos pelo inspector
- [ ] Contatos e Tarefas com CRUD completo
- [ ] Pastas de pesquisa (search folders)
- [ ] Assinatura de código

---

## 🤝 Contribuindo

Veja [CONTRIBUTING.md](CONTRIBUTING.md) e a arquitetura interna em
[Docs/ARCHITECTURE.md](Docs/ARCHITECTURE.md).

## 📄 Licença

**GPL-3.0** — livre para uso pessoal e comercial.

```
Unit Mail — cliente de email para desktop com leitura de PST
Copyright (C) 2025 Roger Oliveira

Este programa é software livre: você pode redistribuí-lo e/ou modificá-lo
sob os termos da Licença Pública Geral GNU v3. É distribuído na esperança
de ser útil, mas SEM NENHUMA GARANTIA.
```

O texto completo está em [LICENSE](LICENSE).

> [!WARNING]
> **Aviso legal** — Unit Mail é um projeto independente, **não afiliado à
> Microsoft**. "Outlook" e o formato PST são marcas/formatos da Microsoft
> Corporation; o acesso a arquivos PST é feito por implementação de
> terceiros (`pst-extractor`), com fins de interoperabilidade e migração
> de dados do próprio usuário.
