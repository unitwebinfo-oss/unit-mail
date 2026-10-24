// i18n minimalista — pt-BR é o idioma-fonte (as strings do código SÃO o
// português). `t('…')` devolve o inglês quando lang==='en', senão o original.
// Strings ainda não mapeadas aparecem em pt-BR — fallback seguro.
import { useSyncExternalStore } from 'react';
import type { Lang } from '../../shared/folderNames';

let current: Lang = 'pt';
const listeners = new Set<() => void>();

export function getLang(): Lang {
  return current;
}

export function initLang(l: Lang | string | undefined): void {
  current = l === 'en' ? 'en' : 'pt';
}

export async function setLang(l: Lang): Promise<void> {
  current = l === 'en' ? 'en' : 'pt';
  await window.outlook.setLanguage?.(current);
  for (const f of listeners) f();
}

/** Locale BCP-47 do idioma atual — para Intl/toLocaleString. */
export function locale(): string {
  return current === 'en' ? 'en-US' : 'pt-BR';
}

/** Re-renderiza o componente quando o idioma muda. */
export function useLang(): Lang {
  return useSyncExternalStore(
    (cb) => {
      listeners.add(cb);
      return () => listeners.delete(cb);
    },
    () => current,
  );
}

/**
 * Traduz uma string-fonte em pt-BR. `{nome}` no texto é interpolado com
 * `vars` — ex.: t('Movendo {n} itens…', { n: 3 }).
 */
export function t(pt: string, vars?: Record<string, string | number>): string {
  let s = current === 'en' ? EN[pt] ?? pt : pt;
  if (vars) for (const [k, v] of Object.entries(vars)) s = s.replaceAll(`{${k}}`, String(v));
  return s;
}

// ---------------------------------------------------------------------------
// Dicionário pt → en. Chave = string exata usada no código.
// ---------------------------------------------------------------------------
const EN: Record<string, string> = {
  // ---- Abas da faixa ----
  'Arquivo': 'File',
  'Página Inicial': 'Home',
  'Enviar/Receber': 'Send / Receive',
  'Pasta': 'Folder',
  'Exibição': 'View',

  // ---- Ribbon: botões ----
  'Novo Email': 'New Email',
  'Novos Itens': 'New Items',
  'Responder': 'Reply',
  'Responder a Todos': 'Reply All',
  'Encaminhar': 'Forward',
  'Excluir': 'Delete',
  'Lixo Eletrônico': 'Junk',
  'Arquivar': 'Archive',
  'Marcar como lida': 'Mark as Read',
  'Marcar como não lida': 'Mark as Unread',
  'Marcar Todas como Lidas': 'Mark All as Read',
  'Sinalizador': 'Follow Up',
  'Categorizar': 'Categorize',
  'Regras': 'Rules',
  'Mover': 'Move',
  'Enviar/Receber Tudo': 'Send/Receive All',
  'Mostrar Progresso': 'Show Progress',
  'Cancelar Tudo': 'Cancel All',
  'Cancelar Tarefa': 'Cancel Task',
  'Baixar Cabeçalhos': 'Download Headers',
  'Atualizar Pasta': 'Update Folder',
  'Configurações': 'Settings',
  'Regras e Alertas': 'Rules & Alerts',
  'Nova Pasta': 'New Folder',
  'Renomear Pasta': 'Rename Folder',
  'Excluir Pasta': 'Delete Folder',
  'Limpar Pasta': 'Clean Up Folder',
  'Excluir Tudo': 'Delete All',
  'Copiar Pasta': 'Copy Folder',
  'Mover Pasta': 'Move Folder',
  'Arquivos de Dados': 'Data Files',
  'Imprimir': 'Print',
  'Enviar': 'Send',
  'Anexar': 'Attach',
  'Salvar': 'Save',
  'Fechar': 'Close',
  'Cancelar': 'Cancel',
  'OK': 'OK',
  'Sim': 'Yes',
  'Não': 'No',
  'Aplicar': 'Apply',
  'Remover': 'Remove',
  'Adicionar': 'Add',
  'Editar': 'Edit',
  'Novo': 'New',
  'Nova': 'New',
  'Voltar': 'Back',
  'Pesquisar': 'Search',
  'Ajuda': 'Help',
  'Assinaturas': 'Signatures',
  'Geral': 'General',
  'Contas': 'Accounts',
  'Pessoas': 'People',
  'Tarefas': 'Tasks',
  'Calendário': 'Calendar',
  'Email para': 'Mail to',

  // ---- Menus de contexto ----
  'Responder a todos': 'Reply all',
  'Marcar como Lida': 'Mark as Read',
  'Marcar como Não Lida': 'Mark as Unread',
  'Mover para…': 'Move to…',
  'Mover para': 'Move to',
  'Executar regras': 'Run rules',
  'Executar Regras': 'Run Rules',
  'Bloquear Remetente': 'Block Sender',
  'Categorias': 'Categories',
  'Excluir pasta': 'Delete folder',
  'Renomear': 'Rename',
  'Criar subpasta': 'New subfolder',
  'Ocultar pasta': 'Hide folder',
  'Mostrar pastas ocultas': 'Show hidden folders',
  'Esvaziar pasta': 'Empty folder',
  'Expandir tudo': 'Expand all',
  'Recolher tudo': 'Collapse all',
  'Nova Pasta…': 'New Folder…',
  'Mover Pasta…': 'Move Folder…',
  'Reexibir Pasta': 'Unhide Folder',
  'Ocultar Pasta': 'Hide Folder',
  'Remover dos Favoritos': 'Remove from Favorites',
  'Adicionar a Favoritos': 'Add to Favorites',
  'Classificar Subpastas A a Z': 'Sort Subfolders A to Z',
  'Mover para cima': 'Move Up',
  'Mover para Baixo': 'Move Down',
  'Nome da nova pasta': 'New folder name',
  'Novo nome': 'New name',
  'Ver progresso de envio/recebimento': 'View send/receive progress',
  'Enviar/receber': 'Send/Receive',
  'Mover "{nome}" para:': 'Move "{nome}" to:',
  'Favoritos': 'Favorites',
  'Nenhuma base aberta': 'No data file open',
  'Nova pasta nesta base': 'New folder in this file',

  // ---- Lista de mensagens ----
  'De': 'From',
  'Assunto': 'Subject',
  'Recebido': 'Received',
  'Tamanho': 'Size',
  'Para': 'To',
  'Recebido em': 'Received',
  'Enviado': 'Sent',
  '(sem destinatário)': '(no recipient)',
  'Cc': 'Cc',
  'Cco': 'Bcc',
  'Data': 'Date',
  'Não lidos': 'Unread',
  'Todos': 'All',
  'Nenhuma mensagem': 'No messages',
  'Nenhum item nesta pasta': 'No items in this folder',
  'Nenhuma mensagem não lida': 'No unread messages',
  'Nenhum resultado': 'No results',
  'Selecione uma mensagem para ler': 'Select a message to read',
  'Pesquisar nesta pasta': 'Search this folder',
  'Pesquisar emails (assunto, remetente, destinatário, prévia)…':
    'Search email (subject, sender, recipient, preview)…',
  'pesquisando…': 'searching…',
  'resultado(s)': 'result(s)',
  'selecionado(s)': 'selected',
  'Todas': 'All',
  'Todas as categorias': 'All categories',
  'Sem categoria': 'No category',
  'Filtrar por categoria': 'Filter by category',
  'Limpar filtro de categoria': 'Clear category filter',
  'Mover…': 'Move…',
  'Regras…': 'Rules…',
  'Ignorar': 'Ignore',
  'Localizar Relacionado': 'Find Related',
  'Mover {alvo} para:': 'Move {alvo} to:',
  'Categorizar {alvo}:': 'Categorize {alvo}:',
  '{n} itens': '{n} items',
  'Limpar Todas as Categorias': 'Clear All Categories',

  // ---- Módulos / status ----
  'Email': 'Mail',
  'Sincronizando': 'Syncing',
  'Conectado': 'Connected',
  'Pronto': 'Ready',
  'Erro': 'Error',
  'Aviso': 'Warning',
  'Itens': 'Items',
  'itens': 'items',
  'item': 'item',
  'não lida': 'unread',
  'não lidas': 'unread',

  // ---- Janela de composição ----
  'Nova Mensagem': 'New Message',
  'Descartar': 'Discard',
  'Salvar Rascunho': 'Save Draft',
  'Anexar Arquivo': 'Attach File',
  'Opções': 'Options',
  'Alta importância': 'High importance',
  'Baixa importância': 'Low importance',
  'Solicitar confirmação de leitura': 'Request read receipt',
  'Sem assunto': '(no subject)',

  // ---- Configurações → Geral ----
  'Aparência': 'Appearance',
  'Tema': 'Theme',
  'Claro': 'Light',
  'Escuro': 'Dark',
  'Sistema': 'System',
  'Idioma': 'Language',
  'Português (Brasil)': 'Portuguese (Brazil)',
  'Inglês': 'English',
  'Inicialização': 'Startup',
  'Sincronização': 'Synchronization',
  'Segurança': 'Security',
  'Remetentes bloqueados': 'Blocked senders',
  'Servidor IMAP': 'IMAP server',
  'IMAP mantém as mensagens no servidor. Nenhum arquivo local será criado ou vinculado automaticamente.': 'IMAP keeps messages on the server. No local data file will be created or linked automatically.',
  'Localizar na lista…': 'Search the list…',
  'Aceitar certificado autoassinado (servidor próprio/hospedagem)': 'Accept self-signed certificate (own server/hosting)',
  'Nenhum bloqueado encontrado para esta busca.': 'No blocked sender matches this search.',
  'Lixo eletrônico': 'Junk email',
  'Verificar vírus': 'Virus scan',
  'Enviar e receber a cada': 'Send/receive every',
  'minutos': 'minutes',
  'Carregar imagens externas': 'Load external images',
  'Pasta de dados': 'Data folder',

  // ---- Backstage ----
  'Informações': 'Info',
  'Abrir e Exportar': 'Open & Export',
  'Salvar Anexos': 'Save Attachments',
  'Opções da Conta': 'Account Settings',
  'Sobre': 'About',
  'Sair': 'Exit',

  // ---- Genérico ----
  'Carregando…': 'Loading…',
  'Carregando': 'Loading',
  'Processando…': 'Processing…',
  'Concluído': 'Done',
  'Detalhes': 'Details',
  'Nome': 'Name',
  'Descrição': 'Description',
  'Tipo': 'Type',
  'Caminho': 'Path',
  'Ações': 'Actions',
  'Confirmar': 'Confirm',
  'Confirmar exclusão': 'Confirm delete',
  'Tem certeza?': 'Are you sure?',
  'Esta ação não pode ser desfeita.': 'This action cannot be undone.',
  'Selecionar tudo': 'Select all',
  'Importante': 'Important',
  'Hoje': 'Today',
  'Ontem': 'Yesterday',
  'Semana': 'Week',
  'Mês': 'Month',
  'Dia': 'Day',
  'Agenda': 'Schedule',
  'Local': 'Location',
  'Início': 'Start',
  'Término': 'End',
  'Lembrete': 'Reminder',
  'Novo compromisso': 'New appointment',
  'Novo contato': 'New contact',
  'Nova tarefa': 'New task',
  'Telefone': 'Phone',
  'Empresa': 'Company',
  'Cargo': 'Job title',
  'Aniversário': 'Birthday',
  'Observações': 'Notes',
  'Movendo para "{pasta}"…': 'Moving to "{pasta}"…',
  'Movendo {n} itens para "{pasta}"…': 'Moving {n} items to "{pasta}"…',
  'Erro ao mover: {msg}': 'Error moving: {msg}',

  // ---- Ribbon: grupos e botões ----
  'Marcas': 'Tags',
  'Compromisso': 'Appointment',
  'Contato': 'Contact',
  'Tarefa': 'Task',
  'Ferramentas': 'Tools',
  'Enviar e Receber': 'Send and Receive',
  'Baixar': 'Download',
  'Limpar': 'Clean Up',
  'Pasta atual': 'Current Folder',
  'Painel de Leitura': 'Reading Pane',
  'Atualizar': 'Refresh',
  'Enviar/Receber Todas as Pastas': 'Send/Receive All Folders',
  'Enviar Tudo': 'Send All',
  'Executar Regras Agora': 'Run Rules Now',
  'À Direita': 'Right',
  'Inferior': 'Bottom',
  'Desativado': 'Off',
  'Novo Compromisso': 'New Appointment',
  'Novo Contato': 'New Contact',
  'Nova Tarefa': 'New Task',
  'Todas as Categorias…': 'All Categories…',
  'Sincronizar a cada': 'Sync every',
  'desligado': 'off',
  '{n} min': '{n} min',
  'Nenhuma conta configurada': 'No accounts configured',
  'Novo nome da pasta': 'New folder name',
  'Grupos de Envio/Recebimento': 'Send/Receive Groups',

  // ---- Backstage ----
  'Informações sobre Contas': 'Account Information',
  '+ Adicionar Conta': '+ Add Account',
  'Configurações de Conta': 'Account Settings',
  'Altere configurações para esta conta ou configure mais conexões.':
    'Change settings for this account or set up more connections.',
  'Configurações de Caixa de Correio': 'Mailbox Settings',
  'Gerencie arquivos de dados, bases locais e onde as mensagens são entregues.':
    'Manage data files, local stores and where messages are delivered.',
  'Gerenciar Regras e Alertas': 'Manage Rules & Alerts',
  'Use Regras e Alertas para organizar as mensagens de email de entrada.':
    'Use Rules & Alerts to organize incoming email messages.',
  '{n} base(s) de dados aberta(s)': '{n} data file(s) open',
  'Abrir': 'Open',
  'Exportar': 'Export',
  'Importar .pst (Outlook)': 'Import .pst (Outlook)',
  'Traz todas as pastas e mensagens de um arquivo .pst para uma base do Unit Mail — o arquivo original nunca é alterado.':
    'Brings all folders and messages from a .pst file into a Unit Mail store — the original file is never changed.',
  'Importar do Thunderbird': 'Import from Thunderbird',
  'Traz todas as pastas e mensagens do perfil do Thunderbird instalado neste computador — o perfil original nunca é alterado.':
    'Brings all folders and messages from the Thunderbird profile installed on this computer — the original profile is never changed.',
  'Novo Arquivo de Dados': 'New Data File',
  'Crie uma base local gravável para receber e organizar mensagens.':
    'Create a writable local store to receive and organize messages.',
  'Exportar configurações (backup)': 'Export settings (backup)',
  'Salva contas, regras e assinaturas num arquivo .json para levar para outra máquina. O arquivo contém as senhas — guarde com segurança.':
    'Saves accounts, rules and signatures to a .json file to take to another machine. The file contains passwords — keep it safe.',
  'Importar configurações': 'Import settings',
  'Restaura contas, regras e assinaturas de um backup. Processo recomendado: primeiro adicione os arquivos de dados, depois importe o backup e re-vincule as contas em Configurações → Email.':
    'Restores accounts, rules and signatures from a backup. Recommended process: first add the data files, then import the backup and re-link accounts in Settings → Mail.',
  'Opções e preferências do aplicativo': 'App options and preferences',
  'Salvar como': 'Save as',
  'Salvar a mensagem aberta como .eml': 'Save the open message as .eml',
  'Abra uma mensagem para salvar': 'Open a message to save',
  'Imprimir a mensagem aberta': 'Print the open message',
  'Abra uma mensagem para imprimir': 'Open a message to print',
  'Atribuir categorias coloridas aos itens selecionados': 'Assign color categories to selected items',
  'Abra uma mensagem para categorizar': 'Open a message to categorize',

  // ---- Janela de composição ----
  'Restaurar tamanho': 'Restore size',
  'Maximizar janela': 'Maximize window',
  'Endereço(s) com formato inválido:': 'Invalid address format:',
  'corrija antes de enviar.': 'fix before sending.',
  'Anexar arquivo': 'Attach file',
  'Enviando…': 'Sending…',
  'Fonte': 'Font',
  'Pequena': 'Small',
  'Normal': 'Normal',
  'Média': 'Medium',
  'Grande': 'Large',
  'Muito grande': 'Very large',
  'Negrito': 'Bold',
  'Itálico': 'Italic',
  'Sublinhado': 'Underline',
  'Tachado': 'Strikethrough',
  'Cor do texto': 'Text color',
  'Marcador de cor do texto (realce)': 'Text highlight color',
  'Alinhar à esquerda': 'Align left',
  'Centralizar': 'Center',
  'Alinhar à direita': 'Align right',
  'Lista com marcadores': 'Bulleted list',
  'Lista numerada': 'Numbered list',
  'Diminuir recuo': 'Decrease indent',
  'Aumentar recuo': 'Increase indent',
  'Inserir link': 'Insert link',
  'Inserir imagem': 'Insert picture',
  'Limpar formatação': 'Clear formatting',
  'Link:': 'Link:',

  // ---- Configurações — cabeçalho e abas ----
  'Contas, arquivos de dados, regras e preferências':
    'Accounts, data files, rules and preferences',
  'Contas de email': 'Email accounts',
  'Recebimento por POP3 ou IMAP; envio por SMTP':
    'Receive via POP3 or IMAP; send via SMTP',
  'Editar conta': 'Edit account',
  'Nova conta de email': 'New email account',
  'Adicionar conta de email': 'Add email account',
  'Testar conexão': 'Test connection',
  'Importar contas de:': 'Import accounts from:',
  'Conta': 'Account',
  'Recebimento': 'Receiving',
  'Entrega': 'Delivery',
  'Situação': 'Status',
  'Conectando…': 'Connecting…',
  'Erro de conexão': 'Connection error',
  'Pronta': 'Ready',
  'Re-sincronizar': 'Re-sync',
  'Servidores': 'Servers',
  'Informações do usuário': 'User information',
  'Nome de exibição': 'Display name',
  'Endereço de email': 'Email address',
  'Protocolo': 'Protocol',
  'Logon': 'Logon',
  'Senha': 'Password',
  'Senha (armazenada criptografada localmente)': 'Password (stored encrypted locally)',
  'Arquivo de dados': 'Data file',
  'Nome do novo arquivo de dados': 'New data file name',
  'Porta': 'Port',
  'Criptografia': 'Encryption',
  'Nenhum': 'None',
  'Regra': 'Rule',
  'Base padrão': 'Default store',
  '(principal)': '(default)',
  'Nenhuma conta cadastrada. Clique em "Adicionar conta" para começar.':
    'No accounts yet. Click "Add account" to get started.',

  // ---- Configurações → Geral ----
  'Tema da interface': 'Interface theme',
  'Igual ao sistema': 'Same as system',
  'Idioma da interface': 'Interface language',
  'Pastas padrão já criadas (Caixa de Entrada, Lixo Eletrônico…) passam a ser exibidas no novo idioma; os nomes gravados no disco nunca são alterados — pastas novas nascem com o nome do idioma escolhido.':
    'Default folders already created (Inbox, Junk…) keep their saved names — they are never renamed on disk. New folders are created with names in the chosen language.',
  'Envio e recebimento': 'Send and receive',
  'Verificar novas mensagens a cada': 'Check for new messages every',
  '1 minuto': '1 minute',
  '{n} minutos': '{n} minutes',
  '1 hora': '1 hour',
  '1 Lembrete': '1 Reminder',
  '{n} Lembretes': '{n} Reminders',
  'Abrir item': 'Open Item',
  'Dispensar': 'Dismiss',
  'Dispensar todos': 'Dismiss All',
  'Adiar': 'Snooze',
  'Lembrete adiado por {n} min': 'Reminder snoozed for {n} min',
  'Privacidade': 'Privacy',
  'Baixar imagens externas automaticamente': 'Download external images automatically',
  'Quando desativado (padrão), imagens da internet dentro de emails HTML ficam bloqueadas — remetentes não podem rastrear quando você abriu a mensagem.':
    'When off (default), internet images inside HTML emails stay blocked — senders cannot track when you opened the message.',
  'Verificar vírus nas mensagens recebidas': 'Scan incoming messages for viruses',
  'Cada mensagem baixada é verificada pelo antivírus do computador{eng}. Mensagens infectadas vão direto para a pasta de quarentena, sem passar pela Caixa de Entrada.':
    'Every downloaded message is scanned by the computer antivirus{eng}. Infected messages go straight to the quarantine folder, skipping the Inbox.',
  'mensagens destes endereços ou domínios caem direto na pasta de lixo eletrônico, sem passar pelas regras. Também dá para bloquear pelo botão direito na mensagem ("Bloquear Remetente").':
    'messages from these addresses or domains go straight to the junk folder, skipping rules. You can also block via the message right-click menu ("Block Sender").',
  'Bloquear': 'Block',
  'Desbloquear': 'Unblock',
  'Remover da lista': 'Remove from list',
  'Configurações salvas.': 'Settings saved.',
  'Não foi possível identificar o endereço do remetente.':
    'Could not identify the sender address.',
  '"{a}" bloqueado — próximas mensagens vão direto para a pasta de lixo eletrônico.':
    '"{a}" blocked — future messages go straight to the junk folder.',
  '{n} remetentes bloqueados — próximas mensagens vão para a pasta de lixo eletrônico.':
    '{n} senders blocked — future messages go to the junk folder.',

  // ---- Configurações → Arquivos de dados / Regras / Assinaturas ----
  'Arquivos de dados': 'Data files',
  'Adicionar arquivo de dados': 'Add data file',
  'Nova base…': 'New store…',
  'Importar .pst…': 'Import .pst…',
  'Thunderbird…': 'Thunderbird…',
  'Procurando…': 'Searching…',
  'Definir principal': 'Set as default',
  'Migrar…': 'Migrate…',
  'Alterar pasta…': 'Change folder…',
  'Regras de recebimento': 'Receive rules',
  'Executar agora': 'Run now',
  'Verificando…': 'Checking…',
  'Preparando…': 'Preparing…',
  'Duplicar': 'Duplicate',
  'Nova assinatura': 'New signature',
  'Criar assinatura': 'Create signature',
  'Excluir assinatura': 'Delete signature',
  'Nenhuma assinatura — clique em "Nova assinatura" para criar.':
    'No signatures — click "New signature" to create one.',
  'Não usada por nenhuma conta': 'Not used by any account',
  'Procurar contas do Outlook instalado no Windows':
    'Look for accounts in Outlook installed on Windows',
  'Procurar contas nos perfis do Thunderbird': 'Look for accounts in Thunderbird profiles',
  'Importar arquivo de exportação de contas (.json)': 'Import account export file (.json)',
  'Arquivo .json…': '.json file…',
  'Recebimento por POP3 ou IMAP; envio por SMTP — igual ao Outlook.':
    'Receive via POP3 or IMAP; send via SMTP — like Outlook.',
  'O sistema não oferece armazenamento seguro de senhas (keyring indisponível). As senhas das contas ficam sem criptografia no arquivo de dados. Instale e ative um cofre de senhas (GNOME Keyring ou KWallet) e recadastre as contas.':
    'The system has no secure password storage (keyring unavailable). Account passwords are stored unencrypted in the data file. Install and enable a password vault (GNOME Keyring or KWallet) and re-add the accounts.',
  'Esquece o que já foi baixado e re-sincroniza tudo do servidor':
    'Forgets what was downloaded and re-syncs everything from the server',
  'Onde as mensagens desta conta são guardadas. Você pode usar um arquivo que já existe ou criar um novo só para esta conta.':
    'Where this account messages are stored. You can use an existing file or create a new one just for this account.',
  'Usar um arquivo de dados existente': 'Use an existing data file',
  'Sem escolha, a conta usa o arquivo principal. Entrega na Caixa de Entrada.':
    'Without a choice, the account uses the default file. Delivers to Inbox.',
  'Criar um novo arquivo de dados': 'Create a new data file',
  'Cria uma base própria ao salvar. Se for o primeiro, vira o principal.':
    'Creates its own store on save. If it is the first, it becomes the default.',
  'Servidor de entrada': 'Incoming server',
  'Servidor de saída': 'Outgoing server',
  'Entrega no servidor': 'Server delivery',
  'Deixar uma cópia das mensagens no servidor': 'Leave a copy of messages on the server',
  'Onde as mensagens são guardadas. Marque uma como': 'Where messages are stored. Mark one as',
  'Principal': 'Default',
  'mensagens sem regra e contas sem base vinculada são entregues na Caixa de Entrada dela.':
    'messages without a rule and accounts without a bound store are delivered to its Inbox.',
  'Cria ou importa bases no formato próprio do app (arquivos originais nunca são alterados):':
    'Creates or imports stores in the app own format (original files are never changed):',
  'Todos os seus dados ficam nesta pasta — contas, regras, assinaturas e as bases de email dentro de':
    'All your data lives in this folder — accounts, rules, signatures and email stores inside',
  '(como o perfil do Thunderbird). Para mover para outro computador, copie a pasta inteira e aponte o app para ela aqui: tudo volta a funcionar sem recadastrar nada.':
    '(like the Thunderbird profile). To move to another computer, copy the whole folder and point the app to it here: everything works again with no re-setup.',
  'Quando chegar email novo, as regras são avaliadas em ordem — a primeira que combinar arquiva a mensagem na pasta escolhida.':
    'When new email arrives, rules are evaluated in order — the first match files the message into the chosen folder.',
  'Regra sem condições combina com TODAS as mensagens':
    'A rule with no conditions matches ALL messages',
  'preencha pelo menos um campo "contém" para filtrar.':
    'fill in at least one "contains" field to filter.',
  'Clique em': 'Click',
  'para abrir o editor e marque': 'to open the editor and check',
  'quais contas usam': 'which accounts use',
  'cada assinatura — a mesma assinatura pode ser aplicada a vários emails, em novas mensagens e/ou respostas.':
    'each signature — the same signature can be applied to several emails, for new messages and/or replies.',
  'Nenhum remetente bloqueado — use o campo acima ou o botão direito numa mensagem.':
    'No blocked senders — use the field above or right-click a message.',
  '{n} bloqueado(s)': '{n} blocked',

  // ---- Editor de regra ----
  'Editar regra': 'Edit rule',
  'Nova regra': 'New rule',
  'Nome da regra': 'Rule name',
  'Aplicar quando chegar email em': 'Apply when email arrives in',
  '— selecione a conta (obrigatório) —': '— select the account (required) —',
  '(conta removida)': '(account removed)',
  'A conta desta regra foi removida — selecione uma conta existente para corrigir.':
    'The account for this rule was removed — select an existing account to fix.',
  'É obrigatório escolher uma conta existente.': 'You must choose an existing account.',
  'Condições — preencha pelo menos uma para filtrar':
    'Conditions — fill at least one to filter',
  'De contém (opcional)': 'From contains (optional)',
  'Para contém (opcional)': 'To contains (optional)',
  'Assunto contém (opcional)': 'Subject contains (optional)',
  'Sem condições, esta regra vale para': 'With no conditions, this rule applies to',
  'todas as mensagens': 'all messages',
  'que chegarem na conta — elas serão arquivadas na pasta abaixo em vez da entrega padrão.':
    'arriving on the account — they will be filed in the folder below instead of the default delivery.',
  'Destino — onde a mensagem é arquivada': 'Destination — where the message is filed',
  'Arquivar na base de dados': 'File into data store',
  '— selecione a base —': '— select the store —',
  'Pasta de destino': 'Destination folder',
  '— selecione a pasta —': '— select the folder —',
  'Nova pasta…': 'New folder…',
  'Nome da pasta (será criada se não existir)': 'Folder name (created if missing)',
  'Selecione a base primeiro': 'Select the store first',
  '← escolher pasta existente': '← choose existing folder',
  'Parar de processar mais regras': 'Stop processing more rules',
  'Salvar regra': 'Save rule',
  'Aplica regras às mensagens de uma pasta existente, sem esperar um novo recebimento.':
    'Applies rules to messages in an existing folder, without waiting for new mail.',
  'Escolha a regra e a pasta — o app verifica cada mensagem da pasta e move para o destino da regra as que combinam. Regras desativadas também podem ser executadas.':
    'Choose the rule and the folder — the app checks each message in the folder and moves matching ones to the rule destination. Disabled rules can also run.',
  '— escolha a regra —': '— choose the rule —',
  '(sem nome)': '(no name)',
  ' (desativada)': ' (disabled)',
  'Atenção: esta regra não tem condições — combina com TODAS as mensagens da pasta escolhida.':
    'Warning: this rule has no conditions — it matches ALL messages in the chosen folder.',
  'Nenhuma regra cadastrada — crie uma na lista abaixo.':
    'No rules yet — create one in the list below.',
  'Abra um arquivo de dados para executar as regras.': 'Open a data file to run rules.',
  'Verificando {done} / {total} mensagens': 'Checking {done} / {total} messages',
  ' — {n} combinaram até agora': ' — {n} matched so far',
  'Nenhuma regra cadastrada — mensagens novas vão para a pasta de entrega da conta.':
    'No rules yet — new messages go to the account delivery folder.',
  'Regra ativa — clique para desativar': 'Rule on — click to disable',
  'Regra desativada — clique para ativar': 'Rule off — click to enable',
  'automática': 'automatic',
  'A conta de email desta regra foi removida. Edite e escolha uma conta existente.':
    'The email account for this rule was removed. Edit and pick an existing account.',
  'conta ausente': 'account missing',
  'em': 'in',
  'qualquer mensagem': 'any message',
  'De contém': 'From contains',
  'Para contém': 'To contains',
  'Assunto contém': 'Subject contains',
  'arquivar em': 'file into',
  'Duplicar regra — abre o editor para escolher a conta de email':
    'Duplicate rule — opens the editor to pick the email account',

  // ---- Assinaturas ----
  'Assinaturas encontradas na importação do Outlook':
    'Signatures found in the Outlook import',
  'Nenhuma conta usa ainda': 'No account uses it yet',
  'novas: {list}': 'new: {list}',
  'respostas: {list}': 'replies: {list}',
  'novas': 'new',
  'respostas': 'replies',
  'Excluir a assinatura "{name}"? Contas que a usam ficarão sem assinatura.':
    'Delete signature "{name}"? Accounts using it will have no signature.',
  'Assinatura {n}': 'Signature {n}',
  'Editar assinatura': 'Edit signature',
  'Nome da assinatura': 'Signature name',
  'Conteúdo': 'Content',
  'Seu Nome — Cargo, Empresa, telefone…': 'Your Name — Title, Company, phone…',
  'Para redimensionar uma imagem, clique nela e arraste a alça do canto inferior direito — a proporção é mantida.':
    'To resize an image, click it and drag the bottom-right handle — aspect ratio is kept.',
  'Contas que usam esta assinatura': 'Accounts using this signature',
  'Nenhuma conta cadastrada — crie uma na aba Email.':
    'No accounts yet — create one in the Mail tab.',
  'Novas mensagens': 'New messages',
  'Respostas e encaminhamentos': 'Replies and forwards',
  'Novas': 'New',
  'Respostas': 'Replies',
  'Os vínculos são aplicados ao clicar em Salvar.': 'Assignments apply when you click Save.',

  // ---- Arquivos de dados — cartões e ações ----
  'Nenhuma base aberta. Crie uma base local ou importe um .pst abaixo.':
    'No store open. Create a local store or import a .pst below.',
  'PST · somente leitura': 'PST · read-only',
  'Base local': 'Local store',
  'vinculada a {email}': 'bound to {email}',
  'sem conta vinculada': 'no account bound',
  'Tornar esta a base principal (destino padrão das mensagens)':
    'Make this the default store (default message destination)',
  'Renomear a base de dados': 'Rename the data store',
  'Copia todo o conteúdo do PST para uma base própria do app — lá, excluir apaga de verdade e libera espaço':
    'Copies all PST content into a native app store — there, deleting really deletes and frees space',
  'Apaga a pasta e todas as mensagens do disco — não pode ser desfeito':
    'Deletes the folder and all messages from disk — cannot be undone',
  'Fecha o PST no app — o arquivo original nunca é apagado':
    'Closes the PST in the app — the original file is never deleted',
  'Alterar pasta de dados': 'Change data folder',
  '"{dir}" já contém um perfil do Unit Mail — contas, regras e bases serão restaurados de lá. Confirmar?':
    '"{dir}" already contains a Unit Mail profile — accounts, rules and stores will be restored from it. Confirm?',
  'Mover seus dados para "{dir}"? As bases locais serão copiadas para a subpasta Mail/ e as referências atualizadas.':
    'Move your data to "{dir}"? Local stores will be copied into the Mail/ subfolder and references updated.',
  'Pasta de dados alterada': 'Data folder changed',
  'Dados copiados. Reinicie o app para concluir.': 'Data copied. Restart the app to finish.',
  'Reiniciar': 'Restart',
  'Abrir a pasta de dados no explorador': 'Open data folder in file explorer',

  // ---- Modal: novo arquivo de dados ----
  'Base própria do Unit Mail — mensagens em .eml': 'Native Unit Mail store — .eml messages',
  'Nome do arquivo de dados': 'Data file name',
  'Ex.: Arquivo 2024, Emails — Trabalho…': 'E.g.: Archive 2024, Mail — Work…',
  'Vincular a uma conta (opcional)': 'Bind to an account (optional)',
  'Sem vínculo — qualquer conta pode usar': 'No binding — any account can use it',
  ' (já tem base)': ' (already has a store)',
  'Sem vínculo, a base fica disponível para receber mensagens de várias contas — todas caem na mesma Caixa de Entrada (use Regras para separar em pastas). Com vínculo, só as mensagens daquela conta são entregues aqui.':
    'Unbound, the store can receive messages from several accounts — all land in the same Inbox (use Rules to split into folders). Bound, only that account messages are delivered here.',
  'Criar na pasta de dados': 'Create in data folder',
  'Criar em outra pasta…': 'Create in another folder…',

  // ---- Modal: migrar/importar PST ----
  'Importar arquivo .pst': 'Import .pst file',
  'Migrar para base própria': 'Migrate to native store',
  'Copia as pastas e mensagens do PST para uma base 100% do Unit Mail — sem a limitação de "somente leitura" do formato .pst. Nessa base, excluir apaga o arquivo de verdade e libera espaço em disco de imediato. Escolha abaixo quais pastas importar (padrão: todas).':
    'Copies the PST folders and messages into a 100% Unit Mail store — without the .pst read-only limitation. There, deleting really deletes files and frees disk space immediately. Pick which folders to import below (default: all).',
  'Pastas a importar': 'Folders to import',
  'Carregando pastas…': 'Loading folders…',
  'Nenhuma pasta encontrada neste arquivo.': 'No folders found in this file.',
  'Nenhuma': 'None',
  '{sel} de {tot} pasta(s)': '{sel} of {tot} folder(s)',
  'Marcar uma pasta inclui as subpastas dela. Mensagens já existentes no destino não são duplicadas.':
    'Checking a folder includes its subfolders. Messages already in the destination are not duplicated.',
  'Pasta de dados não encontrada': 'Data folder not found',
  'A pasta configurada não está acessível agora — se ela está num drive externo/removível, conecte-o ou aguarde a montagem e tente de novo. Nada foi alterado ou sincronizado.':
    'The configured folder is not accessible right now — if it lives on an external/removable drive, plug it in or wait for it to mount and try again. Nothing was changed or synced.',
  'Escolher outra pasta…': 'Choose another folder…',
  'Tentar novamente': 'Try again',
  'O PST original': 'The original PST',
  'não é alterado nem apagado': 'is not changed nor deleted',
  ' só foi aberto para ler os dados e será fechado ao final; o arquivo continua no disco igual estava.':
    ' was only opened to read the data and will be closed at the end; the file stays on disk as it was.',
  ' continua intacto e aberto no app.': ' stays intact and open in the app.',
  'Itens de calendário/contatos/tarefas são convertidos para um registro com assunto e corpo (o formato ICS/vCard original não é preservado).':
    'Calendar/contact/task items are converted into a record with subject and body (the original ICS/vCard format is not preserved).',
  'PSTs grandes podem demorar': 'Large PSTs can take a while',
  'um arquivo de vários GB leva bastante tempo para copiar item a item. O progresso é mostrado durante a importação; não feche o app enquanto ela roda.':
    'a multi-GB file takes a long time to copy item by item. Progress is shown during the import; do not close the app while it runs.',
  'Importar para': 'Import into',
  'Nova base de dados do Unit Mail': 'New Unit Mail data store',
  'base existente': 'existing store',
  'As pastas e mensagens do PST serão mescladas dentro dessa base — itens com nomes de pasta iguais se juntam.':
    'The PST folders and messages will be merged into that store — items with equal folder names join together.',
  'Nome da nova base': 'New store name',
  'Pasta onde criar': 'Folder to create in',
  'Pasta de dados padrão (Mail/) — escolher outra ao lado':
    'Default data folder (Mail/) — pick another beside',
  'Escolher…': 'Browse…',
  'Importar': 'Import',
  'Iniciar migração': 'Start migration',
  'Migrando "{f}"… não feche o app.': 'Migrating "{f}"… do not close the app.',
  ' — {n} já existiam (puladas)': ' — {n} already existed (skipped)',
  'Importação': 'Import',
  'Migração': 'Migration',
  'concluída — base': 'finished — store',
  'criada com todo o conteúdo.': 'created with all content.',
  'O arquivo .pst original não foi alterado — só não fica mais aberto na lista.':
    'The original .pst file was not changed — it just no longer stays open in the list.',
  'O PST original continua aberto e intocado; feche-o quando quiser (aba "Arquivos de Dados" → Fechar).':
    'The original PST stays open and untouched; close it whenever you want (Data Files tab → Close).',
  'Defina a nova base como principal, se desejar.': 'Set the new store as default if you like.',
  'Concluir': 'Finish',

  // ---- Modal: Thunderbird ----
  'mensagens (.mbox) do perfil instalado neste computador':
    'messages (.mbox) from the profile installed on this computer',
  'Nenhum perfil do Thunderbird com mensagens foi encontrado neste computador.':
    'No Thunderbird profile with messages was found on this computer.',
  'Copia TODAS as pastas e mensagens do perfil para uma base do Unit Mail — o perfil do Thunderbird não é alterado. Perfis grandes podem demorar; não feche o app durante a importação.':
    'Copies ALL folders and messages from the profile into a Unit Mail store — the Thunderbird profile is not changed. Large profiles may take a while; do not close the app during import.',
  'Encontrado: {list}': 'Found: {list}',
  'As pastas e mensagens do perfil serão mescladas dentro dessa base — itens com nomes de pasta iguais se juntam.':
    'The profile folders and messages will be merged into that store — items with equal folder names join together.',
  'Importando "{f}"… não feche o app.': 'Importing "{f}"… do not close the app.',
  'Importação concluída — base': 'Import finished — store',
  'criada com todo o conteúdo do perfil. O Thunderbird original não foi alterado.':
    'created with all profile content. The original Thunderbird was not changed.',

  // ---- Confirmações, erros e mensagens dinâmicas ----
  'Regra sem condições': 'Rule without conditions',
  'A regra "{name}" não tem condições — combina com TODAS as mensagens da pasta escolhida e todas serão movidas para "{dest}". Deseja continuar?':
    'Rule "{name}" has no conditions — it matches ALL messages in the chosen folder and all will be moved to "{dest}". Continue?',
  'Executar assim mesmo': 'Run anyway',
  ' ({n} não movidas — o arquivo de dados de destino da regra não está aberto)':
    ' ({n} not moved — the rule destination data file is not open)',
  'Verificação interrompida — {done} mensagens verificadas, {applied} movida(s).':
    'Check interrupted — {done} messages checked, {applied} moved.',
  '{n} mensagem(ns) combinaram com a regra e foram movidas para o destino dela.':
    '{n} message(s) matched the rule and were moved to its destination.',
  'Nenhuma mensagem da pasta combinou com a regra.': 'No message in the folder matched the rule.',
  'Erro ao executar: {msg}': 'Error running: {msg}',
  'Não foi possível criar o arquivo de dados': 'Could not create the data file',
  'Configuração e senha recuperadas automaticamente ({p}). Revise e clique em Salvar.':
    'Settings and password recovered automatically ({p}). Review and click Save.',
  'Configuração importada ({p}). Digite a senha para concluir.':
    'Settings imported ({p}). Type the password to finish.',
  ' Entrega vinculada à base "{label}".': ' Delivery bound to store "{label}".',
  ' O PST de entrega do Outlook ({file}) não foi importado — importe-o na aba "Arquivos de Dados" ou escolha outro arquivo de dados na seção "Arquivo de dados" da conta.':
    ' The Outlook delivery PST ({file}) was not imported — import it in the "Data Files" tab or pick another data file in the account "Data file" section.',
  'Erro ao salvar: {msg}': 'Error saving: {msg}',
  'Remover conta': 'Remove account',
  'Remover a conta "{email}"? A base de dados vinculada "{label}" também será excluída — todas as mensagens nela serão apagadas do disco.':
    'Remove account "{email}"? The bound data store "{label}" will also be deleted — all messages in it will be erased from disk.',
  'Remover a conta "{email}"? A base de dados "{label}" está vinculada só a ela.':
    'Remove account "{email}"? The data store "{label}" is bound only to it.',
  'Apagar também "{label}" — todas as mensagens nela serão excluídas do disco.':
    'Also delete "{label}" — all messages in it will be erased from disk.',
  'Remover a conta "{email}"?{pst}': 'Remove account "{email}"?{pst}',
  'A base de dados vinculada (PST) não é apagada.': 'The bound data file (PST) is not deleted.',
  'Re-sincronizar do servidor': 'Re-sync from server',
  'Re-baixar todas as mensagens de "{email}" que ainda estiverem no servidor? ':
    'Re-download all messages from "{email}" still on the server? ',
  'O que já está na base não duplica, mas a sincronização pode demorar se a caixa for grande.':
    'What is already in the store will not duplicate, but sync may take a while if the mailbox is large.',
  'Sincronização reiniciada — use Enviar/Receber para baixar novamente.':
    'Sync restarted — use Send/Receive to download again.',
  'Excluir base de dados': 'Delete data store',
  'Excluir "{label}"? Todas as mensagens guardadas nela serão apagadas do disco — isso não pode ser desfeito.':
    'Delete "{label}"? All messages stored in it will be deleted from disk — this cannot be undone.',
  'Senhas recuperadas onde estavam salvas no Windows.':
    'Passwords recovered where they were saved in Windows.',
  'O Thunderbird guarda as senhas cifradas: digite a senha ao usar os dados.':
    'Thunderbird stores passwords encrypted: type the password when using the data.',
  'Digite a senha ao usar os dados.': 'Type the password when using the data.',
  '{n} conta(s) encontrada(s) no {name}. {note}': '{n} account(s) found in {name}. {note}',
  'Nenhuma conta encontrada no {name}.': 'No accounts found in {name}.',
  'Erro: {msg}': 'Error: {msg}',
  'Exportado para {path}. Copie este arquivo para o Linux e use "Arquivo" lá.':
    'Exported to {path}. Copy this file to Linux and use "File" there.',
  '{n} conta(s) carregada(s) do arquivo de exportação — digite a senha de cada uma ao importar.':
    '{n} account(s) loaded from the export file — type each password when importing.',
  '(cópia)': '(copy)',
  'Excluir regra': 'Delete rule',
  'Excluir a regra "{name}"?': 'Delete rule "{name}"?',
  'Qualquer conta': 'Any account',
  'Erro ao renomear: {msg}': 'Error renaming: {msg}',
  'Novo arquivo de dados': 'New data file',

  // ---- Aba Email: tooltips, campos e assistente ----
  'Cria uma base vazia — pode receber mensagens de várias contas na mesma Caixa de Entrada (use Regras para separar em pastas)':
    'Creates an empty store — it can receive messages from several accounts in the same Inbox (use Rules to split into folders)',
  'Traz TODAS as pastas e mensagens do arquivo .pst do Outlook para a base que você escolher':
    'Brings ALL folders and messages from the Outlook .pst file into the store you pick',
  'Traz todas as pastas e mensagens do perfil do Thunderbird instalado neste computador':
    'Brings all folders and messages from the Thunderbird profile installed on this computer',
  '•••••••• (senha salva)': '•••••••• (password saved)',
  '(deixe em branco para manter a atual)': '(leave blank to keep current)',
  '✓ Senha salva — preencha apenas se quiser trocar.':
    '✓ Password saved — fill only if you want to change it.',
  'Nenhum arquivo de dados aberto': 'No data file open',
  'Emails — {name}': 'Mail — {name}',
  'conta': 'account',
  'TLS': 'TLS',
  'Servidor SMTP (envio)': 'SMTP server (outgoing)',
  'Servidor {p} (recebimento)': '{p} server (incoming)',
  'Remover do servidor após (dias — 0 = nunca)': 'Remove from server after (days — 0 = never)',
  'Perfil do Thunderbird': 'Thunderbird profile',
  'Importar conta': 'Import account',
  'Etapa 1 de 2 — procurando contas…': 'Step 1 of 2 — searching accounts…',
  'Etapa 1 de 2 — escolha a conta ({n} encontrada(s))': 'Step 1 of 2 — choose the account ({n} found)',
  'Etapa 2 de 2 — {email}': 'Step 2 of 2 — {email}',
  'Escolha a conta que quer importar e clique em': 'Choose the account to import and click',
  'Próximo': 'Next',
  'Próximo →': 'Next →',
  'Uma por vez — as demais ficam aqui para importar depois.':
    'One at a time — the rest stay here to import later.',
  'Procurando contas…': 'Searching accounts…',
  'Lendo o Registro do Windows / perfis — pode levar alguns segundos.':
    'Reading the Windows Registry / profiles — may take a few seconds.',
  'importada': 'imported',
  'já cadastrada': 'already added',
  'senha recuperada': 'password recovered',
  'entrega em {file}': 'delivers to {file}',
  'Grava as contas encontradas num .json para levar ao Linux':
    'Saves the found accounts to a .json to take to Linux',
  'Salvar exportação para levar ao Linux': 'Save export to take to Linux',
  'usuário: {u}': 'user: {u}',
  'senha de SMTP separada recuperada': 'separate SMTP password recovered',
  'Senha recuperada do Windows': 'Password recovered from Windows',
  'Digite a senha da conta': 'Type the account password',
  'A senha foi recuperada — revise ou troque, se necessário.':
    'The password was recovered — review or change it if needed.',
  'As mensagens desta conta entram na Caixa de Entrada dele.':
    'This account messages go into its Inbox.',
  'Base própria só para esta conta.': 'A dedicated store just for this account.',
  'Deixar cópia das mensagens no servidor': 'Leave a copy of messages on the server',
  ' por {n} dias': ' for {n} days',
  'Esta conta tem assinatura do Outlook — depois da importação ela aparece como sugestão na aba Assinaturas.':
    'This account has an Outlook signature — after import it appears as a suggestion in the Signatures tab.',
  '← Voltar': '← Back',
  'Testando…': 'Testing…',
  'Testar conta': 'Test account',
  'Importando…': 'Importing…',
  'Importar esta conta': 'Import this account',
  'Abrir o formulário completo para revisar todos os campos':
    'Open the full form to review all fields',
  'Editar detalhes': 'Edit details',
  'Assinatura — {email}': 'Signature — {email}',
  'Resposta — {email}': 'Reply — {email}',

  // ---- Assistente de primeiro uso ----
  'Bem-vindo ao Unit Mail': 'Welcome to Unit Mail',
  'Configure seu email em poucos passos': 'Set up your email in a few steps',
  'Dados': 'Data',
  'Onde guardar seus dados?': 'Where to store your data?',
  'O Unit Mail guarda': 'Unit Mail stores',
  'tudo': 'everything',
  'numa única pasta — contas, regras, assinaturas e os emails, dentro de':
    'in a single folder — accounts, rules, signatures and emails, inside',
  '(como o perfil do Thunderbird). Para backup ou mudança de computador, basta copiar essa pasta.':
    '(like the Thunderbird profile). For backup or moving to another computer, just copy that folder.',
  'Voltando de outro PC?': 'Coming back from another PC?',
  'Escolha a pasta copiada': 'Pick the copied folder',
  'o app reconhece o perfil e restaura tudo sozinho.':
    'the app recognizes the profile and restores everything by itself.',
  'Pasta padrão — {label}': 'Default folder — {label}',
  'Escolher pasta…': 'Choose folder…',
  'Perfil encontrado nesta pasta.': 'Profile found in this folder.',
  'Pasta nova — seus dados serão criados aqui.': 'New folder — your data will be created here.',
  'Pode escolher outra antes de continuar.': 'You can pick another before continuing.',
  'Restaurando…': 'Restoring…',
  'Restaurar e reiniciar →': 'Restore and restart →',
  'Aplicando…': 'Applying…',
  'Continuar →': 'Continue →',
  'Esta pasta já tem dados do Unit Mail — restaure tudo, ou escolha outra pasta para começar do zero.':
    'This folder already has Unit Mail data — restore everything, or pick another folder to start fresh.',
  'Como você quer começar?': 'How do you want to start?',
  'Configurar conta': 'Set up account',
  'Adicionar uma conta POP3/IMAP e escolher onde ela entrega':
    'Add a POP3/IMAP account and choose where it delivers',
  'Importar contas': 'Import accounts',
  'Trazer as contas do Outlook ou do Thunderbird': 'Bring accounts from Outlook or Thunderbird',
  'Trazer as contas do Thunderbird deste sistema': 'Bring accounts from Thunderbird on this system',
  'O Unit Mail guarda os emails em bases de dados próprias — criadas automaticamente ao configurar a conta.':
    'Unit Mail stores emails in its own data stores — created automatically when setting up the account.',
  'Pular configuração inicial →': 'Skip initial setup →',
  'Outlook: lê as contas instaladas neste Windows, com a senha salva. Thunderbird: lê as contas de todos os perfis deste sistema (as senhas você digita).':
    'Outlook: reads the accounts installed on this Windows, with the saved password. Thunderbird: reads accounts from all profiles on this system (you type the passwords).',
  'Thunderbird: lê as contas de todos os perfis deste sistema. As senhas dele você digita. Também é possível importar um arquivo .json de exportação.':
    'Thunderbird: reads accounts from all profiles on this system. You type its passwords. You can also import a .json export file.',
  'Já tem arquivos': 'Already have',
  'Você pode importá-los depois em Configurações → Arquivos de Dados → "Importar arquivo .pst…".':
    'You can import them later in Settings → Data Files → "Import .pst file…".',
  'Procurar contas do Outlook': 'Search Outlook accounts',
  'Procurar contas do Thunderbird': 'Search Thunderbird accounts',
  'Importar arquivo .json…': 'Import .json file…',
  'Quem copiou a': 'If you copied the',
  'pasta de dados': 'data folder',
  'na etapa anterior já tem tudo — não precisa importar nada aqui. O arquivo':
    'in the previous step you already have everything — no need to import anything here. The',
  'é só para quem não tem a pasta: ele é gerado pelo próprio Unit Mail no Windows (Configurações → Email → "Procurar contas" → "Exportar") e traz apenas as configurações de conta, sem senhas.':
    'file is only for those without the folder: it is generated by Unit Mail itself on Windows (Settings → Mail → "Search accounts" → "Export") and brings only account settings, no passwords.',
  'Selecionar todas': 'Select all',
  'adicionada': 'added',
  'Adicionando…': 'Adding…',
  'Adicionar {n} selecionada(s) →': 'Add {n} selected →',
  'Concluir →': 'Finish →',
  'Sua primeira conta de email': 'Your first email account',
  'Você pode pular e configurar depois em Configurações.':
    'You can skip and set up later in Settings.',
  'Seu nome (ex.: Roger Oliveira)': 'Your name (e.g.: John Smith)',
  'Email (ex.: voce@empresa.com)': 'Email (e.g.: you@company.com)',
  'Senha (opcional agora — pode configurar depois)': 'Password (optional now — can set up later)',
  'Onde as mensagens serão entregues': 'Where messages will be delivered',
  'Nova base na pasta de dados do app (padrão)': 'New store in the app data folder (default)',
  'Nova base em pasta à minha escolha…': 'New store in a folder of my choice…',
  'Sem base de dados por enquanto': 'No data store for now',
  'Cria a base da conta dentro da pasta de dados (Mail/) — os emails ficam no formato próprio do app.':
    'Creates the account store inside the data folder (Mail/) — emails stay in the app own format.',
  'Você escolhe a pasta onde a nova base de dados será criada.':
    'You pick the folder where the new data store is created.',
  'Sem entrega automática — você pode vincular uma base depois.':
    'No automatic delivery — you can bind a store later.',
  'Tem um arquivo .pst do Outlook? Depois use Configurações → Arquivos de Dados → "Importar arquivo .pst…" para trazer tudo para o formato do app.':
    'Have an Outlook .pst file? Later use Settings → Data Files → "Import .pst file…" to bring everything into the app format.',
  'Criando…': 'Creating…',
  'Criar conta': 'Create account',
  'Pular →': 'Skip →',
  'Tudo pronto!': 'All done!',
  'Use': 'Use',
  'no ribbon para gerenciar contas, arquivos de dados, regras e assinaturas a qualquer momento.':
    'on the ribbon to manage accounts, data files, rules and signatures anytime.',
  'Começar a usar': 'Start using',
  '✓ Esta já é a pasta de dados atual — pode continuar.':
    '✓ This is already the current data folder — you can continue.',
  '⚠ Esta pasta já contém um perfil do Unit Mail — restaure ou escolha outra pasta.':
    '⚠ This folder already contains a Unit Mail profile — restore it or pick another folder.',
  'Pasta selecionada — ainda nada foi gravado. Confirme em "Continuar".':
    'Folder selected — nothing saved yet. Confirm with "Continue".',
  '✓ Restaurando perfil — o app vai reiniciar…': '✓ Restoring profile — the app will restart…',
  'Procurando contas do {name}…': 'Searching {name} accounts…',
  'Não foi possível ler as contas do {name} neste sistema.':
    'Could not read {name} accounts on this system.',
  'Nenhum arquivo selecionado.': 'No file selected.',
  'Informe um email válido.': 'Enter a valid email.',
  'Seleção de pasta cancelada — a conta não foi criada.':
    'Folder selection cancelled — the account was not created.',
  'Conta criada, mas a conexão falhou: {msg}. Confira a senha nas Configurações.':
    'Account created, but the connection failed: {msg}. Check the password in Settings.',
  'Erro ao criar conta: {msg}': 'Error creating account: {msg}',

  // ---- Calendário, Contatos e Tarefas ----
  'Organizador': 'Organizer',
  'Obrigatório': 'Required',
  'Participantes': 'Attendees',
  'Recorrência': 'Recurrence',
  'Mês anterior': 'Previous month',
  'Ir para hoje': 'Go to today',
  'Próximo mês': 'Next month',
  'Meus Calendários': 'My Calendars',
  'Nenhum calendário aberto': 'No calendar open',
  'Pesquisar Calendário': 'Search Calendar',
  '{subject} — duplo clique abre os detalhes': '{subject} — double-click opens details',
  '+{n} mais': '+{n} more',
  'Pesquisar contatos…': 'Search contacts…',
  '{n} contato(s)': '{n} contact(s)',
  ' em {folder}': ' in {folder}',
  'Nenhum contato nesta pasta': 'No contacts in this folder',
  'Nenhuma tarefa nesta pasta': 'No tasks in this folder',
  'Selecione uma pasta de tarefas': 'Select a tasks folder',
  'Vence {date}': 'Due {date}',
  'Sem vencimento': 'No due date',
  // App — erros e avisos
  'Erro ao salvar anexo: {msg}': 'Error saving attachment: {msg}',
  'Erro ao abrir anexo: {msg}': 'Error opening attachment: {msg}',
  '{n} de {total} anexos salvos em "{dir}"': '{n} of {total} attachments saved to "{dir}"',
  'Erro ao salvar anexos: {msg}': 'Error saving attachments: {msg}',
  '{n} item(ns) copiado(s) — o arquivo PST é somente leitura, o original foi preservado.': '{n} item(s) copied — the PST file is read-only, the original was preserved.',
  'Erro ao imprimir: {msg}': 'Error printing: {msg}',
  'Erro ao ler pasta: {msg}': 'Error reading folder: {msg}',
  'Erro no envio/recebimento: {errs}': 'Send/receive error: {errs}',
  'Erro no envio/recebimento: {msg}': 'Send/receive error: {msg}',
  '"{subj}" enviada para {folder}.': '"{subj}" moved to {folder}.',
  'Mensagem': 'Message',
  'mensagem': 'message',
  'Falha ao enviar "{subj}": {err} — a mensagem ficou em {folder}.': 'Failed to send "{subj}": {err} — the message stays in {folder}.',
  'Lembrete: {subj}{quando}{local}': 'Reminder: {subj}{quando}{local}',
  'Erro ao ler item: {msg}': 'Error reading item: {msg}',
  'Erro ao abrir mensagem: {msg}': 'Error opening message: {msg}',
  'Sincronização parcial: {errs}': 'Partial sync: {errs}',
  'Erro ao excluir: {msg}': 'Error deleting: {msg}',
  'Erro ao mover pasta: {msg}': 'Error moving folder: {msg}',
  'Erro ao criar pasta: {msg}': 'Error creating folder: {msg}',
  'Erro ao excluir pasta: {msg}': 'Error deleting folder: {msg}',
  '"{folder}" é uma pasta padrão e não pode ser excluída.': '"{folder}" is a default folder and cannot be deleted.',
  'Erro ao ocultar pasta: {msg}': 'Error hiding folder: {msg}',
  'Erro ao copiar pasta: {msg}': 'Error copying folder: {msg}',
  '{n} item(ns) removido(s) da pasta.': '{n} item(s) removed from the folder.',
  'A pasta já está vazia.': 'The folder is already empty.',
  'Erro ao limpar pasta: {msg}': 'Error emptying folder: {msg}',
  'Erro ao marcar como lidas: {msg}': 'Error marking as read: {msg}',
  'Backup salvo em: {path}': 'Backup saved to: {path}',
  'Erro ao exportar: {msg}': 'Error exporting: {msg}',
  'Backup importado: {parts}.{warn}': 'Backup imported: {parts}.{warn}',
  'Atenção: {w}': 'Warning: {w}',
  'Erro ao importar backup: {msg}': 'Error importing backup: {msg}',
  'Mensagem salva em: {path}': 'Message saved to: {path}',
  'O arquivo de dados principal não está aberto — abra-o em Configurações → Arquivos de Dados.': 'The primary data file is not open — open it in Settings → Data Files.',
  'Executando regras na pasta…': 'Running rules in the folder…',
  '{n} mensagem(ns) movida(s) pelas regras.': '{n} message(s) moved by rules.',
  'Nenhuma mensagem correspondeu às regras.': 'No messages matched the rules.',
  'Erro ao executar regras: {msg}': 'Error running rules: {msg}',
  'Erro: {err}': 'Error: {err}',
  'Erro ao sincronizar: {msg}': 'Error syncing: {msg}',
  'Enviando resposta ao organizador…': 'Sending response to the organizer…',
  'Convite aceito — compromisso adicionado ao Calendário.': 'Invitation accepted — appointment added to Calendar.',
  'Convite recusado — resposta enviada ao organizador.': 'Invitation declined — response sent to the organizer.',
  'Nova proposta de horário enviada ao organizador.': 'New time proposal sent to the organizer.',
  'Erro ao responder convite: {msg}': 'Error responding to invitation: {msg}',
  'Excluindo {n} itens…': 'Deleting {n} items…',
  'Excluindo…': 'Deleting…',
  'Carregando bases…': 'Loading stores…',
  'Movendo pasta…': 'Moving folder…',
  'Movendo {n} itens…': 'Moving {n} items…',
  'Movendo…': 'Moving…',
  // Progresso de envio/recebimento
  'Cancelado': 'Cancelled',
  'Aguardando': 'Waiting',
  'Processando': 'Processing',
  '{n} item(ns)': '{n} item(s)',
  'Progresso de envio/recebimento no Unit Mail': 'Unit Mail Send/Receive Progress',
  '{done} de {total} tarefas foram concluídas': '{done} of {total} tasks completed',
  '{n} com erro': '{n} with errors',
  'Erros': 'Errors',
  'Progresso': 'Progress',
  'Restante': 'Remaining',
  'Nenhuma tarefa em execução.': 'No tasks running.',
  'Nenhum erro nesta execução.': 'No errors in this run.',
  'Cancelar "{label}"': 'Cancel "{label}"',
  'Selecione uma tarefa em execução na lista': 'Select a running task in the list',
  'Cancelar tarefa': 'Cancel task',
  'Cancelar tudo': 'Cancel All',
  'Resumir': 'Compact',
  'Ocultar': 'Hide',
  // Novo item (compromisso / contato / tarefa)
  'Informe o nome do contato.': 'Enter the contact name.',
  'Informe o assunto.': 'Enter a subject.',
  'O arquivo de dados principal não está aberto. Abra-o ou defina outro em Configurações → Arquivos de Dados.': 'The primary data file is not open. Open it or set another in Settings → Data Files.',
  'Arquivo de dados do item não está aberto': "The item's data file is not open",
  'Arquivo de dados principal não está aberto': 'Primary data file is not open',
  'Nome completo': 'Full name',
  'Sobrenome': 'Last name',
  'Departamento': 'Department',
  'Telefone comercial': 'Business phone',
  'Celular': 'Mobile',
  'Participantes (emails, separados por ;)': 'Attendees (emails, separated by ;)',
  'Evento de dia inteiro': 'All day event',
  'Sem lembrete': 'No reminder',
  'Na hora': 'On time',
  '5 minutos antes': '5 minutes before',
  '15 minutos antes': '15 minutes before',
  '30 minutos antes': '30 minutes before',
  '1 hora antes': '1 hour before',
  '1 dia antes': '1 day before',
  'Vencimento': 'Due date',
  '% concluída': '% complete',
  'Tarefa concluída': 'Task completed',
  'Anotações': 'Notes',
  'Salvando…': 'Saving…',
  // Categorias
  'Todas as Categorias': 'All Categories',
  'Cor da categoria — clique para trocar': 'Category color — click to change',
  'Excluir categoria': 'Delete category',
  'Nova categoria': 'New category',
  'Nenhuma categoria — clique em "Nova" para criar.': 'No categories — click "New" to create one.',
  // Editor de texto rico
  'Tamanho da fonte': 'Font size',
  'Imagem selecionada': 'Selected image',
  'Arraste a alça do canto inferior direito para redimensionar.': 'Drag the bottom-right grip to resize.',
  'Arraste para redimensionar': 'Drag to resize',
  // Painel de leitura
  'Para proteger sua privacidade, as imagens externas desta mensagem foram bloqueadas. Ative em Configurações de Conta → Geral.': 'To protect your privacy, external images in this message were blocked. Enable them in Account Settings → General.',
  '(sem conteúdo)': '(no content)',
  'Tel. Comercial': 'Business Phone',
  'Tel. Residencial': 'Home Phone',
  'Endereço Comercial': 'Business Address',
  'Endereço Residencial': 'Home Address',
  'Status': 'Status',
  'Concluída em': 'Completed on',
  'Responsável': 'Owner',
  'Não iniciada': 'Not started',
  'Em andamento': 'In progress',
  'Concluída': 'Completed',
  'Adiada': 'Deferred',
  'Esta reunião foi cancelada pelo organizador.': 'This meeting was cancelled by the organizer.',
  'Esta mensagem é uma resposta a um convite de reunião.': 'This message is a response to a meeting invitation.',
  'Convite aceito': 'Invitation accepted',
  'Convite recusado': 'Invitation declined',
  'Convite enviado': 'Invitation sent',
  'Salvar na pasta Rascunhos': 'Save to Drafts folder',
  'Erro ao salvar rascunho: {msg}': 'Error saving draft: {msg}',
  'Aceitar': 'Accept',
  'Recusar': 'Decline',
  'Cancelar proposta': 'Cancel proposal',
  'Propor novo horário…': 'Propose new time…',
  'Enviar proposta': 'Send proposal',
  'Nova proposta de horário enviada': 'New time proposal sent',
  'Selecione um item para ler': 'Select an item to read',
  'Mais ações': 'More actions',
  'Salvar como…': 'Save as…',
  'Imprimir…': 'Print…',
  'Opções do anexo': 'Attachment options',
  'Salvar todos': 'Save all',
  // Janela de mensagem (duplo clique)
  '(sem assunto)': '(no subject)',
  'Minimizar': 'Minimize',
  'Restaurar': 'Restore',
  'Maximizar': 'Maximize',
  // Tela de boas-vindas / composição
  'Selecione arquivos de dados ou configure suas contas': 'Select data files or set up your accounts',
  'Arquivos de dados…': 'Data files…',
  'Configurar contas…': 'Set up accounts…',
  'Selecione uma conta de envio.': 'Select a sending account.',
  // Convites de reunião
  'Convite enviado para {n} participante(s).': 'Invitation sent to {n} attendee(s).',
  'Compromisso salvo, mas nenhuma conta de email está configurada para enviar os convites.':
    'Appointment saved, but no email account is configured to send the invitations.',
  'Compromisso salvo, mas falha ao enviar convites: {msg}':
    'Appointment saved, but failed to send invitations: {msg}',
  'Convite atualizado enviado para {n} participante(s).':
    'Updated invitation sent to {n} attendee(s).',
  'As alterações serão enviadas aos participantes.':
    'Changes will be sent to attendees.',
  'Editar Compromisso': 'Edit Appointment',
  'Editar Contato': 'Edit Contact',
  'Editar Tarefa': 'Edit Task',
  'erro desconhecido': 'unknown error',
};
