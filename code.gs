/**
 * ARQUIVO: Code.gs
 * DESCRIÇÃO: Script principal que lida com o bot do Telegram, webhook e lógica de negócio.
 * VERSÃO: Inclui menus com botões inline para uma navegação mais intuitiva.
 */

// --- CONFIGURAÇÕES E UTILITÁRIOS GERAIS ---

// Nome da planilha onde estão todas as abas. Ajuste se o nome da sua planilha for diferente.
const SPREADSHEET_NAME = "Sistema Inteligente de Gestão de Vendas"; 

// Mapeamento dos nomes das abas para facilitar o acesso.
const SHEET_NAMES = {
  CONFIGURACOES: "CONFIGURACOES",
  VENDEDORES: "Vendedores_Revendedores", // ABA DEDICADA PARA VENDEDORES
  PALAVRAS_CHAVE: "PALAVRAS_CHAVE",
  ESTOQUE_PRONTO: "ESTOQUE_PRONTO",
  ESTOQUE_MATERIA: "ESTOQUE_MATERIA",
  ESTOQUE_CONSIGNACAO: "ESTOQUE_CONSIGNAÇÃO",
  VENDAS_LOG: "VENDAS_LOG",
  PRODUCAO_FILA: "PRODUCAO_FILA", // Aba antiga, pode ser depreciada ou mantida para histórico
  ESTOQUE_TECIDOS: "ESTOQUE_TECIDOS",
  MENSAGENS_PADRAO: "CONFIGURACOES",
  FLUXO_CAIXA: "FLUXO_CAIXA",
  // NOVAS ABAS PARA O MÓDULO DE PRODUÇÃO
  LOTES_PRODUCAO: "LOTES_PRODUCAO",
  ITENS_LOTE: "ITENS_LOTE"
};


// Objeto para armazenar as configurações carregadas da planilha.
let CONFIG = {};

// Propriedades do script para armazenar estados de conversação
const userProperties = PropertiesService.getUserProperties();

// Função para obter a planilha ativa ou uma planilha pelo nome.
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByName(SPREADSHEET_NAME);
}

// Função auxiliar para obter uma aba específica.
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`A aba "${sheetName}" não foi encontrada. Verifique o nome.`);
  }
  return sheet;
}

// Função para ler dados de uma aba, pulando a primeira linha (cabeçalho).
function readSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return []; // Retorna vazio se não houver dados além do cabeçalho
  return values.slice(1); // Ignora o cabeçalho
}

function getHeaders(sheetName) {
    const sheet = getSheet(sheetName);
    if (sheet.getLastRow() < 1) return [];
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// Função para carregar as configurações da aba 'CONFIGURACOES'.
function loadConfigurations() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'CONFIG_CACHE';
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    Logger.log("Carregando CONFIG do cache.");
    CONFIG = JSON.parse(cachedData);
    return;
  }
 
  Logger.log("Cache de CONFIG não encontrado. Lendo da planilha.");
  const configSheet = getSheet(SHEET_NAMES.CONFIGURACOES);
  const data = configSheet.getDataRange().getValues();
 
  CONFIG = {}; // Reinicia CONFIG para garantir que esteja atualizado a cada carregamento
  let currentSection = '';
  let skipNextRow = false;

  const sectionMap = {
    'SISTEMA (Tokens e URLs):': 'SISTEMA',
    'CHAT_IDS (Usuários e Grupos):': 'CHAT_IDS',
    'VENDEDORES:': 'VENDEDORES',
    'PREÇOS:': 'PRECOS',
    'ALERTAS:': 'ALERTAS',
    'EVENTOS:': 'EVENTOS',
    'PARAMETROS_SISTEMA:': 'PARAMETROS_SISTEMA',
    'MENSAGENS_PADRAO:': 'MENSAGENS_PADRAO',
    'DESCONTOS_ATACADO:': 'DESCONTOS_ATACADO'
  };

  data.forEach(row => {
    if (skipNextRow) {
      skipNextRow = false;
      return;
    }

    const key = String(row[0]).trim();
   
    if (sectionMap[key]) {
      currentSection = sectionMap[key];
      if (!CONFIG[currentSection]) {
        CONFIG[currentSection] = {};
        if (['CHAT_IDS', 'VENDEDORES', 'PRECOS', 'ALERTAS', 'EVENTOS', 'MENSAGENS_PADRAO', 'DESCONTOS_ATACADO'].includes(currentSection)) {
          CONFIG[currentSection]['lista'] = [];
          skipNextRow = true;
        }
      }
    } 
    else if (currentSection && key) { 
      if (currentSection === 'SISTEMA' || currentSection === 'PARAMETROS_SISTEMA') {
        CONFIG[currentSection][key] = row[1];
      } else if (currentSection === 'CHAT_IDS') {
        if (key === 'chatId') {
          CONFIG[currentSection].lista.push({ id: row[1], nome: row[2], grupo: row[3] });
        } else {
          CONFIG[currentSection][key] = row[1];
        }
      } else if (currentSection === 'VENDEDORES') {
        // NÃO FAZ NADA AQUI - Esta seção foi movida para ser lida da aba Vendedores_Revendedores
        // Isso previne que a lista de vendedores na aba CONFIGURACOES seja usada
      } else if (currentSection === 'PRECOS') {
        CONFIG[currentSection].lista.push({
          modelo: key,
          preco_base: row[1],
          margem: row[2],
          preco_venda: row[3],
          categoria: row[4]
        });
      } else if (currentSection === 'ALERTAS') {
        CONFIG[currentSection].lista.push({
          tipo_alerta: key,
          produto: row[1],
          estoque_minimo: row[2],
          notificar_chat: row[3],
          ativo: row[4]
        });
      } else if (currentSection === 'EVENTOS') {
        CONFIG[currentSection].lista.push({
          nome: key,
          data: row[1],
          local: row[2],
          vendedores: row[3],
          status: row[4],
          meta_vendas: row[5],
          produtos_foco: row[6]
        });
      } else if (currentSection === 'MENSAGENS_PADRAO') {
        CONFIG[currentSection].lista.push({
          tipo_mensagem: key,
          texto: row[1],
          usar_emoji: row[2]
        });
      } else if (currentSection === 'DESCONTOS_ATACADO') { 
        CONFIG[currentSection].lista.push({
          faixa_min: parseInt(row[0], 10),
          faixa_max: parseInt(row[1], 10),
          desconto_percentual: String(row[2]).trim(),
          condicao_pagamento: String(row[3]).trim().toUpperCase()
        });
      }
    }
  });

  // --- INÍCIO DA CORREÇÃO ---
  // Carrega a lista de Vendedores/Revendedores da aba dedicada
  // Isso centraliza a fonte da verdade para o Bot e para o Dashboard
  Logger.log("Carregando VENDEDORES da aba Vendedores_Revendedores...");
  try {
    const vendedoresSheet = getSheet(SHEET_NAMES.VENDEDORES);
    const vendedoresData = vendedoresSheet.getDataRange().getValues();
    const vendedoresHeaders = vendedoresData.shift(); // Pega cabeçalho

    const nomeIdx = vendedoresHeaders.indexOf('Nome');
    const telegramIdIdx = vendedoresHeaders.indexOf('Telegram ID');
    const permissoesIdx = vendedoresHeaders.indexOf('Permissões');
    const statusIdx = vendedoresHeaders.indexOf('Status');
    // Trata ambos os nomes de coluna "Comissão" ou "Comissao"
    const comissaoIdx = vendedoresHeaders.indexOf('Comissão') !== -1 ? vendedoresHeaders.indexOf('Comissão') : vendedoresHeaders.indexOf('Comissao');
    const metaIdx = vendedoresHeaders.indexOf('Meta Mensal');

    if (!CONFIG.VENDEDORES) CONFIG.VENDEDORES = {};
    CONFIG.VENDEDORES.lista = []; // Limpa qualquer lista antiga vinda da aba CONFIGURACOES

    vendedoresData.forEach(row => {
        const status = String(row[statusIdx]).trim().toLowerCase();
        if (status === 'ativo') { // Carrega apenas vendedores ativos
            const telegramId = String(row[telegramIdIdx]).trim();
            // Só adiciona se tiver um Telegram ID, essencial para o Bot
            if (telegramId) { 
              CONFIG.VENDEDORES.lista.push({
                  nome: String(row[nomeIdx]).trim(),
                  telegram_id: telegramId,
                  permissoes: String(row[permissoesIdx]).trim().toLowerCase() || 'vendedor', // Padrão 'vendedor'
                  status: status,
                  comissao: parseFloat(String(row[comissaoIdx]).replace(',', '.')) || 0,
                  meta_mensal: parseFloat(String(row[metaIdx]).replace(',', '.')) || 0
              });
            }
        }
    });
     Logger.log(`Carregados ${CONFIG.VENDEDORES.lista.length} vendedores ativos (com Telegram ID) da aba Vendedores_Revendedores.`);
  } catch (e) {
      Logger.log(`ERRO ao carregar Vendedores da aba Vendedores_Revendedores: ${e.message}`);
      // Se falhar, inicializa uma lista vazia para evitar que o resto do bot quebre
      if (!CONFIG.VENDEDORES) CONFIG.VENDEDORES = { lista: [] };
  }
  // --- FIM DA CORREÇÃO ---

  const cacheDuration = (parseInt(CONFIG.PARAMETROS_SISTEMA?.CACHE_DURACAO, 10) || 5) * 60; // Em segundos
  cache.put(cacheKey, JSON.stringify(CONFIG), cacheDuration);
  Logger.log(`CONFIG armazenado no cache por ${cacheDuration / 60} minutos.`);
}


// Função para carregar o dicionário de PALAVRAS_CHAVE.
function loadKeywords() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'KEYWORDS_CACHE';
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    Logger.log("Carregando KEYWORDS do cache.");
    return JSON.parse(cachedData);
  }
 
  Logger.log("Cache de KEYWORDS não encontrado. Lendo da planilha.");
  const keywordSheet = getSheet(SHEET_NAMES.PALAVRAS_CHAVE);
  const data = keywordSheet.getDataRange().getValues().slice(1);
  const keywords = {};
  data.forEach(row => {
    const palavraUsuario = String(row[0]).toLowerCase().trim();
    if (!palavraUsuario) return;
    const palavraSistema = String(row[1]).trim();
    const categoria = String(row[2]).trim();
    const prioridade = parseInt(row[3], 10) || 99;
    if (!keywords[categoria]) {
      keywords[categoria] = [];
    }
    keywords[categoria].push({ palavraUsuario, palavraSistema, prioridade });
  });

  // A variável global CONFIG já deve ter sido carregada em doPost
  const cacheDuration = (parseInt(CONFIG.PARAMETROS_SISTEMA?.CACHE_DURACAO, 10) || 5) * 60; // Em segundos
  cache.put(cacheKey, JSON.stringify(keywords), cacheDuration);
  Logger.log(`KEYWORDS armazenado no cache por ${cacheDuration / 60} minutos.`);
 
  return keywords;
}

// Variável global para armazenar as palavras-chave carregadas.
let KEYWORDS = {};

// --- FUNÇÕES DE INTERAÇÃO COM O TELEGRAM API ---

// Função principal para enviar mensagens ao Telegram.
function sendTelegramMessage(chat_id, text, reply_markup = null) {
  if (!CONFIG || !CONFIG.SISTEMA) loadConfigurations();
  const telegramToken = CONFIG.SISTEMA.TELEGRAM_TOKEN;
  if (!telegramToken) {
    Logger.log("TELEGRAM_TOKEN não configurado na aba CONFIGURACOES.");
    return;
  }

  const url = `https://api.telegram.org/bot${telegramToken}/sendMessage`;
 
  const payload = {
    chat_id: String(chat_id),
    text: text,
    parse_mode: 'HTML'
  };

  if (reply_markup !== null) {
    payload.reply_markup = reply_markup;
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  try {
    UrlFetchApp.fetch(url, options);
    Logger.log(`Mensagem enviada com sucesso para chat_id: ${chat_id}`);
  } catch (e) {
    Logger.log(`Erro ao enviar mensagem ao Telegram: ${e.message}`);
  }
}

// Função para editar mensagens existentes.
function editTelegramMessage(chat_id, message_id, text, reply_markup = null) {
  if (!CONFIG || !CONFIG.SISTEMA) loadConfigurations();
  const telegramToken = CONFIG.SISTEMA.TELEGRAM_TOKEN;
  if (!telegramToken) {
    Logger.log("TELEGRAM_TOKEN não configurado na aba CONFIGURACOES.");
    return;
  }

  const url = `https://api.telegram.org/bot${telegramToken}/editMessageText`;
 
  const payload = {
    chat_id: String(chat_id),
    message_id: message_id,
    text: text,
    parse_mode: 'HTML'
  };

  if (reply_markup !== null) {
    payload.reply_markup = reply_markup;
  }
  
  Logger.log(`PAYLOAD para editTelegramMessage: ${JSON.stringify(payload)}`);

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if(response.getResponseCode() !== 200) {
       Logger.log(`Erro ao editar mensagem (Código: ${response.getResponseCode()}): ${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`Exceção ao editar mensagem do Telegram: ${e.message}`);
  }
}

// Função para responder a um callback (clique em botão).
function answerCallbackQuery(callback_query_id, text = "") {
  try {
    if (!CONFIG || !CONFIG.SISTEMA) loadConfigurations();
    const telegramToken = CONFIG.SISTEMA.TELEGRAM_TOKEN;
    if (!telegramToken) {
      Logger.log("TELEGRAM_TOKEN não configurado para answerCallbackQuery.");
      return;
    }
    const url = `https://api.telegram.org/bot${telegramToken}/answerCallbackQuery`;
    const payload = {
      callback_query_id: callback_query_id,
      text: text,
      show_alert: false
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true 
    };
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log(`Info: Erro ao executar answerCallbackQuery (pode ser ignorado em testes): ${e.message}`);
  }
}

// --- PONTO DE ENTRADA E ROTEAMENTO ---

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      Logger.log("Requisição recebida sem dados válidos.");
      return;
    }
    
    const contents = JSON.parse(e.postData.contents);
    Logger.log(`Update Recebido: ${JSON.stringify(contents)}`);
    
    loadConfigurations();
    KEYWORDS = loadKeywords(); 
    
    if (contents.callback_query) {
      handleCallbackQuery(contents.callback_query);
    } else if (contents.message) {
      const { chat, text, from } = contents.message;
      const user = { id: from.id, name: from.first_name || 'Desconhecido' };
      
      const vendedorInfo = CONFIG.VENDEDORES.lista.find(v => String(v.telegram_id) === String(from.id));
      if (!vendedorInfo) {
        sendTelegramMessage(chat.id, "🚫 Acesso Negado. Você não tem permissão para usar este bot.");
        return;
      }
      processCommand(chat.id, text, vendedorInfo.nome, user);
    }
  } catch (error) {
    Logger.log(`Erro fatal no doPost: ${error.stack}`);
  }
}

function handleCallbackQuery(callback_query) {
    const { from, message, data } = callback_query;
    const user = { id: from.id, name: from.first_name || 'Desconhecido' };
    
    answerCallbackQuery(callback_query.id);

    // Roteador para o fluxo de venda
    if (data.startsWith('venda_')) {
        handleSaleFlow(callback_query);
        return;
    }

    // NOVO: Roteador para o fluxo de criação de lote
    if (data.startsWith('lote_')) {
        handleLoteFlow(callback_query);
        return;
    }

    if (data.startsWith('onboarding_')) {
        handleOnboardingFlow(callback_query);
        return;
    }
    
    if (data === 'menu_principal') {
        sendMainMenu(message.chat.id, message.message_id);
        return;
    }

    // Roteamento para outros menus
    const menuActions = {
        'menu_estoque': showEstoqueMenu,
        'menu_producao': showProducaoMenu,
        'menu_relatorios': showRelatoriosMenu,
        'menu_caixa': showCaixaMenu,
        'menu_consignacao': showConsignacaoMenu
    };

    if (menuActions[data]) {
        menuActions[data](message.chat.id, message.message_id, user.id);
        return;
    }
    
    // Roteamento para ações que pedem input do utilizador
    const inputActions = {
        'action_venda': () => startSaleFlow(user),
        'action_consultar_estoque': () => setupNextAction(user.id, 'consultar_estoque', "🔍 **Consultar Estoque**\n\nDigite o nome do produto ou matéria-prima que deseja consultar."),
        'action_add_materia': () => setupNextAction(user.id, 'add_materia', "➕ **Adicionar Matéria-Prima**\n\nDigite os detalhes no formato:\n`add materia <qtd> <descrição completa>`"),
        'action_concluir_lote': () => setupNextAction(user.id, 'concluir_lote', "✅ **Concluir Lote**\n\nDigite o ID do lote a concluir:\n`/concluir lote LOTE-240115-123456`"),
        'action_cancelar_lote': () => setupNextAction(user.id, 'cancelar_lote', "❌ **Cancelar Lote**\n\nDigite o ID do lote a ser cancelado:\n`cancelar lote LOTE-240115-123456`"),
        'action_iniciar_lote': () => iniciarNovoLote(user), // ATUALIZADO
        'consignar_enviar': () => setupNextAction(user.id, 'consignar_enviar', "🚚 **Enviar para Consignação**\n\nDigite no formato:\n`consignar <qtd> <produto> para <revendedor>`"),
        'consignar_venda': () => setupNextAction(user.id, 'consignar_venda', "💰 **Venda Consignada**\n\nDigite no formato:\n`venda c <qtd> <produto> de <revendedor>`"),
        'consignar_retorno': () => setupNextAction(user.id, 'consignar_retorno', "↩️ **Retorno de Consignação**\n\nDigite no formato:\n`retorno c <qtd> <produto> de <revendedor>`"),
        'consignar_estoque': () => setupNextAction(user.id, 'consignar_estoque', "🔍 **Consultar Estoque Consignado**\n\nDigite no formato:\n`estoque c <revendedor>`")
    };

    if (inputActions[data]) {
        inputActions[data]();
        return;
    }

    // Ações de Relatórios
    if (data.startsWith('report_')) {
        const reportType = data.split('_')[1];
        let command = "relatorio ";
        if (reportType === 'dia') command += "do dia";
        else if (reportType === 'semana') command += "vendas semana";
        else if (reportType === 'top') command += "top produtos";
        else if (reportType === 'producao') command += "producao";
        handleReportQueryCommand(message.chat.id, command);
        return;
    }

    sendTelegramMessage(message.chat.id, `Funcionalidade para '${data}' ainda não implementada.`);
}

function processCommand(chat_id, commandText, vendedorNome, user) {
    const lowerCommand = (commandText || '').toLowerCase();
    const nextAction = userProperties.getProperty('next_action_' + user.id);
    
    if (nextAction) {
        // Limpa a ação ANTES do uso para evitar loops
        userProperties.deleteProperty('next_action_' + user.id); 
        const actionHandlers = {
            'caixa': () => handleCashFlowCommand(chat_id, commandText, vendedorNome),
            'consultar_estoque': () => handleStockQueryCommand(chat_id, commandText),
            'add_materia': () => handleMateriaPrimaCommand(chat_id, commandText, vendedorNome),
            'concluir_lote': () => concluirLoteDeProducao(commandText, user),
            'cancelar_lote': () => cancelarLoteDeProducao(commandText, user),
            'consignar_enviar': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
            'consignar_venda': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
            'consignar_retorno': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
            'consignar_estoque': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
            'lote_descricao': () => {
                const stateJSON = userProperties.getProperty('state_' + user.id);
                if (!stateJSON) return;
                let state = JSON.parse(stateJSON);
                state.details.descricao = commandText;
                state.step = 'menu_itens';
                userProperties.setProperty('state_' + user.id, JSON.stringify(state));
                askLoteQuestion(user, state, null); // message_id é nulo porque estamos numa nova mensagem
            },
            // CORREÇÃO: Adiciona handlers para estampa e quantidade digitada
            'lote_estampa': () => {
                const stateJSON = userProperties.getProperty('state_' + user.id);
                if (!stateJSON) return;
                let state = JSON.parse(stateJSON);
                state.currentItem.estampa = commandText;
                state.step = 'get_quantidade';
                userProperties.setProperty('state_' + user.id, JSON.stringify(state));
                askLoteQuestion(user, state, null);
            },
            'lote_quantidade': () => {
                const stateJSON = userProperties.getProperty('state_' + user.id);
                if (!stateJSON) return;
                let state = JSON.parse(stateJSON);
                const qty = parseInt(commandText, 10);
                if (isNaN(qty) || qty <= 0) {
                    sendTelegramMessage(chat_id, "Quantidade inválida. Por favor, digite um número maior que zero.");
                    setupNextAction(user.id, 'lote_quantidade', "Digite a quantidade desejada:");
                    return;
                }
                state.currentItem.quantidade = qty;
                const ci = state.currentItem;
                const produtoCompleto = `${ci.modelo} ${ci.genero || ''} ${ci.cor} ${ci.cor_manga ? 'Manga ' + ci.cor_manga : ''} ${ci.tamanho} ${ci.estampa}`.replace(/\s+/g, ' ').trim();
                state.details.itens.push({ produto: produtoCompleto, quantidade: ci.quantidade });
                delete state.currentItem;
                state.step = 'menu_itens';
                userProperties.setProperty('state_' + user.id, JSON.stringify(state));
                askLoteQuestion(user, state, null);
            }
        };
        if (actionHandlers[nextAction]) {
            actionHandlers[nextAction]();
            return;
        }
    }
    
    if (lowerCommand === '/start' || lowerCommand === '/ajuda') {
        const onboardingCompleted = userProperties.getProperty('onboarding_completed_' + user.id);
        if (!onboardingCompleted) {
            startOnboarding(user);
        } else {
            sendMainMenu(chat_id);
        }
        return;
    }

    // Fallback para comandos de texto genéricos que não dependem de um estado anterior
    const commandHandlers = {
        'vendi': () => handleSaleCommand(chat_id, commandText, vendedorNome),
        'tem': () => handleStockQueryCommand(chat_id, commandText),
        'add materia': () => handleMateriaPrimaCommand(chat_id, commandText, vendedorNome),
        'entrada': () => handleCashFlowCommand(chat_id, commandText, vendedorNome),
        'saida': () => handleCashFlowCommand(chat_id, commandText, vendedorNome),
        'concluir lote': () => concluirLoteDeProducao(commandText, user),
        'cancelar lote': () => cancelarLoteDeProducao(commandText, user),
        'consignar': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
        'venda c': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
        'retorno c': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
        'estoque c': () => handleConsignmentCommand(chat_id, commandText, vendedorNome),
        'relatorio': () => handleReportQueryCommand(chat_id, lowerCommand),
    };

    for (const key in commandHandlers) {
        if (lowerCommand.startsWith(key)) {
            commandHandlers[key]();
            return;
        }
    }

    sendTelegramMessage(chat_id, "Comando não reconhecido. Use /ajuda para ver as opções.");
}


function setupNextAction(userId, action, message) {
    userProperties.setProperty('next_action_' + userId, action);
    sendTelegramMessage(userId, message);
}

// --- FLUXOS DE CONVERSA E MENUS ---

function startOnboarding(user) {
    const text = "👋 **Bem-vindo(a) ao Gestor PRO!**\n\nEu sou seu assistente virtual para gerir vendas, estoque e produção.\n\nVamos fazer um tour rápido pelas principais funcionalidades?";
    const keyboard = { inline_keyboard: [[{ text: "🚀 Sim, vamos lá!", callback_data: "onboarding_next_2" }]] };
    sendTelegramMessage(user.id, text, keyboard);
}

function handleOnboardingFlow(callback_query) {
  const { from, message, data } = callback_query;
  const user = { id: from.id, name: from.first_name };
  const { chat, message_id } = message;
  
  Logger.log(`Iniciando handleOnboardingFlow. User: ${user.id}, Chat: ${chat.id}, Msg: ${message_id}, Data: ${data}`);

  const nextStep = parseInt(data.split('_').pop(), 10);
  let text = "";
  let keyboard = { inline_keyboard: [] };

  switch (nextStep) {
    case 2:
      text = "🛒 **Vendas Rápidas**\n\nCom o botão 'Venda Rápida', você será guiado passo a passo para registrar uma venda. Esqueça os comandos complicados!";
      keyboard.inline_keyboard.push([{ text: "Próximo Passo ➡️", callback_data: "onboarding_next_3" }]);
      break;
    case 3:
      text = "📦 **Gestão de Estoque e Produção**\n\nConsulte o estoque de qualquer item, adicione matéria-prima e inicie lotes de produção, tudo pelos menus.";
      keyboard.inline_keyboard.push([{ text: "Entendi! Ir para o Menu ➡️", callback_data: "onboarding_next_4" }]);
      break;
    case 4:
      userProperties.setProperty('onboarding_completed_' + user.id, 'true');
      sendMainMenu(chat.id, message_id);
      return;
  }
  
  editTelegramMessage(chat.id, message_id, text, keyboard);
}


function sendMainMenu(chat_id, message_id = null) {
  const text = "Olá! 👋 Sou seu assistente de gestão. Escolha uma opção abaixo:";
  const keyboard = getMainMenuKeyboard(chat_id);
  
  if (message_id) {
    editTelegramMessage(chat_id, message_id, text, keyboard);
  } else {
    sendTelegramMessage(chat_id, text, keyboard);
  }
}

function getMainMenuKeyboard(chat_id) {
    if(!CONFIG || !CONFIG.VENDEDORES) loadConfigurations();
    const user = CONFIG.VENDEDORES.lista.find(v => String(v.telegram_id) === String(chat_id));
    const permissions = user ? user.permissoes : '';

    let keyboard = {
        inline_keyboard: [
            [{ text: "📦 Estoque", callback_data: "menu_estoque" }, { text: "🛒 Venda Rápida", callback_data: "action_venda" }],
            [{ text: "🏭 Produção", callback_data: "menu_producao" }, { text: "🚚 Consignação", callback_data: "menu_consignacao" }],
            [{ text: "📊 Relatórios", callback_data: "menu_relatorios" }, { text: "💰 Fluxo de Caixa", callback_data: "menu_caixa" }]
        ]
    };
    
    if (permissions === 'admin') {
        const dashboardUrl = ScriptApp.getService().getUrl();
        keyboard.inline_keyboard.push([{ text: "📊 Abrir Painel de Gestão", url: dashboardUrl }]);
    }
    return keyboard;
}

// Funções para mostrar submenus
function showEstoqueMenu(chat_id, message_id, user_id) {
    const text = "📦 **Estoque**\nO que você gostaria de fazer?";
    const keyboard = { inline_keyboard: [
        [{ text: "Consultar Estoque", callback_data: "action_consultar_estoque" }],
        [{ text: "Adicionar Matéria-Prima", callback_data: "action_add_materia" }],
        [{ text: "⬅️ Voltar ao Menu", callback_data: "menu_principal" }]
    ]};
    editTelegramMessage(chat_id, message_id, text, keyboard);
}

function showProducaoMenu(chat_id, message_id, user_id) {
    const text = "🏭 **Produção**\nO que você gostaria de fazer?";
    const keyboard = { inline_keyboard: [
        [{ text: "Iniciar Novo Lote", callback_data: "action_iniciar_lote" }],
        [{ text: "Concluir Lote (por ID)", callback_data: "action_concluir_lote" }],
        [{ text: "❌ Cancelar Lote (por ID)", callback_data: "action_cancelar_lote" }],
        [{ text: "⬅️ Voltar ao Menu", callback_data: "menu_principal" }]
    ]};
    editTelegramMessage(chat_id, message_id, text, keyboard);
}

function showRelatoriosMenu(chat_id, message_id, user_id) {
    const text = "📊 **Relatórios**\nQual relatório você deseja gerar?";
    const keyboard = { inline_keyboard: [
        [{ text: "Vendas do Dia", callback_data: "report_dia" }, { text: "Vendas da Semana", callback_data: "report_semana" }],
        [{ text: "Top Produtos", callback_data: "report_top" }, { text: "Lotes Pendentes", callback_data: "report_producao" }],
        [{ text: "⬅️ Voltar ao Menu", callback_data: "menu_principal" }]
    ]};
    editTelegramMessage(chat_id, message_id, text, keyboard);
}

function showCaixaMenu(chat_id, message_id, user_id) {
    const text = "💰 **Fluxo de Caixa**\n\nPara registar um movimento, digite no formato:\n`entrada <valor> <descrição>`\nou\n`saida <valor> <descrição>`";
    const keyboard = { inline_keyboard: [[{ text: "⬅️ Voltar ao Menu", callback_data: "menu_principal" }]] };
    editTelegramMessage(chat_id, message_id, text, keyboard);
    userProperties.setProperty('next_action_' + user_id, 'caixa');
}


function showConsignacaoMenu(chat_id, message_id, user_id) {
    const text = "🚚 **Consignação**\nO que você gostaria de fazer?";
    const keyboard = { inline_keyboard: [
        [{ text: "Enviar para Revendedor", callback_data: "consignar_enviar" }],
        [{ text: "Registrar Venda Consignada", callback_data: "consignar_venda" }],
        [{ text: "Registrar Retorno", callback_data: "consignar_retorno" }],
        [{ text: "Consultar Estoque de Revendedor", callback_data: "consignar_estoque" }],
        [{ text: "⬅️ Voltar ao Menu", callback_data: "menu_principal" }]
    ]};
    editTelegramMessage(chat_id, message_id, text, keyboard);
}


// --- FLUXO DE VENDA CONVERSACIONAL ---

function startSaleFlow(user) {
    const state = {
        flow: 'venda',
        step: 'get_model',
        details: {}
    };
    userProperties.setProperty('state_' + user.id, JSON.stringify(state));
    askSaleQuestion(user, state);
}

function handleSaleFlow(callback_query) {
    const user = { id: callback_query.from.id, name: callback_query.from.first_name || 'Desconhecido' };
    const chat_id = callback_query.message.chat.id;
    const data = callback_query.data;

    const stateJSON = userProperties.getProperty('state_' + user.id);
    if (!stateJSON) {
        sendTelegramMessage(chat_id, "Ocorreu um erro ou a sua sessão expirou. Por favor, comece de novo.");
        sendMainMenu(chat_id);
        return;
    }
    let state = JSON.parse(stateJSON);

    const [flow, action, ...valueParts] = data.split('_');
    const value = valueParts.join('_');

    if (action === 'set') {
        const field = valueParts[0];
        const fieldValue = valueParts.slice(1).join('_');

        state.details[field] = fieldValue;

        // Avança para o próximo passo
        if (field === 'modelo') state.step = 'get_cor';
        else if (field === 'cor') state.step = 'get_tamanho';
        else if (field === 'tamanho') state.step = 'get_estampa';
        else if (field === 'estampa') state.step = 'get_quantidade';
        else if (field === 'qtd') {
            const qty = parseInt(fieldValue, 10);
            if (isNaN(qty) || qty <= 0) {
                sendTelegramMessage(chat_id, "Quantidade inválida. Por favor, digite um número maior que zero.");
                state.step = 'awaiting_quantity'; // Pede novamente
            } else {
                state.details.quantidade = qty;
                state.step = 'get_confirmacao';
            }
        }
       
    } else if (action === 'confirm') {
        const vendedorInfo = CONFIG.VENDEDORES.lista.find(v => String(v.telegram_id) === String(user.id));
        const vendedorNome = vendedorInfo ? vendedorInfo.nome : 'Painel Admin';

        const d = state.details;
        const produtoCompleto = `${d.modelo} ${d.genero || ''} ${d.cor} ${d.cor_manga ? 'Manga ' + d.cor_manga : ''} ${d.tamanho} ${d.estampa}`.replace(/\s+/g, ' ').trim();
        const commandText = `vendi ${d.quantidade} ${produtoCompleto}`;
       
        handleSaleCommand(chat_id, commandText, vendedorNome);
        userProperties.deleteProperty('state_' + user.id); // Limpa o estado
        return; // Sai para não enviar outra pergunta
   
    } else if (action === 'other') {
        if(value === 'qtd') {
            state.step = 'awaiting_quantity';
        }
    }

    userProperties.setProperty('state_' + user.id, JSON.stringify(state));
    askSaleQuestion(user, state);
}

function askSaleQuestion(user, state) {
    const chat_id = user.id;
    const sheetData = readSheetData(SHEET_NAMES.ESTOQUE_PRONTO);
    const headers = getHeaders(SHEET_NAMES.ESTOQUE_PRONTO);

    const getUniqueValues = (data, headers, columnName, filters = {}) => {
        const colIndex = headers.indexOf(columnName);
        const filterKeys = Object.keys(filters);
        const values = new Set();
        data.filter(row => parseInt(row[headers.indexOf('Quantidade')], 10) > 0) // Filtra apenas itens com estoque
            .forEach(row => {
            const passesFilters = filterKeys.every(key => {
                const filterColIndex = headers.indexOf(key);
                return (row[filterColIndex] || '').toString().trim() === filters[key];
            });
            if (passesFilters) {
                const value = (row[colIndex] || '').toString().trim();
                if (value) values.add(value);
            }
        });
        return Array.from(values).sort();
    };

    let text = "";
    let keyboard = { inline_keyboard: [] };
   
    switch (state.step) {
        case 'get_model':
            text = "🛒 **Venda Rápida (1/5)**\n\nSelecione o modelo:";
            const models = getUniqueValues(sheetData, headers, 'Modelo');
            keyboard.inline_keyboard = models.map(m => [{ text: m, callback_data: `venda_set_modelo_${m}` }]);
            break;
       
        case 'get_cor':
            text = `🛒 **Venda Rápida (2/5)**\n\nModelo: ${state.details.modelo}\nSelecione a cor:`;
            const colors = getUniqueValues(sheetData, headers, 'Cor', { 'Modelo': state.details.modelo });
            keyboard.inline_keyboard = colors.map(c => [{ text: c, callback_data: `venda_set_cor_${c}` }]);
            break;

        case 'get_tamanho':
            text = `🛒 **Venda Rápida (3/5)**\n\n${state.details.modelo} ${state.details.cor}\nSelecione o tamanho:`;
            const tamanhos = getUniqueValues(sheetData, headers, 'Tamanho', { 'Modelo': state.details.modelo, 'Cor': state.details.cor });
            keyboard.inline_keyboard = tamanhos.map(t => [{ text: t, callback_data: `venda_set_tamanho_${t}` }]);
            break;
       
        case 'get_estampa':
            text = `🛒 **Venda Rápida (4/5)**\n\n${state.details.modelo} ${state.details.cor} ${state.details.tamanho}\nSelecione a estampa:`;
            const estampas = getUniqueValues(sheetData, headers, 'Estampa', { 'Modelo': state.details.modelo, 'Cor': state.details.cor, 'Tamanho': state.details.tamanho });
            keyboard.inline_keyboard = estampas.map(e => [{ text: e, callback_data: `venda_set_estampa_${e}` }]);
            break;
       
        case 'get_quantidade':
            text = `🛒 **Venda Rápida (5/5)**\n\nProduto selecionado:\n${state.details.modelo} ${state.details.cor} ${state.details.tamanho} ${state.details.estampa}\n\nSelecione a quantidade:`;
            keyboard.inline_keyboard.push(
                [1, 2, 3].map(q => ({ text: q.toString(), callback_data: `venda_set_qtd_${q}` })),
                [{ text: "Outro valor...", callback_data: "venda_other_qtd" }]
            );
            break;

        case 'awaiting_quantity':
             sendTelegramMessage(chat_id, "Por favor, digite a quantidade desejada e envie.");
             return; 

        case 'get_confirmacao':
            const d = state.details;
            const produto = `${d.modelo} ${d.cor} ${d.tamanho} ${d.estampa}`;
            text = `📝 **Confirmação da Venda**\n\n` +
                   `**Produto:** ${produto}\n` +
                   `**Quantidade:** ${d.quantidade}\n\n` +
                   `Confirma o registo desta venda?`;
            keyboard.inline_keyboard.push(
                [{ text: "✅ Confirmar Venda", callback_data: "venda_confirm" }]
            );
            break;

        default:
            sendTelegramMessage(chat_id, "Ocorreu um erro, a começar de novo.");
            userProperties.deleteProperty('state_' + chat_id);
            sendMainMenu(chat_id);
            return;
    }
   
    keyboard.inline_keyboard.push([{ text: "❌ Cancelar", callback_data: "menu_principal" }]);
    sendTelegramMessage(chat_id, text, keyboard);
}


// --- LÓGICA DE NEGÓCIO (Vendas, Estoque, Produção, etc.) ---

function handleConsignmentCommand(chat_id, commandText, vendedorNome) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const lowerCommand = commandText.toLowerCase();

        if (lowerCommand.startsWith('consignar')) {
            // Comando: consignar <qtd> <produto> para <revendedor>
            const parts = lowerCommand.split(' para ');
            if (parts.length < 2) {
                sendTelegramMessage(chat_id, "Formato inválido. Use: `consignar <qtd> <produto> para <revendedor>`");
                return;
            }
            const revendedor = parts.pop().trim();
            const productInfo = parts.join(' para ').replace('consignar', '').trim();
            
            const detalhes = extractProductDetails(productInfo);
            const { modelo, cor, tamanho, estampa, cor_manga, genero, quantidade } = detalhes;

            if (!modelo || !cor || !tamanho || !estampa) {
                sendTelegramMessage(chat_id, "Não identifiquei o produto completo (modelo, cor, tamanho, estampa).");
                return;
            }
            const produtoCompleto = `${modelo} ${genero || ''} ${cor} ${cor_manga ? 'Manga ' + cor_manga : ''} ${tamanho} ${estampa}`.replace(/\s+/g, ' ').trim();

            // 1. Dar baixa no ESTOQUE_PRONTO
            const estoqueProntoSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
            const estoqueProntoData = estoqueProntoSheet.getDataRange().getValues();
            const estoqueProntoHeaders = estoqueProntoData[0];
            let produtoProntoRowIndex = -1;

            for (let i = 1; i < estoqueProntoData.length; i++) {
                const row = estoqueProntoData[i];
                const nomeProdutoNaPlanilha = `${row[estoqueProntoHeaders.indexOf('Modelo')]} ${row[estoqueProntoHeaders.indexOf('Gênero')] || ''} ${row[estoqueProntoHeaders.indexOf('Cor')]} ${row[estoqueProntoHeaders.indexOf('Cor_Manga')] ? 'Manga ' + row[estoqueProntoHeaders.indexOf('Cor_Manga')] : ''} ${row[estoqueProntoHeaders.indexOf('Tamanho')]} ${row[estoqueProntoHeaders.indexOf('Estampa')]}`.replace(/\s+/g, ' ').trim().toLowerCase();
                if (nomeProdutoNaPlanilha === produtoCompleto.toLowerCase()) {
                    const qtdDisponivel = parseInt(row[estoqueProntoHeaders.indexOf('Quantidade')], 10);
                    if (qtdDisponivel >= quantidade) {
                        produtoProntoRowIndex = i + 1;
                        estoqueProntoSheet.getRange(produtoProntoRowIndex, estoqueProntoHeaders.indexOf('Quantidade') + 1).setValue(qtdDisponivel - quantidade);
                    } else {
                        sendTelegramMessage(chat_id, `Estoque insuficiente para "${produtoCompleto.toUpperCase()}". Disponível: ${qtdDisponivel}.`);
                        return;
                    }
                    break;
                }
            }

            if (produtoProntoRowIndex === -1) {
                sendTelegramMessage(chat_id, `Produto "${produtoCompleto.toUpperCase()}" não encontrado no estoque pronto.`);
                return;
            }

            // 2. Adicionar/Atualizar no ESTOQUE_CONSIGNACAO
            const consignacaoSheet = getSheet(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
            const consignacaoData = consignacaoSheet.getDataRange().getValues();
            const consignacaoHeaders = consignacaoData[0];
            let consignacaoRowIndex = -1;

            for (let i = 1; i < consignacaoData.length; i++) {
                const row = consignacaoData[i];
                const revendedorNaPlanilha = String(row[consignacaoHeaders.indexOf('Revendedor')]).trim().toLowerCase();
                const produtoNaPlanilha = String(row[consignacaoHeaders.indexOf('Produto')]).trim().toLowerCase();
                if (revendedorNaPlanilha === revendedor.toLowerCase() && produtoNaPlanilha === produtoCompleto.toLowerCase()) {
                    consignacaoRowIndex = i + 1;
                    break;
                }
            }

            if (consignacaoRowIndex !== -1) { // Atualiza linha existente
                const qtdEnviadaCell = consignacaoSheet.getRange(consignacaoRowIndex, consignacaoHeaders.indexOf('Qtd_Enviada') + 1);
                const qtdRestanteCell = consignacaoSheet.getRange(consignacaoRowIndex, consignacaoHeaders.indexOf('Qtd_Restante') + 1);
                qtdEnviadaCell.setValue(qtdEnviadaCell.getValue() + quantidade);
                qtdRestanteCell.setValue(qtdRestanteCell.getValue() + quantidade);
            } else { // Cria nova linha
                consignacaoSheet.appendRow([revendedor, '', produtoCompleto.toUpperCase(), new Date(), quantidade, 0, 0, quantidade, 'Em consignação']);
            }
            sendTelegramMessage(chat_id, `✅ ${quantidade}x ${produtoCompleto.toUpperCase()} enviado(s) para ${revendedor}.`);

        } else if (lowerCommand.startsWith('venda c')) {
            // Comando: venda c <qtd> <produto> de <revendedor>
            const parts = lowerCommand.split(' de ');
             if (parts.length < 2) {
                sendTelegramMessage(chat_id, "Formato inválido. Use: `venda c <qtd> <produto> de <revendedor>`");
                return;
            }
            const revendedor = parts.pop().trim();
            const productInfo = parts.join(' de ').replace('venda c', '').trim();
            const detalhes = extractProductDetails(productInfo);
            const { produtoCompleto, quantidade } = getFullProductAndQty(detalhes);

            const consignacaoSheet = getSheet(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
            const consignacaoData = consignacaoSheet.getDataRange().getValues();
            const consignacaoHeaders = consignacaoData[0];
            let consignacaoRowIndex = -1;

            for (let i = 1; i < consignacaoData.length; i++) {
                const row = consignacaoData[i];
                 const revendedorNaPlanilha = String(row[consignacaoHeaders.indexOf('Revendedor')]).trim().toLowerCase();
                 const produtoNaPlanilha = String(row[consignacaoHeaders.indexOf('Produto')]).trim().toLowerCase();
                 if (revendedorNaPlanilha === revendedor.toLowerCase() && produtoNaPlanilha === produtoCompleto.toLowerCase()) {
                     const qtdRestante = parseInt(row[consignacaoHeaders.indexOf('Qtd_Restante')], 10);
                     if (qtdRestante >= quantidade) {
                         consignacaoRowIndex = i + 1;
                         const qtdVendidaCell = consignacaoSheet.getRange(i + 1, consignacaoHeaders.indexOf('Qtd_Vendida') + 1);
                         const qtdRestanteCell = consignacaoSheet.getRange(i + 1, consignacaoHeaders.indexOf('Qtd_Restante') + 1);
                         qtdVendidaCell.setValue(qtdVendidaCell.getValue() + quantidade);
                         qtdRestanteCell.setValue(qtdRestante - quantidade);
                         
                         // Logar venda
                         const precoConfig = CONFIG.PRECOS?.lista.find(p => p.modelo.toLowerCase() === detalhes.modelo.toLowerCase());
                         const precoPadrao = precoConfig?.preco_venda || 70.00;
                         let saleValue = precoPadrao * quantidade;
                         getSheet(SHEET_NAMES.VENDAS_LOG).appendRow([ `VENDA-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss')}`, new Date(), revendedor, produtoCompleto.toUpperCase(), quantidade, 'consignado', 'Telegram Bot', parseFloat(saleValue.toFixed(2)), '', 'concluída', '' ]);

                         sendTelegramMessage(chat_id, `✅ Venda consignada de ${quantidade}x ${produtoCompleto.toUpperCase()} por ${revendedor} registrada.`);
                     } else {
                         sendTelegramMessage(chat_id, `Estoque consignado insuficiente para "${produtoCompleto.toUpperCase()}" com ${revendedor}. Restam: ${qtdRestante}.`);
                     }
                     return;
                 }
            }
            sendTelegramMessage(chat_id, `Produto "${produtoCompleto.toUpperCase()}" não encontrado no estoque consignado de ${revendedor}.`);

        } else if (lowerCommand.startsWith('retorno c')) {
            // Comando: retorno c <qtd> <produto> de <revendedor>
            const parts = lowerCommand.split(' de ');
            if (parts.length < 2) {
                sendTelegramMessage(chat_id, "Formato inválido. Use: `retorno c <qtd> <produto> de <revendedor>`");
                return;
            }
            const revendedor = parts.pop().trim();
            const productInfo = parts.join(' de ').replace('retorno c', '').trim();
            const detalhes = extractProductDetails(productInfo);
            const { produtoCompleto, quantidade } = getFullProductAndQty(detalhes);

            // 1. Atualizar ESTOQUE_CONSIGNACAO
            const consignacaoSheet = getSheet(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
            const consignacaoData = consignacaoSheet.getDataRange().getValues();
            const consignacaoHeaders = consignacaoData[0];

            for (let i = 1; i < consignacaoData.length; i++) {
                 const row = consignacaoData[i];
                 const revendedorNaPlanilha = String(row[consignacaoHeaders.indexOf('Revendedor')]).trim().toLowerCase();
                 const produtoNaPlanilha = String(row[consignacaoHeaders.indexOf('Produto')]).trim().toLowerCase();
                 if (revendedorNaPlanilha === revendedor.toLowerCase() && produtoNaPlanilha === produtoCompleto.toLowerCase()) {
                     const qtdRestante = parseInt(row[consignacaoHeaders.indexOf('Qtd_Restante')], 10);
                     if (qtdRestante >= quantidade) {
                         const qtdRetornadaCell = consignacaoSheet.getRange(i + 1, consignacaoHeaders.indexOf('Qtd_Retornada') + 1);
                         const qtdRestanteCell = consignacaoSheet.getRange(i + 1, consignacaoHeaders.indexOf('Qtd_Restante') + 1);
                         qtdRetornadaCell.setValue(qtdRetornadaCell.getValue() + quantidade);
                         qtdRestanteCell.setValue(qtdRestante - quantidade);

                         // 2. Devolver ao ESTOQUE_PRONTO
                         const estoqueProntoSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
                         const estoqueProntoData = estoqueProntoSheet.getDataRange().getValues();
                         const estoqueProntoHeaders = estoqueProntoData[0];

                         for (let j = 1; j < estoqueProntoData.length; j++) {
                             const nomeProdutoPronto = `${estoqueProntoData[j][estoqueProntoHeaders.indexOf('Modelo')]} ${estoqueProntoData[j][estoqueProntoHeaders.indexOf('Gênero')] || ''} ${estoqueProntoData[j][estoqueProntoHeaders.indexOf('Cor')]} ${estoqueProntoData[j][estoqueProntoHeaders.indexOf('Cor_Manga')] ? 'Manga ' + estoqueProntoData[j][estoqueProntoHeaders.indexOf('Cor_Manga')] : ''} ${estoqueProntoData[j][estoqueProntoHeaders.indexOf('Tamanho')]} ${estoqueProntoData[j][estoqueProntoHeaders.indexOf('Estampa')]}`.replace(/\s+/g, ' ').trim().toLowerCase();
                             if (nomeProdutoPronto === produtoCompleto.toLowerCase()) {
                                 const qtdCell = estoqueProntoSheet.getRange(j + 1, estoqueProntoHeaders.indexOf('Quantidade') + 1);
                                 qtdCell.setValue(qtdCell.getValue() + quantidade);
                                 sendTelegramMessage(chat_id, `✅ Retorno de ${quantidade}x ${produtoCompleto.toUpperCase()} de ${revendedor} registrado. Estoque principal atualizado.`);
                                 return;
                             }
                         }
                     } else {
                         sendTelegramMessage(chat_id, `Quantidade de retorno inválida para "${produtoCompleto.toUpperCase()}" de ${revendedor}. Restam: ${qtdRestante}.`);
                     }
                     return;
                 }
            }
             sendTelegramMessage(chat_id, `Produto "${produtoCompleto.toUpperCase()}" não encontrado no estoque consignado de ${revendedor}.`);

        } else if (lowerCommand.startsWith('estoque c')) {
            // Comando: estoque c <revendedor>
            const revendedor = lowerCommand.replace('estoque c', '').trim();
            const consignacaoData = readSheetData(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
            const consignacaoHeaders = getHeaders(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
            let report = `📦 **Estoque Consignado de ${revendedor.toUpperCase()}:**\n`;
            let found = false;
            consignacaoData.forEach(row => {
                if (String(row[consignacaoHeaders.indexOf('Revendedor')]).trim().toLowerCase() === revendedor.toLowerCase()) {
                    const restante = parseInt(row[consignacaoHeaders.indexOf('Qtd_Restante')], 10);
                    if (restante > 0) {
                        report += `\n- ${row[consignacaoHeaders.indexOf('Produto')]}: <b>${restante} un.</b>`;
                        found = true;
                    }
                }
            });
            if (!found) {
                report = `Nenhum produto encontrado no estoque consignado de ${revendedor}.`;
            }
            sendTelegramMessage(chat_id, report);
        } else {
            sendTelegramMessage(chat_id, "Comando de consignação não reconhecido. Use o menu para ajuda.");
        }
    } catch (e) {
        Logger.log(`Erro em handleConsignmentCommand: ${e.stack}`);
        sendTelegramMessage(chat_id, `Ocorreu um erro ao processar a consignação: ${e.message}`);
    }
    finally {
        lock.releaseLock();
    }
}

// --- NOVO FLUXO DE CRIAÇÃO DE LOTE ---

function iniciarNovoLote(user) {
    const state = {
        flow: 'lote',
        step: 'get_descricao',
        details: { itens: [], solicitante: user.name }
    };
    userProperties.setProperty('state_' + user.id, JSON.stringify(state));
    setupNextAction(user.id, 'lote_descricao', "📝 **Novo Lote de Produção**\n\nPrimeiro, digite a descrição do lote (ex: 'Coleção Outono 2025').");
}

function handleLoteFlow(callback_query) {
    const user = { id: callback_query.from.id, name: callback_query.from.first_name || 'Desconhecido' };
    const chat_id = callback_query.message.chat.id;
    const message_id = callback_query.message.message_id;
    const data = callback_query.data;

    const stateJSON = userProperties.getProperty('state_' + user.id);
    if (!stateJSON) {
        editTelegramMessage(chat_id, message_id, "Ocorreu um erro ou a sua sessão expirou. Por favor, comece de novo.");
        sendMainMenu(chat_id);
        return;
    }
    let state = JSON.parse(stateJSON);

    const [flow, action, ...valueParts] = data.split('_');
    const value = valueParts.join('_');

    if (action === 'add' && value === 'item') {
        state.step = 'get_model';
        state.currentItem = {}; // Inicia um novo item temporário
    } else if (action === 'set') {
        const field = valueParts[0];
        const fieldValue = valueParts.slice(1).join('_');

        state.currentItem[field] = fieldValue;

        if (field === 'modelo') state.step = 'get_cor';
        else if (field === 'cor') state.step = 'get_tamanho';
        else if (field === 'tamanho') state.step = 'get_estampa';
        else if (field === 'estampa') state.step = 'get_quantidade';
        else if (field === 'qtd') {
            const qty = parseInt(fieldValue, 10);
            state.currentItem.quantidade = qty;

            const ci = state.currentItem;
            const produtoCompleto = `${ci.modelo} ${ci.genero || ''} ${ci.cor} ${ci.cor_manga ? 'Manga ' + ci.cor_manga : ''} ${ci.tamanho} ${ci.estampa}`.replace(/\s+/g, ' ').trim();
            state.details.itens.push({ produto: produtoCompleto, quantidade: ci.quantidade });
            
            delete state.currentItem;
            state.step = 'menu_itens';
        }
    } else if (action === 'other' && value === 'qtd') {
        state.step = 'awaiting_quantity';
        setupNextAction(user.id, 'lote_quantidade', "Digite a quantidade desejada para este item:");
    } else if (action === 'finish') {
        finalizarLote(user, state.details, message_id);
        userProperties.deleteProperty('state_' + user.id);
        return;
    } else if (action === 'cancel') {
        userProperties.deleteProperty('state_' + user.id);
        sendMainMenu(chat_id, message_id);
        return;
    }

    userProperties.setProperty('state_' + user.id, JSON.stringify(state));
    askLoteQuestion(user, state, message_id);
}

function askLoteQuestion(user, state, message_id) {
    const chat_id = user.id;
    const materiaData = readSheetData(SHEET_NAMES.ESTOQUE_MATERIA);
    const materiaHeaders = getHeaders(SHEET_NAMES.ESTOQUE_MATERIA);

    const getUniqueValues = (data, headers, columnName, filters = {}) => {
        const colIndex = headers.indexOf(columnName);
        if (colIndex === -1) return [];
        const filterKeys = Object.keys(filters);
        const values = new Set();
        data.forEach(row => {
            const passesFilters = filterKeys.every(key => {
                const filterColIndex = headers.indexOf(key);
                return (row[filterColIndex] || '').toString().trim() === filters[key];
            });
            if (passesFilters) {
                const value = (row[colIndex] || '').toString().trim();
                if (value) values.add(value);
            }
        });
        return Array.from(values).sort();
    };

    let text = "";
    let keyboard = { inline_keyboard: [] };
    let isNewMessage = message_id === null;

    switch (state.step) {
        case 'menu_itens':
            text = `📝 **Lote: ${state.details.descricao}**\n\n`;
            if (state.details.itens.length > 0) {
                text += "<b>Itens adicionados:</b>\n";
                state.details.itens.forEach(item => {
                    text += `- ${item.quantidade}x ${item.produto}\n`;
                });
            } else {
                text += "Nenhum item adicionado ainda.\n";
            }
            text += "\nO que deseja fazer?";
            keyboard.inline_keyboard.push([{ text: "➕ Adicionar Item", callback_data: "lote_add_item" }]);
            if (state.details.itens.length > 0) {
                keyboard.inline_keyboard.push([{ text: "✅ Finalizar e Criar Lote", callback_data: "lote_finish" }]);
            }
            break;
        
        case 'get_model':
            text = "➕ **Adicionar Item (1/5)**\n\nSelecione o modelo da peça lisa:";
            const models = getUniqueValues(materiaData, materiaHeaders, 'Modelo');
            keyboard.inline_keyboard = models.map(m => [{ text: m, callback_data: `lote_set_modelo_${m}` }]);
            break;

        case 'get_cor':
            text = `➕ **Adicionar Item (2/5)**\n\nModelo: ${state.currentItem.modelo}\nSelecione a cor:`;
            const colors = getUniqueValues(materiaData, materiaHeaders, 'Cor', { 'Modelo': state.currentItem.modelo });
            keyboard.inline_keyboard = colors.map(c => [{ text: c, callback_data: `lote_set_cor_${c}` }]);
            break;

        case 'get_tamanho':
            text = `➕ **Adicionar Item (3/5)**\n\n${state.currentItem.modelo} ${state.currentItem.cor}\nSelecione o tamanho:`;
            const tamanhos = getUniqueValues(materiaData, materiaHeaders, 'Tamanho', { 'Modelo': state.currentItem.modelo, 'Cor': state.currentItem.cor });
            keyboard.inline_keyboard = tamanhos.map(t => [{ text: t, callback_data: `lote_set_tamanho_${t}` }]);
            break;
       
        case 'get_estampa':
            text = `➕ **Adicionar Item (4/5)**\n\n${state.currentItem.modelo} ${state.currentItem.cor} ${state.currentItem.tamanho}\nDigite o nome da estampa:`;
            setupNextAction(user.id, 'lote_estampa', text);
            return; // Sai da função para aguardar o input do utilizador

        case 'get_quantidade':
            const ci = state.currentItem;
            text = `➕ **Adicionar Item (5/5)**\n\nItem: ${ci.modelo} ${ci.cor} ${ci.tamanho} ${ci.estampa}\n\nSelecione a quantidade:`;
            keyboard.inline_keyboard.push(
                [1, 2, 3, 5, 10].map(q => ({ text: q.toString(), callback_data: `lote_set_qtd_${q}` })),
                [{ text: "Outro valor...", callback_data: "lote_other_qtd" }]
            );
            break;
    }

    keyboard.inline_keyboard.push([{ text: "❌ Cancelar Lote", callback_data: "lote_cancel" }]);
    
    if (isNewMessage) {
        sendTelegramMessage(chat_id, text, keyboard);
    } else {
        editTelegramMessage(chat_id, message_id, text, keyboard);
    }
}


function finalizarLote(user, loteDetails, message_id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 

  try {
    const lote = loteDetails;
    if (!lote || lote.itens.length === 0) {
        sendTelegramMessage(user.id, "Seu lote está vazio. Adicione itens antes de finalizar.");
        return;
    }
   
    const materiaPrimaSheet = getSheet(SHEET_NAMES.ESTOQUE_MATERIA);
    const materiaHeaders = materiaPrimaSheet.getRange(1, 1, 1, materiaPrimaSheet.getLastColumn()).getValues()[0];
    const materiaData = materiaPrimaSheet.getDataRange().getValues();

    let reservas = [];
    let pickingList = "";
    let erroEstoque = "";
    let custoTotalDoLote = 0; 

    for (const item of lote.itens) {
      const detalhesItem = extractProductDetails(item.produto); 
     
      let materiaPrimaNecessaria = `${detalhesItem.modelo}`;
      if (detalhesItem.genero) materiaPrimaNecessaria += ` ${detalhesItem.genero}`;
      materiaPrimaNecessaria += ` ${detalhesItem.cor}`;
      if (detalhesItem.cor_manga) materiaPrimaNecessaria += ` Manga ${detalhesItem.cor_manga}`;
      materiaPrimaNecessaria += ` ${detalhesItem.tamanho}`;
      materiaPrimaNecessaria = materiaPrimaNecessaria.trim().toLowerCase();

      let materiaEncontrada = false;
      for (let i = 1; i < materiaData.length; i++) { 
        const row = materiaData[i];
        const rowModel = String(row[materiaHeaders.indexOf('Modelo')]).toLowerCase();
        const rowCor = String(row[materiaHeaders.indexOf('Cor')]).toLowerCase();
        const rowTamanho = String(row[materiaHeaders.indexOf('Tamanho')]).toLowerCase();
        const rowCorManga = String(row[materiaHeaders.indexOf('Cor_Manga')] || '').toLowerCase();
        const rowGenero = String(row[materiaHeaders.indexOf('Gênero')] || '').toLowerCase();
       
        let nomeMateriaCompleto = `${rowModel}`;
        if (rowGenero) nomeMateriaCompleto += ` ${rowGenero}`;
        nomeMateriaCompleto += ` ${rowCor}`;
        if (rowCorManga) nomeMateriaCompleto += ` Manga ${rowCorManga}`;
        nomeMateriaCompleto += ` ${rowTamanho}`;
        nomeMateriaCompleto = nomeMateriaCompleto.trim();
       
        if (nomeMateriaCompleto === materiaPrimaNecessaria) {
          materiaEncontrada = true;
          const qtdAtual = parseInt(row[materiaHeaders.indexOf('Qtd_Atual')], 10);
          const qtdReservada = parseInt(row[materiaHeaders.indexOf('Qtd_Reservada')] || 0, 10);
          const disponivel = qtdAtual - qtdReservada;

          if (disponivel >= item.quantidade) {
            reservas.push({ rowIndex: i + 1, qtd: item.quantidade });
            pickingList += ` • ${item.quantidade}x ${nomeMateriaCompleto.toUpperCase()}\n`;
           
            const custoUnitarioMateria = parseFloat(String(row[materiaHeaders.indexOf('Custo')] || '0').replace(',', '.')) || 0;
            if (custoUnitarioMateria > 0) {
              custoTotalDoLote += custoUnitarioMateria * item.quantidade;
            } else {
              erroEstoque = `Custo da matéria-prima "${nomeMateriaCompleto.toUpperCase()}" não definido na planilha ESTOQUE_MATERIA.`;
            }

          } else {
            erroEstoque = `Estoque insuficiente para "${nomeMateriaCompleto.toUpperCase()}". Necessário: ${item.quantidade}, Disponível: ${disponivel}.`;
          }
          break;
        }
      }
      if (!materiaEncontrada) {
        erroEstoque = `Matéria-prima "${materiaPrimaNecessaria.toUpperCase()}" não encontrada no estoque.`;
      }
      if (erroEstoque) break;
    }

    if (erroEstoque) {
      editTelegramMessage(user.id, message_id, `🔴 ERRO: Não foi possível criar o lote. ${erroEstoque}`);
      return;
    }

    reservas.forEach(res => {
      const qtdReservadaIndex = materiaHeaders.indexOf('Qtd_Reservada') + 1;
      const cell = materiaPrimaSheet.getRange(res.rowIndex, qtdReservadaIndex);
      const valorAtual = cell.getValue() || 0;
      cell.setValue(valorAtual + res.qtd);
    });

    const novoIdLote = `LOTE-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd-HHmmss')}`;
    const lotesSheet = getSheet(SHEET_NAMES.LOTES_PRODUCAO);
    lotesSheet.appendRow([novoIdLote, lote.descricao, lote.solicitante, new Date(), 'Aguardando Produção', '', custoTotalDoLote.toFixed(2)]);

    const itensLoteSheet = getSheet(SHEET_NAMES.ITENS_LOTE);
    lote.itens.forEach(item => {
      itensLoteSheet.appendRow([novoIdLote, item.produto, item.quantidade]);
    });
   
    const grupoProducaoId = CONFIG.CHAT_IDS.GRUPO_ALERTAS; 
    editTelegramMessage(user.id, message_id, `✅ Lote "${lote.descricao}" (ID: ${novoIdLote}) criado com sucesso! Notificação enviada para a produção.`);
    sendTelegramMessage(grupoProducaoId, 
      `📢 **Novo Lote de Produção**\n\n` +
      `**ID:** <code>${novoIdLote}</code>\n` +
      `**Descrição:** ${lote.descricao}\n` +
      `**Solicitante:** ${lote.solicitante}\n\n` +
      `📋 **Lista para Coleta (Picking List):**\n${pickingList}`
    );
   
    userProperties.deleteProperty('state_' + user.id);

  } finally {
    lock.releaseLock();
  }
}

function concluirLoteDeProducao(command, user) {
  const loteId = command.toLowerCase().replace('concluir lote', '').trim().toUpperCase();
  if (!loteId) {
    sendTelegramMessage(user.id, "Especifique o ID do lote. Ex: <code>concluir lote LOTE-12345</code>");
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const lotesSheet = getSheet(SHEET_NAMES.LOTES_PRODUCAO);
    const lotesData = lotesSheet.getDataRange().getValues();
    const lotesHeaders = lotesData[0];
    let loteRowIndex = -1, loteRowData = null;

    for (let i = 1; i < lotesData.length; i++) {
      if (lotesData[i][lotesHeaders.indexOf('ID_Lote')] === loteId) {
        loteRowData = lotesData[i];
        if (String(loteRowData[lotesHeaders.indexOf('Status')]).trim() === 'Aguardando Produção') loteRowIndex = i + 1;
        break;
      }
    }

    if (!loteRowData) { sendTelegramMessage(user.id, `Lote "${loteId}" não encontrado.`); return; }
    if (loteRowIndex === -1) { sendTelegramMessage(user.id, `Lote "${loteId}" já foi concluído ou cancelado.`); return; }

    const itensDoLote = readSheetData(SHEET_NAMES.ITENS_LOTE).filter(row => row[0] === loteId);
   
    const custoTotalDoLote = parseFloat(loteRowData[lotesHeaders.indexOf('Custo_Total')] || 0);
    const totalItensProduzidos = itensDoLote.reduce((sum, item) => sum + parseInt(item[2], 10), 0);
    const custoUnitarioMedio = totalItensProduzidos > 0 ? (custoTotalDoLote / totalItensProduzidos) : 0;

    const materiaPrimaSheet = getSheet(SHEET_NAMES.ESTOQUE_MATERIA);
    const materiaData = materiaPrimaSheet.getDataRange().getValues();
    const materiaHeaders = materiaData[0];

    const estoqueProntoSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
    const estoqueProntoData = estoqueProntoSheet.getDataRange().getValues();
    const estoqueProntoHeaders = estoqueProntoData[0];

    for (const item of itensDoLote) {
      const [ , produtoFinalCompleto, quantidadeProduzidaStr ] = item;
      const quantidadeProduzida = parseInt(quantidadeProduzidaStr, 10);
      const detalhesItem = extractProductDetails(produtoFinalCompleto);
      const materiaPrimaNecessaria = `${detalhesItem.modelo || ''} ${detalhesItem.genero || ''} ${detalhesItem.cor || ''} ${detalhesItem.cor_manga ? 'Manga ' + detalhesItem.cor_manga : ''} ${detalhesItem.tamanho || ''}`.replace(/\s+/g, ' ').trim().toLowerCase();
     
      for (let i = 1; i < materiaData.length; i++) {
        let nomeMateriaCompleto = `${String(materiaData[i][materiaHeaders.indexOf('Modelo')] || '')} ${String(materiaData[i][materiaHeaders.indexOf('Gênero')] || '')} ${String(materiaData[i][materiaHeaders.indexOf('Cor')] || '')} ${materiaData[i][materiaHeaders.indexOf('Cor_Manga')] ? 'Manga ' + materiaData[i][materiaHeaders.indexOf('Cor_Manga')] : ''} ${String(materiaData[i][materiaHeaders.indexOf('Tamanho')] || '')}`.replace(/\s+/g, ' ').trim().toLowerCase();
        if (nomeMateriaCompleto === materiaPrimaNecessaria) {
          const qtdAtualCell = materiaPrimaSheet.getRange(i + 1, materiaHeaders.indexOf('Qtd_Atual') + 1);
          const qtdReservadaCell = materiaPrimaSheet.getRange(i + 1, materiaHeaders.indexOf('Qtd_Reservada') + 1);
          qtdAtualCell.setValue(qtdAtualCell.getValue() - quantidadeProduzida);
          qtdReservadaCell.setValue(qtdReservadaCell.getValue() - quantidadeProduzida);
          break;
        }
      }

      let produtoFinalEncontrado = false;
      for (let j = 1; j < estoqueProntoData.length; j++) {
        let nomeProdutoPronto = `${String(estoqueProntoData[j][estoqueProntoHeaders.indexOf('Modelo')] || '')} ${String(estoqueProntoData[j][estoqueProntoHeaders.indexOf('Gênero')] || '')} ${String(estoqueProntoData[j][estoqueProntoHeaders.indexOf('Cor')] || '')} ${String(estoqueProntoData[j][estoqueProntoHeaders.indexOf('Cor_Manga')] || '') ? 'Manga ' + estoqueProntoData[j][estoqueProntoHeaders.indexOf('Cor_Manga')] : ''} ${String(estoqueProntoData[j][estoqueProntoHeaders.indexOf('Tamanho')] || '')} ${String(estoqueProntoData[j][estoqueProntoHeaders.indexOf('Estampa')] || '')}`.replace(/\s+/g, ' ').trim().toLowerCase();
        if (nomeProdutoPronto === produtoFinalCompleto.toLowerCase()) {
          produtoFinalEncontrado = true;
          const qtdCell = estoqueProntoSheet.getRange(j + 1, estoqueProntoHeaders.indexOf('Quantidade') + 1);
          qtdCell.setValue(qtdCell.getValue() + quantidadeProduzida);
          estoqueProntoSheet.getRange(j + 1, estoqueProntoHeaders.indexOf('Custo_Unitario') + 1).setValue(custoUnitarioMedio.toFixed(2));
          estoqueProntoSheet.getRange(j + 1, estoqueProntoHeaders.indexOf('Data_Atualização') + 1).setValue(new Date());
          estoqueProntoSheet.getRange(j + 1, estoqueProntoHeaders.indexOf('Status') + 1).setValue('disponivel');
          break;
        }
      }
      if (!produtoFinalEncontrado) {
        const precoConfig = CONFIG.PRECOS?.lista.find(p => p.modelo.toLowerCase() === detalhesItem.modelo.toLowerCase());
        const precoPadrao = precoConfig?.preco_venda || 70.00;
        const novoId = `${(detalhesItem.modelo || 'P').substring(0,2)}-${(detalhesItem.estampa || 'E').replace(/\s/g, '').substring(0,3)}-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'mmss')}`.toUpperCase();
        estoqueProntoSheet.appendRow([novoId, detalhesItem.modelo, detalhesItem.cor, detalhesItem.cor_manga || '', detalhesItem.tamanho, detalhesItem.genero || '', detalhesItem.estampa, quantidadeProduzida, precoPadrao, custoUnitarioMedio.toFixed(2), new Date(), 'disponivel']);
      }
    }
   
    lotesSheet.getRange(loteRowIndex, lotesHeaders.indexOf('Status') + 1).setValue('Concluído');
    lotesSheet.getRange(loteRowIndex, lotesHeaders.indexOf('Data_Conclusao') + 1).setValue(new Date());

    sendTelegramMessage(user.id, `✅ Lote ${loteId} concluído com sucesso e estoque atualizado!`);
    sendTelegramMessage(CONFIG.CHAT_IDS.GRUPO_ALERTAS, `🎉 Lote ${loteId} foi marcado como CONCLUÍDO por ${user.name}.`);

  } finally {
    lock.releaseLock();
  }
}

/**
 * NOVO: Função para cancelar um lote de produção.
 */
function cancelarLoteDeProducao(command, user) {
  const loteId = command.toLowerCase().replace('cancelar lote', '').trim().toUpperCase();
  if (!loteId) {
    sendTelegramMessage(user.id, "Especifique o ID do lote. Ex: <code>cancelar lote LOTE-12345</code>");
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const lotesSheet = getSheet(SHEET_NAMES.LOTES_PRODUCAO);
    const lotesData = lotesSheet.getDataRange().getValues();
    const lotesHeaders = lotesData[0];
    let loteRowIndex = -1;

    for (let i = 1; i < lotesData.length; i++) {
      if (lotesData[i][lotesHeaders.indexOf('ID_Lote')] === loteId) {
        if (String(lotesData[i][lotesHeaders.indexOf('Status')]).trim() === 'Aguardando Produção') {
          loteRowIndex = i + 1;
        } else {
          sendTelegramMessage(user.id, `Não é possível cancelar o lote "${loteId}", pois o seu status é "${lotesData[i][lotesHeaders.indexOf('Status')]}".`);
          return;
        }
        break;
      }
    }

    if (loteRowIndex === -1) {
      sendTelegramMessage(user.id, `Lote "${loteId}" não encontrado ou já processado.`);
      return;
    }

    const itensDoLote = readSheetData(SHEET_NAMES.ITENS_LOTE).filter(row => row[0] === loteId);
    
    if (itensDoLote.length === 0) {
      // Se não há itens, apenas cancela o lote
      lotesSheet.getRange(loteRowIndex, lotesHeaders.indexOf('Status') + 1).setValue('Cancelado');
      sendTelegramMessage(user.id, `✅ Lote ${loteId} (sem itens) cancelado com sucesso.`);
      return;
    }

    const materiaPrimaSheet = getSheet(SHEET_NAMES.ESTOQUE_MATERIA);
    const materiaData = materiaPrimaSheet.getDataRange().getValues();
    const materiaHeaders = materiaData[0];

    for (const item of itensDoLote) {
      const [ , produtoFinalCompleto, quantidadeReservadaStr ] = item;
      const quantidadeReservada = parseInt(quantidadeReservadaStr, 10);
      const detalhesItem = extractProductDetails(produtoFinalCompleto);
      const materiaPrimaNecessaria = `${detalhesItem.modelo || ''} ${detalhesItem.genero || ''} ${detalhesItem.cor || ''} ${detalhesItem.cor_manga ? 'Manga ' + detalhesItem.cor_manga : ''} ${detalhesItem.tamanho || ''}`.replace(/\s+/g, ' ').trim().toLowerCase();
     
      let materiaEncontrada = false;
      for (let i = 1; i < materiaData.length; i++) {
        let nomeMateriaCompleto = `${String(materiaData[i][materiaHeaders.indexOf('Modelo')] || '')} ${String(materiaData[i][materiaHeaders.indexOf('Gênero')] || '')} ${String(materiaData[i][materiaHeaders.indexOf('Cor')] || '')} ${materiaData[i][materiaHeaders.indexOf('Cor_Manga')] ? 'Manga ' + materiaData[i][materiaHeaders.indexOf('Cor_Manga')] : ''} ${String(materiaData[i][materiaHeaders.indexOf('Tamanho')] || '')}`.replace(/\s+/g, ' ').trim().toLowerCase();
        if (nomeMateriaCompleto === materiaPrimaNecessaria) {
          materiaEncontrada = true;
          const qtdReservadaCell = materiaPrimaSheet.getRange(i + 1, materiaHeaders.indexOf('Qtd_Reservada') + 1);
          const valorAtual = qtdReservadaCell.getValue() || 0;
          qtdReservadaCell.setValue(Math.max(0, valorAtual - quantidadeReservada));
          break;
        }
      }
      if (!materiaEncontrada) {
         Logger.log(`AVISO: Matéria-prima "${materiaPrimaNecessaria}" do lote cancelado não foi encontrada para devolver ao estoque.`);
      }
    }
   
    lotesSheet.getRange(loteRowIndex, lotesHeaders.indexOf('Status') + 1).setValue('Cancelado');

    sendTelegramMessage(user.id, `✅ Lote ${loteId} cancelado! A matéria-prima que estava reservada foi devolvida ao estoque.`);
    sendTelegramMessage(CONFIG.CHAT_IDS.GRUPO_ALERTAS, `❌ O Lote ${loteId} foi CANCELADO por ${user.name}.`);

  } finally {
    lock.releaseLock();
  }
}


function extractProductDetails(text) {
    const lowerText = text.toLowerCase();
    const details = { modelo: null, cor: null, tamanho: null, estampa: null, tipo_tecido: null, cor_manga: null, genero: null, quantidade: 1, forma_pagamento: null };
    let tempText = lowerText;
    const qtyMatch = tempText.match(/(\b\d+\s*)/);
    if (qtyMatch) {
        details.quantidade = parseInt(qtyMatch[0], 10);
        tempText = tempText.replace(qtyMatch[0], '').trim();
    }
    let cleanText = tempText.replace(/vendi|tem|estoque|addprod|relatorio|entrada|saida|consignar|para|venda consignada|por|retorno consignado|de|estoque consignado/g, ' ').trim();
   
    const orderedCategories = ['tamanho', 'genero', 'cor_manga', 'cor', 'modelo', 'tipo_tecido', 'forma_pagamento'];
   
    for (const categoria of orderedCategories) {
        if (!KEYWORDS[categoria]) continue;
       
        let bestMatch = KEYWORDS[categoria]
            .filter(kw => new RegExp(`\\b${kw.palavraUsuario}\\b`, 'i').test(cleanText))
            .sort((a, b) => (b.prioridade - a.prioridade) || (b.palavraUsuario.length - a.palavraUsuario.length))[0];
       
        if (bestMatch) {
            details[categoria] = bestMatch.palavraSistema;
            cleanText = cleanText.replace(new RegExp(`\\b${bestMatch.palavraUsuario}\\b`, 'i'), ' ').trim();
        }
    }

    if (cleanText) {
      details.estampa = cleanText.split(' ').filter(Boolean).map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
    }
   
    return details;
}


function calculateWholesaleDiscount(quantity, paymentMethod) {
  if (!CONFIG.DESCONTOS_ATACADO || !CONFIG.DESCONTOS_ATACADO.lista) return 0;
  let applicableDiscount = 0;
  const sortedRules = CONFIG.DESCONTOS_ATACADO.lista.sort((a, b) => b.faixa_min - a.faixa_min);

  for (const rule of sortedRules) {
    if (quantity >= rule.faixa_min && (rule.faixa_max === '*' || quantity <= rule.faixa_max)) {
      if (rule.condicao_pagamento === '*' || (paymentMethod && rule.condicao_pagamento === paymentMethod.toUpperCase())) {
        const discountStr = String(rule.desconto_percentual);
        if (discountStr.endsWith('%')) {
          applicableDiscount = parseFloat(discountStr) / 100;
        } else {
          applicableDiscount = parseFloat(discountStr);
        }
        break; 
      }
    }
  }
  return applicableDiscount;
}
function getFullProductAndQty(detalhes) {
    const { modelo, cor, tamanho, estampa, cor_manga, genero, quantidade } = detalhes;
    if (!modelo || !cor || !tamanho || !estampa) {
        throw new Error("Não identifiquei o produto completo (modelo, cor, tamanho, estampa).");
    }
    const produtoCompleto = `${modelo} ${genero || ''} ${cor} ${cor_manga ? 'Manga ' + cor_manga : ''} ${tamanho} ${estampa}`.replace(/\s+/g, ' ').trim();
    return { produtoCompleto, quantidade };
}

function handleSaleCommand(chat_id, commandText, vendedorNome) {
    const { modelo, cor, tamanho, estampa, cor_manga, genero, quantidade, forma_pagamento } = extractProductDetails(commandText);
    if (!modelo || !cor || !tamanho || !estampa) {
        sendTelegramMessage(chat_id, "Não identifiquei o produto completo (modelo, cor, tamanho, estampa).");
        return;
    }
    const estoqueSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
    const estoqueData = estoqueSheet.getDataRange().getValues();
    const headers = estoqueData[0];
    const productNameToFind = `${modelo || ''} ${genero || ''} ${cor || ''} ${cor_manga ? 'Manga ' + cor_manga : ''} ${tamanho || ''} ${estampa || ''}`.replace(/\s+/g, ' ').trim().toLowerCase();

    for (let i = 1; i < estoqueData.length; i++) {
        const row = estoqueData[i];
        const productNameInSheet = `${String(row[headers.indexOf('Modelo')] || '')} ${String(row[headers.indexOf('Gênero')] || '')} ${String(row[headers.indexOf('Cor')] || '')} ${row[headers.indexOf('Cor_Manga')] ? 'Manga ' + row[headers.indexOf('Cor_Manga')] : ''} ${String(row[headers.indexOf('Tamanho')] || '')} ${String(row[headers.indexOf('Estampa')] || '')}`.replace(/\s+/g, ' ').trim().toLowerCase();
       
        if (productNameInSheet === productNameToFind) {
            const qtyCell = estoqueSheet.getRange(i + 1, headers.indexOf('Quantidade') + 1);
            const currentQty = parseInt(qtyCell.getValue(), 10);
            if (currentQty >= quantidade) {
                const newQty = currentQty - quantidade;
                qtyCell.setValue(newQty);
                estoqueSheet.getRange(i + 1, headers.indexOf('Data_Atualização') + 1).setValue(new Date());
                estoqueSheet.getRange(i + 1, headers.indexOf('Status') + 1).setValue(newQty > 0 ? 'disponivel' : 'indisponivel');
               
                const custoUnitario = parseFloat(String(row[headers.indexOf('Custo_Unitario')] || '0').replace(',', '.')) || 0;
                const custoTotalVenda = custoUnitario * quantidade;
                let saleValue = parseFloat(String(row[headers.indexOf('Preço')] || '0').replace(',', '.')) * quantidade;
               
                const discount = calculateWholesaleDiscount(quantidade, forma_pagamento);
                const tipoVenda = discount > 0 ? 'atacado' : 'varejo';
                if (discount > 0) saleValue *= (1 - discount);

                const lucro = saleValue - custoTotalVenda;

                const vendasLogSheet = getSheet(SHEET_NAMES.VENDAS_LOG);
                const logHeaders = getHeaders(SHEET_NAMES.VENDAS_LOG);
                const newLogRow = new Array(logHeaders.length).fill('');

                logHeaders.forEach((header, index) => {
                    const trimmedHeader = String(header).trim();
                    switch(trimmedHeader) {
                        case 'ID': newLogRow[index] = `VENDA-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss')}`; break;
                        case 'Data_Hora': newLogRow[index] = new Date(); break;
                        case 'Vendedor': newLogRow[index] = vendedorNome; break;
                        case 'Produto_Completo': newLogRow[index] = productNameToFind.toUpperCase(); break;
                        case 'Quantidade': newLogRow[index] = quantidade; break;
                        case 'Tipo_Venda': newLogRow[index] = tipoVenda; break;
                        case 'Canal': newLogRow[index] = 'Telegram Bot'; break;
                        case 'Valor': newLogRow[index] = parseFloat(saleValue.toFixed(2)); break;
                        case 'Custo_Total': newLogRow[index] = parseFloat(custoTotalVenda.toFixed(2)); break;
                        case 'Lucro': newLogRow[index] = parseFloat(lucro.toFixed(2)); break;
                        case 'Status': newLogRow[index] = 'concluída'; break;
                        case 'Forma_Pagamento': newLogRow[index] = forma_pagamento || ''; break;
                    }
                });
               
                vendasLogSheet.appendRow(newLogRow);
               
                sendTelegramMessage(chat_id, `✅ Venda registrada! ${productNameToFind.toUpperCase()}. Restam ${newQty} unidades.`);
            } else {
                sendTelegramMessage(chat_id, `Ops! Estoque insuficiente para ${productNameToFind.toUpperCase()}. Temos apenas ${currentQty}.`);
            }
            return;
        }
    }
    sendTelegramMessage(chat_id, `Produto "${productNameToFind.toUpperCase()}" não encontrado no estoque.`);
}

function handleStockQueryCommand(chat_id, commandText) {
    const lowerCommand = commandText.toLowerCase().replace('tem', '').trim();
    const searchTerms = lowerCommand.split(' ').filter(Boolean);

    let response = "📦 **Consulta de Estoque:**\n";
    let foundMateria = false;
    let foundPronto = false;

    // Estoque de Matéria Prima
    const materiaData = readSheetData(SHEET_NAMES.ESTOQUE_MATERIA);
    const materiaHeaders = getHeaders(SHEET_NAMES.ESTOQUE_MATERIA);
    let materiaResults = "";
    materiaData.forEach(row => {
        const nomeMateria = `${String(row[materiaHeaders.indexOf('Modelo')] || '')} ${String(row[materiaHeaders.indexOf('Gênero')] || '')} ${String(row[materiaHeaders.indexOf('Cor')] || '')} ${row[materiaHeaders.indexOf('Cor_Manga')] ? 'Manga ' + row[materiaHeaders.indexOf('Cor_Manga')] : ''} ${String(row[materiaHeaders.indexOf('Tamanho')] || '')}`.replace(/\s+/g, ' ').trim().toLowerCase();
       
        const allTermsMatch = searchTerms.every(term => nomeMateria.includes(term));

        if (allTermsMatch) {
            const disponivel = parseInt(row[materiaHeaders.indexOf('Qtd_Atual')]) - (parseInt(row[materiaHeaders.indexOf('Qtd_Reservada')] || 0));
            if (disponivel > 0) {
              materiaResults += `\n- ${nomeMateria.toUpperCase()}: ${disponivel} (disponível)`;
              foundMateria = true;
            }
        }
    });
    if(foundMateria) response += "\n**Matéria-Prima (Peças Lisas):**" + materiaResults;

    // Estoque Pronto
    const prontoData = readSheetData(SHEET_NAMES.ESTOQUE_PRONTO);
    const prontoHeaders = getHeaders(SHEET_NAMES.ESTOQUE_PRONTO);
    let prontoResults = "";
    prontoData.forEach(row => {
        const nomeProduto = `${String(row[prontoHeaders.indexOf('Modelo')] || '')} ${String(row[prontoHeaders.indexOf('Gênero')] || '')} ${String(row[prontoHeaders.indexOf('Cor')] || '')} ${row[prontoHeaders.indexOf('Cor_Manga')] ? 'Manga ' + row[prontoHeaders.indexOf('Cor_Manga')] : ''} ${String(row[prontoHeaders.indexOf('Tamanho')] || '')} ${String(row[prontoHeaders.indexOf('Estampa')] || '')}`.replace(/\s+/g, ' ').trim().toLowerCase();
       
        const allTermsMatch = searchTerms.every(term => nomeProduto.includes(term));

        if (allTermsMatch) {
            const quantidade = parseInt(row[prontoHeaders.indexOf('Quantidade')], 10) || 0;
            if (quantidade > 0) {
              prontoResults += `\n- ${nomeProduto.toUpperCase()}: ${quantidade} unidades`;
              foundPronto = true;
            }
        }
    });
    if(foundPronto) response += "\n\n**Produtos Prontos:**" + prontoResults;

    if (!foundMateria && !foundPronto) {
      response = "Nenhum item encontrado no estoque com os termos da sua pesquisa.";
    }

    sendTelegramMessage(chat_id, response);
}


function handleCashFlowCommand(chat_id, commandText, vendedorNome) {
    const parts = commandText.trim().split(/\s+/);
    const type = parts[0].toLowerCase();
    const value = parseFloat(parts[1]?.replace(',', '.'));
    const description = parts.slice(2).join(' ');

    if ((type !== 'entrada' && type !== 'saida') || isNaN(value) || !description) {
        sendTelegramMessage(chat_id, "Formato inválido. Use: `entrada/saida <valor> <descrição>`");
        return;
    }

    const fluxoSheet = getSheet(SHEET_NAMES.FLUXO_CAIXA);
    fluxoSheet.appendRow([ new Date(), type === 'entrada' ? 'Entrada' : 'Saída', description, type === 'entrada' ? value : '', type === 'saida' ? value : '', '', '', '', vendedorNome ]);
    sendTelegramMessage(chat_id, `✅ ${type.charAt(0).toUpperCase() + type.slice(1)} de R$ ${value.toFixed(2)} (${description}) registrada.`);
}

function handleMateriaPrimaCommand(chat_id, commandText, vendedorNome) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const commandContent = commandText.toLowerCase().replace('add materia', '').trim();
    const parts = commandContent.split(/\s+/);
    const quantidade = parseInt(parts[0], 10);
   
    if (isNaN(quantidade) || quantidade <= 0) {
      sendTelegramMessage(chat_id, "Formato inválido. Use: `add materia <quantidade> <descrição da peça lisa>`");
      return;
    }
   
    const productText = parts.slice(1).join(' ');
    const detalhes = extractProductDetails(productText);
    const { modelo, cor, tamanho, cor_manga, genero } = detalhes;

    if (!modelo || !cor || !tamanho) {
      sendTelegramMessage(chat_id, "Não identifiquei a matéria-prima completa. Forneça pelo menos modelo, cor e tamanho.");
      return;
    }

    const materiaSheet = getSheet(SHEET_NAMES.ESTOQUE_MATERIA);
    const headers = getHeaders(SHEET_NAMES.ESTOQUE_MATERIA);
    const data = readSheetData(SHEET_NAMES.ESTOQUE_MATERIA);

    const inputModelo = (modelo || '').trim().toLowerCase();
    const inputGenero = (genero || '').trim().toLowerCase();
    const inputCor = (cor || '').trim().toLowerCase();
    const inputCorManga = (cor_manga || '').trim().toLowerCase();
    const inputTamanho = (tamanho || '').trim().toLowerCase();

    let matches = data
        .map((row, index) => ({ row, index }))
        .filter(({ row }) => {
            const rowModelo = (row[headers.indexOf('Modelo')] || '').trim().toLowerCase();
            const rowCor = (row[headers.indexOf('Cor')] || '').trim().toLowerCase();
            const rowTamanho = (row[headers.indexOf('Tamanho')] || '').trim().toLowerCase();

            if (rowModelo !== inputModelo || rowCor !== inputCor || rowTamanho !== inputTamanho) {
                return false;
            }
            if (inputGenero) {
                const rowGenero = (row[headers.indexOf('Gênero')] || '').trim().toLowerCase();
                if (rowGenero !== inputGenero) return false;
            }
            if (inputCorManga) {
                const rowCorManga = (row[headers.indexOf('Cor_Manga')] || '').trim().toLowerCase();
                if (rowCorManga !== inputCorManga) return false;
            }
            return true;
        });

    if (matches.length === 0 && !inputGenero && !inputCorManga) {
        const baseMatches = data
            .map((row, index) => ({ row, index }))
            .filter(({ row }) => {
                const rowModelo = (row[headers.indexOf('Modelo')] || '').trim().toLowerCase();
                const rowCor = (row[headers.indexOf('Cor')] || '').trim().toLowerCase();
                const rowTamanho = (row[headers.indexOf('Tamanho')] || '').trim().toLowerCase();
                return rowModelo === inputModelo && rowCor === inputCor && rowTamanho === inputTamanho;
            });
        if (baseMatches.length === 1) {
            matches = baseMatches;
        }
    }


    if (matches.length === 1) {
        const rowIndex = matches[0].index;
        const qtdAtualColIndex = headers.indexOf('Qtd_Atual');
        const qtdAtual = parseInt(data[rowIndex][qtdAtualColIndex]) || 0;
        const novaQtd = qtdAtual + quantidade;
        const rowData = matches[0].row;
        const nomeCompleto = `${rowData[headers.indexOf('Modelo')]} ${rowData[headers.indexOf('Gênero')] || ''} ${rowData[headers.indexOf('Cor')]} ${rowData[headers.indexOf('Cor_Manga')] ? 'Manga '+rowData[headers.indexOf('Cor_Manga')] : ''} ${rowData[headers.indexOf('Tamanho')]}`.replace(/\s+/g, ' ').trim();
       
        materiaSheet.getRange(rowIndex + 2, qtdAtualColIndex + 1).setValue(novaQtd);
        sendTelegramMessage(chat_id, `✅ Stock de "${nomeCompleto.toUpperCase()}" atualizado. Novo total: <b>${novaQtd}</b>.`);
   
    } else if (matches.length > 1) {
        sendTelegramMessage(chat_id, 'Ambiguidade: Mais de um item corresponde à sua descrição. Por favor, forneça mais detalhes (como gênero ou cor da manga).');
   
    } else {
        const novaLinha = new Array(headers.length).fill('');
        novaLinha[headers.indexOf('Modelo')] = modelo;
        novaLinha[headers.indexOf('Cor')] = cor;
        novaLinha[headers.indexOf('Tamanho')] = tamanho;
        novaLinha[headers.indexOf('Qtd_Atual')] = quantidade;
        novaLinha[headers.indexOf('Qtd_Reservada')] = 0;
        novaLinha[headers.indexOf('Data_Entrada')] = new Date();
        novaLinha[headers.indexOf('Fornecedor')] = `Telegram (${vendedorNome})`;
        novaLinha[headers.indexOf('Gênero')] = genero || '';
        novaLinha[headers.indexOf('Cor_Manga')] = cor_manga || '';
       
        materiaSheet.appendRow(novaLinha);
        const nomeCompleto = `${modelo} ${genero || ''} ${cor} ${cor_manga ? 'Manga ' + cor_manga : ''} ${tamanho}`.replace(/\s+/g, ' ').trim();
        sendTelegramMessage(chat_id, `✅ Nova matéria-prima adicionada: ${quantidade}x ${nomeCompleto.toUpperCase()}.`);
    }

  } catch (e) {
    Logger.log(`Erro em handleMateriaPrimaCommand: ${e.stack}`);
    sendTelegramMessage(chat_id, `Ocorreu um erro ao adicionar a matéria-prima: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

function handleReportQueryCommand(chat_id, commandText) {
    let reportText = "📊 <b>Relatório solicitado:</b>\n\n";
    const today = new Date();
    const vendasLogData = readSheetData(SHEET_NAMES.VENDAS_LOG);
    const vendasLogSheet = getSheet(SHEET_NAMES.VENDAS_LOG);
    const headers = vendasLogSheet.getDataRange().getValues()[0];

    const lotesData = readSheetData(SHEET_NAMES.LOTES_PRODUCAO);
    const lotesSheet = getSheet(SHEET_NAMES.LOTES_PRODUCAO);
    const lotesHeaders = lotesSheet.getDataRange().getValues()[0];

    const lowerCommand = commandText.toLowerCase();

    if (lowerCommand.includes('relatorio do dia') || lowerCommand.includes('vendas hoje')) {
        const todaySales = vendasLogData.filter(row => {
            const saleDate = new Date(row[headers.indexOf('Data_Hora')]);
            return saleDate.toDateString() === today.toDateString();
        });
        const totalValueToday = todaySales.reduce((sum, row) => sum + parseFloat(row[headers.indexOf('Valor')]), 0);
        const totalItemsToday = todaySales.reduce((sum, row) => sum + parseInt(row[headers.indexOf('Quantidade')], 10), 0);
        reportText += `<b>📈 Vendas de Hoje (${Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy')}):</b>\n`;
        reportText += `  - Valor Total: R$ ${totalValueToday.toFixed(2)}\n`;
        reportText += `  - Total de Itens: ${totalItemsToday}\n`;
        if (todaySales.length > 0) {
            reportText += `  - Detalhes:\n`;
            todaySales.forEach(sale => {
                reportText += `    - ${sale[headers.indexOf('Produto_Completo')]} (${sale[headers.indexOf('Quantidade')]} un.) por ${sale[headers.indexOf('Vendedor')]} (R$ ${parseFloat(sale[headers.indexOf('Valor')]).toFixed(2)})\n`;
            });
        }
    } else if (lowerCommand.includes('relatorio vendas semana')) {
        const startOfWeek = new Date(today);
        startOfWeek.setDate(today.getDate() - today.getDay());
        startOfWeek.setHours(0,0,0,0);
        const weekSales = vendasLogData.filter(row => new Date(row[headers.indexOf('Data_Hora')]) >= startOfWeek);
        const totalValueWeek = weekSales.reduce((sum, row) => sum + parseFloat(row[headers.indexOf('Valor')]), 0);
        const totalItemsWeek = weekSales.reduce((sum, row) => sum + parseInt(row[headers.indexOf('Quantidade')], 10), 0);
        reportText += `<b>📈 Vendas desta Semana (a partir de ${Utilities.formatDate(startOfWeek, Session.getScriptTimeZone(), 'dd/MM')}):</b>\n`;
        reportText += `  - Valor Total: R$ ${totalValueWeek.toFixed(2)}\n`;
        reportText += `  - Total de Itens: ${totalItemsWeek}\n`;
    } else if (lowerCommand.includes('relatorio vendas mes')) {
        const monthSales = vendasLogData.filter(row => new Date(row[headers.indexOf('Data_Hora')]).getMonth() === today.getMonth());
        const totalValueMonth = monthSales.reduce((sum, row) => sum + parseFloat(row[headers.indexOf('Valor')]), 0);
        const totalItemsMonth = monthSales.reduce((sum, row) => sum + parseInt(row[headers.indexOf('Quantidade')], 10), 0);
        reportText += `<b>📈 Vendas deste Mês (${Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM')}):</b>\n`;
        reportText += `  - Valor Total: R$ ${totalValueMonth.toFixed(2)}\n`;
        reportText += `  - Total de Itens: ${totalItemsMonth}\n`;
    } else if (lowerCommand.includes('top produtos')) {
        const productSales = {};
        vendasLogData.forEach(row => {
            const productName = row[headers.indexOf('Produto_Completo')];
            const quantity = parseInt(row[headers.indexOf('Quantidade')], 10);
            productSales[productName] = (productSales[productName] || 0) + quantity;
        });
        const sortedProducts = Object.entries(productSales).sort(([, a], [, b]) => b - a);
        reportText += "<b>⭐ Top 5 Produtos Mais Vendidos:</b>\n";
        sortedProducts.slice(0, 5).forEach(([product, qty], index) => {
            reportText += ` ${index + 1}. ${product}: ${qty} un.\n`;
        });
    } else if (lowerCommand.includes('relatorio producao')) {
        const pendingOrders = lotesData.filter(row => String(row[lotesHeaders.indexOf('Status')]).toLowerCase().trim() === 'aguardando produção');
        reportText += "<b>🏭 Lotes Aguardando Produção:</b>\n";
        if (pendingOrders.length > 0) {
            pendingOrders.forEach(order => { 
                reportText += ` - <b>ID:</b> <code>${order[lotesHeaders.indexOf('ID_Lote')]}</code>\n`;
                reportText += `   <b>Descrição:</b> ${order[lotesHeaders.indexOf('Descricao')]}\n`;
                reportText += `   <b>Data:</b> ${Utilities.formatDate(new Date(order[lotesHeaders.indexOf('Data_Pedido')]), Session.getScriptTimeZone(), 'dd/MM/yyyy')}\n`;
            });
        } else {
            reportText += "  Nenhum lote pendente.\n";
        }
    } else {
        reportText += "Comando de relatório não reconhecido. Tente 'relatorio do dia', '...semana', '...mes', 'top produtos' ou 'relatorio producao'.";
    }

    sendTelegramMessage(chat_id, reportText);
}


// --- FUNÇÕES DE ADMINISTRAÇÃO, GATILHOS E TESTES ---

function setWebhook() {
  try {
    loadConfigurations();
    const telegramToken = CONFIG.SISTEMA.TELEGRAM_TOKEN;
    const webAppUrl = CONFIG.SISTEMA.WEBHOOK_URL; 

    if (!telegramToken || !webAppUrl) {
      const errorMessage = `Erro: ${!telegramToken ? 'TELEGRAM_TOKEN' : 'WEBHOOK_URL'} não configurado na aba CONFIGURACOES.`;
      Logger.log(errorMessage);
      SpreadsheetApp.getUi().alert("Erro de Configuração", errorMessage, SpreadsheetApp.getUi().ButtonSet.OK); 
      return errorMessage;
    }

    const url = `https://api.telegram.org/bot${telegramToken}/setWebhook?url=${webAppUrl}`;
 
    const response = UrlFetchApp.fetch(url);
    const responseText = response.getContentText();
    Logger.log(`Webhook configurado: ${responseText}`);
    SpreadsheetApp.getUi().alert("Sucesso!", `Webhook configurado: ${responseText}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return "Webhook configurado com sucesso!";
  } catch (e) {
    Logger.log(`Erro ao configurar o webhook: ${e.message}`);
    SpreadsheetApp.getUi().alert("Erro!", `Erro ao configurar o webhook: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return `Erro ao configurar o webhook: ${e.message}`;
  }
}

function setupTimeDrivenTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'checkAllAlerts') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  loadConfigurations(); 
  const intervalMinutes = CONFIG.PARAMETROS_SISTEMA.ALERTA_INTERVALO_MINUTOS || 60; 
  if (CONFIG.SISTEMA.SISTEMA_ATIVO === 'SIM') {
     ScriptApp.newTrigger('checkAllAlerts')
        .timeBased()
        .everyMinutes(intervalMinutes)
        .create();
     Logger.log(`Gatilho de alertas criado para executar a cada ${intervalMinutes} minutos.`);
     return `Gatilho de alertas configurado para executar a cada ${intervalMinutes} minutos.`;
  } else {
     Logger.log("Sistema não está ativo. Gatilho de alertas não foi configurado.");
     return "Sistema não está ativo. Gatilho de alertas não foi configurado.";
  }
}

function deleteAllTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log("Todos os gatilhos do projeto foram removidos.");
  return "Todos os gatilhos do projeto foram removidos.";
}

function testOnboardingFlow() {
  const mock_callback_query = {
    "id": "1234567890123456789",
    "from": { "id": 1000505271, "is_bot": false, "first_name": "Breno Andrade" },
    "message": { "message_id": 12345, "chat": { "id": 1000505271, "type": "private" }},
    "data": "onboarding_next_2" 
  };

  loadConfigurations();
  
  handleOnboardingFlow(mock_callback_query);

  Logger.log("-> testOnboardingFlow executada. É esperado que apareça um erro nos logs sobre 'message to edit not found', porque o message_id é falso. O importante é analisar o log 'PAYLOAD' que foi gerado logo antes do erro.");
}


function testDoPost() {
  const mockEvent = {
    postData: {
      contents: JSON.stringify({
        message: {
          chat: { id: 1000505271 }, // Substitua pelo seu ID de chat do Telegram
          from: { id: 1000505271, first_name: 'Breno Andrade' }, // Substitua pelo seu ID e nome
          text: '/ajuda' // Testando o comando /ajuda
        }
      })
    }
  };
  doPost(mockEvent);
}

// PONTO DE ENTRADA PARA O APLICATIVO WEB (DASHBOARD)
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard.html')
    .setTitle("Painel de Gestão PRO")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkAllAlerts() {
  loadConfigurations();
  const alertsConfig = CONFIG.ALERTAS?.lista;
  if (!alertsConfig) {
    Logger.log("Nenhuma configuração de alerta encontrada.");
    return;
  }
 
  const today = new Date();
 
  alertsConfig.forEach(rule => {
    if (rule.ativo !== 'SIM') return;

    const notifyChatId = rule.notificar_chat || CONFIG.CHAT_IDS.GRUPO_ALERTAS;
    if (!notifyChatId) return;

    if (rule.tipo_alerta === 'estoque_baixo') {
      const minStock = parseInt(rule.estoque_minimo, 10);
      const materiaData = getSheet(SHEET_NAMES.ESTOQUE_MATERIA).getDataRange().getValues();
      const materiaHeaders = materiaData.shift();
      materiaData.forEach(row => {
        const currentQty = parseInt(row[materiaHeaders.indexOf('Qtd_Atual')], 10);
        if (currentQty <= minStock) {
          const productName = `${row[materiaHeaders.indexOf('Modelo')]} ${row[materiaHeaders.indexOf('Cor')]} ${row[materiaHeaders.indexOf('Tamanho')]}`;
          sendTelegramMessage(notifyChatId, `🔴 ALERTA DE ESTOQUE (MATÉRIA-PRIMA): ${productName} com apenas ${currentQty} unidades.`);
        }
      });
      const prontoData = getSheet(SHEET_NAMES.ESTOQUE_PRONTO).getDataRange().getValues();
      const prontoHeaders = prontoData.shift();
      prontoData.forEach(row => {
        const currentQty = parseInt(row[prontoHeaders.indexOf('Quantidade')], 10);
        if (currentQty <= minStock) {
          const productName = `${row[materiaHeaders.indexOf('Modelo')]} ${row[materiaHeaders.indexOf('Cor')]} ${row[materiaHeaders.indexOf('Tamanho')]} ${row[materiaHeaders.indexOf('Estampa')]}`;
          sendTelegramMessage(notifyChatId, `🔴 ALERTA DE ESTOQUE (PRODUTO PRONTO): ${productName} com apenas ${currentQty} unidades.`);
        }
      });
    }
   
    else if (rule.tipo_alerta === 'producao_atrasada') {
      const maxDays = parseInt(rule.estoque_minimo, 10); 
      const lotesData = readSheetData(SHEET_NAMES.LOTES_PRODUCAO);
      const lotesHeaders = getSheet(SHEET_NAMES.LOTES_PRODUCAO).getRange(1, 1, 1, getSheet(SHEET_NAMES.LOTES_PRODUCAO).getLastColumn()).getValues()[0];
     
      lotesData.forEach(lote => {
        const status = lote[lotesHeaders.indexOf('Status')];
        const dataPedido = new Date(lote[lotesHeaders.indexOf('Data_Pedido')]);
        const daysDiff = (today.getTime() - dataPedido.getTime()) / (1000 * 3600 * 24);
       
        if (status === 'Aguardando Produção' && daysDiff > maxDays) {
          const loteId = lote[lotesHeaders.indexOf('ID_Lote')];
          const descricao = lote[lotesHeaders.indexOf('Descricao')];
          sendTelegramMessage(notifyChatId, `⏰ ALERTA DE PRODUÇÃO: O lote "${descricao}" (ID: ${loteId}) está aguardando há mais de ${maxDays} dias.`);
        }
      });
    }
  });
}

/**
 * NOVO: FUNÇÃO PARA TESTAR O FLUXO DE CRIAÇÃO DE LOTE
 * Execute esta função diretamente no editor para simular a criação de um lote.
 * Altere os valores dentro da função para testar diferentes cenários.
 */
function testLoteCreationFlow() {
  // 1. Simular o utilizador que está a iniciar o fluxo
  const mockUser = { id: 1000505271, name: 'Breno Andrade' }; // Substitua pelo seu ID e nome
  
  // 2. Simular o início do fluxo
  Logger.log("--- INICIANDO TESTE DE CRIAÇÃO DE LOTE ---");
  iniciarNovoLote(mockUser);
  Logger.log("Estado inicial salvo para o utilizador: " + userProperties.getProperty('state_' + mockUser.id));

  // 3. Simular a resposta do utilizador para a descrição
  const mockDescricao = "Lote de Teste Automático";
  processCommand(mockUser.id, mockDescricao, mockUser.name, mockUser);
  Logger.log(`Descrição "${mockDescricao}" processada. Novo estado: ` + userProperties.getProperty('state_' + mockUser.id));

  // 4. Simular o clique no botão "Adicionar Item"
  const mockCallbackAdicionarItem = {
    from: { id: mockUser.id, first_name: mockUser.name },
    message: { chat: { id: mockUser.id }, message_id: 12345 }, // message_id é um placeholder
    data: "lote_add_item"
  };
  handleCallbackQuery(mockCallbackAdicionarItem);
  Logger.log("Clique em 'Adicionar Item' processado. Novo estado: " + userProperties.getProperty('state_' + mockUser.id));

  // 5. Simular a seleção de Modelo, Cor e Tamanho
  const selections = [
    { data: "lote_set_modelo_T-Shirt", log: "Modelo 'T-Shirt' selecionado." },
    { data: "lote_set_cor_Preta", log: "Cor 'Preta' selecionada." },
    { data: "lote_set_tamanho_G", log: "Tamanho 'G' selecionado." }
  ];

  selections.forEach(selection => {
    const mockCallback = {
      from: { id: mockUser.id, first_name: mockUser.name },
      message: { chat: { id: mockUser.id }, message_id: 12345 },
      data: selection.data
    };
    handleCallbackQuery(mockCallback);
    Logger.log(selection.log + " Novo estado: " + userProperties.getProperty('state_' + mockUser.id));
  });

  // 6. Simular a digitação da estampa
  const mockEstampa = "Estampa de Teste";
  processCommand(mockUser.id, mockEstampa, mockUser.name, mockUser);
  Logger.log(`Estampa "${mockEstampa}" processada. Novo estado: ` + userProperties.getProperty('state_' + mockUser.id));

  // 7. Simular a seleção da quantidade
  const mockCallbackQtd = {
      from: { id: mockUser.id, first_name: mockUser.name },
      message: { chat: { id: mockUser.id }, message_id: 12345 },
      data: "lote_set_qtd_5"
  };
  handleCallbackQuery(mockCallbackQtd);
  Logger.log("Quantidade '5' selecionada. Novo estado: " + userProperties.getProperty('state_' + mockUser.id));
  
  // 8. Simular a finalização do lote
  Logger.log("--- SIMULANDO FINALIZAÇÃO DO LOTE ---");
  const mockCallbackFinalizar = {
      from: { id: mockUser.id, first_name: mockUser.name },
      message: { chat: { id: mockUser.id }, message_id: 12345 },
      data: "lote_finish"
  };
  handleCallbackQuery(mockCallbackFinalizar);
  Logger.log("--- TESTE CONCLUÍDO ---");
  Logger.log("Verifique as planilhas 'LOTES_PRODUCAO' e 'ITENS_LOTE' para ver o resultado. Verifique também o seu Telegram e o grupo de produção para as mensagens.");
  
  // Limpa o estado de teste no final
  userProperties.deleteProperty('state_' + mockUser.id);
}

