// --- FUNÇÕES DE GESTÃO DO WEBHOOK ---
// Execute estas funções manualmente no editor do Apps Script para gerir a conexão com o Telegram.

/**
 * NOVO: LIMPA O CACHE DO SCRIPT
 * Execute esta função se as configurações (como o token) não parecerem atualizar.
 * Isso força o script a reler os dados da planilha na próxima execução.
 */
function clearScriptCache() {
  try {
    const cache = CacheService.getScriptCache();
    // Limpa especificamente os caches que guardam as configurações e palavras-chave
    cache.removeAll(['CONFIG_CACHE', 'KEYWORDS_CACHE']);
    const message = 'Cache do script (CONFIG e KEYWORDS) foi limpo com sucesso!';
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    const errorMsg = `Erro ao limpar o cache: ${e.message}`;
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
  }
}


/**
 * VERIFICA O STATUS ATUAL DO WEBHOOK
 * Use esta função para ver qual URL o Telegram está a usar no momento.
 */
function getWebhookInfo() {
  try {
    // É necessário carregar as configurações para obter o token
    if (typeof loadConfigurations !== 'function') {
      throw new Error("A função 'loadConfigurations' não foi encontrada. Verifique se ela existe no arquivo Code.gs.");
    }
    loadConfigurations(); 
    
    const telegramToken = CONFIG.SISTEMA.TELEGRAM_TOKEN;
    if (!telegramToken) {
      const message = "TELEGRAM_TOKEN não encontrado na aba CONFIGURACOES.";
      Logger.log(message);
      SpreadsheetApp.getUi().alert(message); // Mostra um alerta na planilha
      return;
    }
    const url = `https://api.telegram.org/bot${telegramToken}/getWebhookInfo`;
    const response = UrlFetchApp.fetch(url);
    const responseText = response.getContentText();
    Logger.log(`Informação do Webhook: ${responseText}`);
    // Mostra o resultado num pop-up fácil de ler
    SpreadsheetApp.getUi().alert('Informação Atual do Webhook', responseText, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    const errorMsg = `Erro ao obter informação do webhook: ${e.message}`;
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
  }
}

/**
 * NOVO: REGISTA O URL ATUAL DO SCRIPT
 * Execute esta função para ver qual URL o seu script está a usar.
 * O URL correto para o webhook deve terminar com /exec.
 */
function logWebAppUrl() {
  try {
    const webAppUrl = ScriptApp.getService().getUrl();
    const message = `O URL atual do seu aplicativo da web é:\n\n${webAppUrl}\n\nCertifique-se de que este é o URL que aparece no getWebhookInfo e que ele termina com /exec.`;
    Logger.log(message);
    SpreadsheetApp.getUi().alert('URL do Aplicativo da Web', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    const errorMsg = `Erro ao obter o URL do aplicativo da web: ${e.message}`;
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
  }
}

/**
 * DELETA O WEBHOOK ATUAL
 * Execute esta função se precisar de limpar a configuração atual.
 */
function deleteWebhook() {
  try {
    if (typeof loadConfigurations !== 'function') {
      throw new Error("A função 'loadConfigurations' não foi encontrada.");
    }
    loadConfigurations();
    
    const telegramToken = CONFIG.SISTEMA.TELEGRAM_TOKEN;
    if (!telegramToken) {
      throw new Error("TELEGRAM_TOKEN não encontrado.");
    }
    const url = `https://api.telegram.org/bot${telegramToken}/deleteWebhook`;
    const response = UrlFetchApp.fetch(url);
    const responseText = response.getContentText();
    Logger.log(`Resposta da exclusão do webhook: ${responseText}`);
    SpreadsheetApp.getUi().alert('Sucesso!', `Webhook deletado: ${responseText}`, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    const errorMsg = `Erro ao deletar o webhook: ${e.message}`;
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
  }
}

/**
 * NOVO: LIMPA O CACHE DO SCRIPT
 * Execute esta função se as configurações (como o token) não parecerem atualizar.
 * Isso força o script a reler os dados da planilha na próxima execução.
 */
function clearScriptCache() {
  try {
    const cache = CacheService.getScriptCache();
    // Limpa especificamente os caches que guardam as configurações e palavras-chave
    cache.removeAll(['CONFIG_CACHE', 'KEYWORDS_CACHE']);
    const message = 'Cache do script (CONFIG e KEYWORDS) foi limpo com sucesso!';
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    const errorMsg = `Erro ao limpar o cache: ${e.message}`;
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
  }
}

/**
 * LOGA O URL ATUAL DO APLICATIVO WEB
 * Execute esta função para ver o URL que deve ser colocado na sua planilha de configurações.
 */
function logWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  Logger.log(`URL do Web App: ${url}`);
  SpreadsheetApp.getUi().alert("URL do Aplicativo Web", `O URL do seu script é:\n\n${url}\n\nCopie e cole este valor na sua aba 'CONFIGURACOES' na linha 'WEBHOOK_URL'.`, SpreadsheetApp.getUi().ButtonSet.OK);
}
