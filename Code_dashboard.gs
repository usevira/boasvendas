/**
 * ARQUIVO: Code_Dashboard.gs
 * DESCRIÇÃO: Atua como o backend para o Dashboard.html, com lógica atualizada para
 * criar uma ordem de separação de matéria-prima consolidada e gerir formulários de massa.
 */

// --- Funções Auxiliares Genéricas ---

function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(sheetName) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        throw new Error(`A aba "${sheetName}" não foi encontrada.`);
    }
    return sheet;
}

function getHeaders(sheetName) {
    const sheet = getSheet(sheetName);
    // Garante que leia apenas colunas que realmente têm cabeçalhos
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return []; // Aba vazia
    const range = sheet.getRange(1, 1, 1, lastCol);
    if (!range) return [];
    const headers = range.getValues()[0];
    // Remove colunas vazias do final, se houver
    while (headers.length > 0 && headers[headers.length - 1] === "") {
        headers.pop();
    }
    return headers;
}


function readSheetData(sheetName) {
    const sheet = getSheet(sheetName);
    // Evita ler a planilha inteira se ela for muito grande ou tiver muitas linhas vazias
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // Retorna vazio se só tiver cabeçalho ou estiver vazia
    const lastCol = getHeaders(sheetName).length; // Usa o número real de headers
    if (lastCol === 0) return []; // Não há colunas com cabeçalho
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    return data;
}


// =================================================================
// PONTO DE ENTRADA PRINCIPAL PARA O DASHBOARD
// =================================================================

function getDashboardData() {
  try {
    Logger.log("Iniciando getDashboardData...");
    
    // Assegura que todas as abas necessárias existem
    Object.values(SHEET_NAMES).forEach(name => getSheet(name)); // Valida existência

    Logger.log("1. Carregando dados...");
    const estoqueProntoData = readSheetData(SHEET_NAMES.ESTOQUE_PRONTO);
    const estoqueMateriaData = readSheetData(SHEET_NAMES.ESTOQUE_MATERIA);
    const vendasData = readSheetData(SHEET_NAMES.VENDAS_LOG);
    const lotesData = readSheetData(SHEET_NAMES.LOTES_PRODUCAO);
    const itensLoteData = readSheetData(SHEET_NAMES.ITENS_LOTE);
    const consignadoData = readSheetData(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
    const fluxoCaixaData = readSheetData(SHEET_NAMES.FLUXO_CAIXA);
    const vendedoresData = readSheetData(SHEET_NAMES.VENDEDORES);
    Logger.log("   -> Dados carregados.");

    Logger.log("2. Processando dados...");
    const estoqueProntoHeaders = getHeaders(SHEET_NAMES.ESTOQUE_PRONTO);
    const estoquePronto = processEstoquePronto_(estoqueProntoData, estoqueProntoHeaders);

    const estoqueMateriaHeaders = getHeaders(SHEET_NAMES.ESTOQUE_MATERIA);
    const estoqueMateria = processEstoqueMateria_(estoqueMateriaData, estoqueMateriaHeaders);

    const vendasHeaders = getHeaders(SHEET_NAMES.VENDAS_LOG);
    const vendas = processVendasLog_(vendasData, vendasHeaders);

    const lotesHeaders = getHeaders(SHEET_NAMES.LOTES_PRODUCAO);
    const itensLoteHeaders = getHeaders(SHEET_NAMES.ITENS_LOTE);
    const lotes = processLotesProducao_(lotesData, lotesHeaders, itensLoteData, itensLoteHeaders);
    
    const ordemProducao = getOrdemProducao_(lotes); // Usa os lotes já processados

    const consignadoHeaders = getHeaders(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
    const estoqueConsignado = processEstoqueConsignado_(consignadoData, consignadoHeaders);

    const fluxoCaixaHeaders = getHeaders(SHEET_NAMES.FLUXO_CAIXA);
    const fluxoCaixa = processFluxoCaixa_(fluxoCaixaData, fluxoCaixaHeaders);

    const vendedoresHeaders = getHeaders(SHEET_NAMES.VENDEDORES);
    const nomeColIndex = vendedoresHeaders.indexOf('Nome');
    const statusColIndex = vendedoresHeaders.indexOf('Status');
    
    const vendedoresAtivos = vendedoresData
      .filter(row => (row[statusColIndex] || '').toString().trim().toLowerCase() === 'ativo')
      .map(row => row[nomeColIndex]);
      
    const revendedoresAtivos = [...vendedoresAtivos]; // Assume que todos ativos podem ser revendedores

    Logger.log("3. Calculando estatísticas e gráficos...");
    const stats = getDashboardStats_(estoquePronto, vendas, lotes);
    const chartData = getChartData_(vendas);

    Logger.log("4. Montando opções de formulário...");
    const modelos = [...new Set(estoqueMateriaData.map(row => row[estoqueMateriaHeaders.indexOf('Modelo')]).filter(Boolean))].sort();
    const cores = [...new Set(estoqueMateriaData.map(row => row[estoqueMateriaHeaders.indexOf('Cor')]).filter(Boolean))].sort();
    const sizeOrder = { 'P': 1, 'M': 2, 'G': 3, 'GG': 4, 'XG': 5 };
    const tamanhos = [...new Set(estoqueMateriaData.map(row => row[estoqueMateriaHeaders.indexOf('Tamanho')]).filter(Boolean))]
                      .sort((a, b) => (sizeOrder[a] || 99) - (sizeOrder[b] || 99));
    const formOptions = { modelos, cores, tamanhos };

    Logger.log("5. Montando objeto de retorno...");
    const dashboardData = {
      stats: stats,
      estoquePronto: estoquePronto,
      estoqueMateria: estoqueMateria,
      estoqueConsignado: estoqueConsignado, 
      vendasRecentes: vendas.slice(0, 50), // Limita a 50 para performance
      lotes: lotes,
      ordemProducao: ordemProducao,
      fluxoCaixa: fluxoCaixa.slice(0, 50), // Limita a 50 para performance
      vendedores: vendedoresAtivos, 
      revendedores: revendedoresAtivos,
      // Lista de produtos para dropdowns (Usa todos, não só os com estoque)
      produtos: estoquePronto.map(p => p.nomeCompleto), 
      charts: chartData,
      formOptions: formOptions 
    };

    Logger.log("getDashboardData concluído com sucesso. Retornando dados.");
    return dashboardData;

  } catch (e) {
    Logger.log(`Erro em getDashboardData: ${e.message} ${e.stack}`);
    // Retorna um erro estruturado para o frontend
    return { error: `Ocorreu um erro ao buscar os dados: ${e.message}` };
  }
}

// =================================================================
// FUNÇÕES DE PROCESSAMENTO DE DADOS (Chamadas por getDashboardData)
// =================================================================

function processEstoquePronto_(data, headers) {
  const modeloIndex = headers.indexOf('Modelo');
  const generoIndex = headers.indexOf('Gênero');
  const corIndex = headers.indexOf('Cor');
  const corMangaIndex = headers.indexOf('Cor_Manga');
  const tamanhoIndex = headers.indexOf('Tamanho');
  const estampaIndex = headers.indexOf('Estampa');
  const qtdIndex = headers.indexOf('Quantidade');
  const precoIndex = headers.indexOf('Preço');
  const custoIndex = headers.indexOf('Custo_Unitario'); 
  const statusIndex = headers.indexOf('Status');

  return data.map(row => {
    const modelo = row[modeloIndex] || '';
    const genero = row[generoIndex] || '';
    const cor = row[corIndex] || '';
    const corManga = row[corMangaIndex] || '';
    const tamanho = row[tamanhoIndex] || '';
    const estampa = row[estampaIndex] || '';
    
    return {
      modelo, genero, cor, corManga, tamanho, estampa, 
      nomeCompleto: `${modelo} ${genero} ${cor} ${corManga ? 'Manga ' + corManga : ''} ${tamanho} ${estampa}`.replace(/\s+/g, ' ').trim(),
      quantidade: parseInt(row[qtdIndex], 10) || 0,
      preco: parseFloat(String(row[precoIndex]).replace(",", ".")) || 0,
      custo: parseFloat(String(row[custoIndex]).replace(",", ".")) || 0,
      status: row[statusIndex]
    };
  }).sort((a,b) => a.nomeCompleto.localeCompare(b.nomeCompleto));
}

function processEstoqueMateria_(data, headers) {
  const modeloIndex = headers.indexOf('Modelo');
  const generoIndex = headers.indexOf('Gênero');
  const corIndex = headers.indexOf('Cor');
  const corMangaIndex = headers.indexOf('Cor_Manga');
  const tamanhoIndex = headers.indexOf('Tamanho');
  const qtdAtualIndex = headers.indexOf('Qtd_Atual');
  const qtdReservadaIndex = headers.indexOf('Qtd_Reservada');

  return data.map(row => {
    const modelo = row[modeloIndex] || '';
    const genero = row[generoIndex] || '';
    const cor = row[corIndex] || '';
    const corManga = row[corMangaIndex] || '';
    const tamanho = row[tamanhoIndex] || '';
    const qtdAtual = parseInt(row[qtdAtualIndex], 10) || 0;
    const qtdReservada = parseInt(row[qtdReservadaIndex], 10) || 0;

    return {
      nomeCompleto: `${modelo} ${genero} ${cor} ${corManga ? 'Manga ' + corManga : ''} ${tamanho}`.replace(/\s+/g, ' ').trim(),
      qtdAtual: qtdAtual,
      qtdReservada: qtdReservada,
      qtdDisponivel: qtdAtual - qtdReservada
    };
  }).sort((a,b) => a.nomeCompleto.localeCompare(b.nomeCompleto));
}

function processEstoqueConsignado_(data, headers) {
    const revendedorCol = headers.indexOf('Revendedor');
    const produtoCol = headers.indexOf('Produto');
    const dataEnvioCol = headers.indexOf('Data_Envio');
    const qtdEnviadaCol = headers.indexOf('Qtd_Enviada');
    const qtdVendidaCol = headers.indexOf('Qtd_Vendida');
    const qtdRetornadaCol = headers.indexOf('Qtd_Retornada');
    const qtdRestanteCol = headers.indexOf('Qtd_Restante');
    const statusCol = headers.indexOf('Status');

    const localKeywords = loadKeywords_(); 
    const sizeOrder = { 'P': 1, 'M': 2, 'G': 3, 'GG': 4, 'XG': 5 };

    const mappedData = data.map(row => {
        const dataEnvio = row[dataEnvioCol];
        const produto = row[produtoCol] || '';
        const detalhes = extractProductDetails_(produto, localKeywords);

        return {
            revendedor: row[revendedorCol],
            produto: produto,
            dataEnvio: dataEnvio instanceof Date ? dataEnvio.toISOString() : dataEnvio, // Formato ISO para JS
            qtdEnviada: parseInt(row[qtdEnviadaCol], 10) || 0,
            qtdVendida: parseInt(row[qtdVendidaCol], 10) || 0,
            qtdRetornada: parseInt(row[qtdRetornadaCol], 10) || 0,
            qtdRestante: parseInt(row[qtdRestanteCol], 10) || 0,
            status: row[statusCol],
            _modelo: detalhes.modelo || '',
            _cor: detalhes.cor || '',
            _estampa: detalhes.estampa || '',
            _tamanhoOrdem: sizeOrder[detalhes.tamanho] || 99, 
        };
    });

    // Ordenação mais robusta
    return mappedData.sort((a, b) => {
        const revendedorCompare = (a.revendedor || '').localeCompare(b.revendedor || '');
        if (revendedorCompare !== 0) return revendedorCompare;
        
        // Ordena por data de envio mais recente primeiro dentro do mesmo revendedor
        const dataCompare = new Date(b.dataEnvio) - new Date(a.dataEnvio);
        if (dataCompare !== 0) return dataCompare;

        // Se datas iguais, ordena por produto
        const modeloCompare = a._modelo.localeCompare(b._modelo);
        if (modeloCompare !== 0) return modeloCompare;
        const corCompare = a._cor.localeCompare(b._cor);
        if (corCompare !== 0) return corCompare;
        const estampaCompare = a._estampa.localeCompare(b._estampa);
        if (estampaCompare !== 0) return estampaCompare;
        return a._tamanhoOrdem - b._tamanhoOrdem;
    });
}


function processVendasLog_(data, headers) {
  const dataHoraIndex = headers.indexOf('Data_Hora');
  const vendedorIndex = headers.indexOf('Vendedor');
  const produtoIndex = headers.indexOf('Produto_Completo');
  const qtdIndex = headers.indexOf('Quantidade');
  const valorIndex = headers.indexOf('Valor');
  const custoIndex = headers.indexOf('Custo_Total'); 
  const lucroIndex = headers.indexOf('Lucro');     

  return data.map(row => {
    const dataHora = row[dataHoraIndex];
    return {
      dataHora: dataHora instanceof Date ? dataHora.toISOString() : dataHora, // Formato ISO para JS
      vendedor: row[vendedorIndex],
      produto: row[produtoIndex],
      quantidade: parseInt(row[qtdIndex], 10) || 0,
      valor: parseFloat(String(row[valorIndex]).replace(",", ".")) || 0,
      custo: parseFloat(String(row[custoIndex]).replace(",", ".")) || 0,
      lucro: parseFloat(String(row[lucroIndex]).replace(",", ".")) || 0
    };
  }).sort((a, b) => new Date(b.dataHora) - new Date(a.dataHora)); // Mais recentes primeiro
}

function processLotesProducao_(lotesData, lotesHeaders, itensData, itensHeaders) {
  const loteIdIndex = lotesHeaders.indexOf('ID_Lote');
  const descIndex = lotesHeaders.indexOf('Descricao');
  const solIndex = lotesHeaders.indexOf('Solicitante');
  const dataIndex = lotesHeaders.indexOf('Data_Criacao');
  const statusIndex = lotesHeaders.indexOf('Status');
  
  const itemLoteIdIndex = itensHeaders.indexOf('ID_Lote');
  const itemProdIndex = itensHeaders.indexOf('Produto_Final_Completo');
  const itemQtdIndex = itensHeaders.indexOf('Quantidade');

  const itensPorLote = itensData.reduce((acc, itemRow) => {
    const loteId = itemRow[itemLoteIdIndex];
    if (!acc[loteId]) acc[loteId] = [];
    acc[loteId].push({ 
      produto: itemRow[itemProdIndex], 
      quantidade: parseInt(itemRow[itemQtdIndex], 10) || 0 // Garante número
    });
    return acc;
  }, {});

  return lotesData.map(row => {
    const loteId = row[loteIdIndex];
    const dataPedido = row[dataIndex]; 
    return {
      id: loteId,
      descricao: row[descIndex],
      solicitante: row[solIndex],
      dataPedido: dataPedido instanceof Date ? dataPedido.toISOString() : dataPedido, // Formato ISO para JS
      status: row[statusIndex],
      itens: itensPorLote[loteId] || [] 
    };
  }).sort((a, b) => new Date(b.dataPedido) - new Date(a.dataPedido)); // Mais recentes primeiro
}

function processFluxoCaixa_(data, headers) {
  const dataIndex = headers.indexOf('Data');
  const tipoIndex = headers.indexOf('Tipo');
  // Handle potential variations in header names
  const descIndex = headers.indexOf('Descrição') !== -1 ? headers.indexOf('Descrição') : headers.indexOf('Descricao'); 
  const entradaIndex = headers.indexOf('Entrada');
  const saidaIndex = headers.indexOf('Saída') !== -1 ? headers.indexOf('Saída') : headers.indexOf('Saida'); 
  const respIndex = headers.indexOf('Responsável') !== -1 ? headers.indexOf('Responsável') : headers.indexOf('Responsavel'); 

  return data.map(row => {
    const dataFluxo = row[dataIndex];
    return {
      data: dataFluxo instanceof Date ? dataFluxo.toISOString() : dataFluxo, // Formato ISO para JS
      tipo: row[tipoIndex],
      descricao: row[descIndex],
      entrada: parseFloat(String(row[entradaIndex]).replace(",", ".")) || 0,
      saida: parseFloat(String(row[saidaIndex]).replace(",", ".")) || 0,
      responsavel: row[respIndex]
    };
  }).sort((a, b) => new Date(b.data) - new Date(a.data)); // Mais recentes primeiro
}

// =================================================================
// FUNÇÕES DE CÁLCULO DE ESTATÍSTICAS E GRÁFICOS
// =================================================================

function getDashboardStats_(estoque, vendas, lotes) {
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();

  const vendasMes = vendas.filter(v => {
    const dataVenda = new Date(v.dataHora);
    return dataVenda.getMonth() === mesAtual && dataVenda.getFullYear() === anoAtual;
  });

  const faturamentoMes = vendasMes.reduce((acc, v) => acc + (v.valor || 0), 0);
  const lucroMes = vendasMes.reduce((acc, v) => acc + (v.lucro || 0), 0); // Usa o lucro já calculado

  return {
    totalItensEstoque: estoque.reduce((acc, item) => acc + item.quantidade, 0),
    valorEstoque: estoque.reduce((acc, item) => acc + (item.quantidade * item.preco), 0),
    faturamentoMes: faturamentoMes,
    lucroMes: lucroMes,
    lotesPendentes: lotes.filter(l => l.status === 'Aguardando Produção').length,
  };
}

function getChartData_(vendas) {
    const hoje = new Date();
    
    // Vendas nos Últimos 30 Dias (Agrupado por Dia)
    const salesLast30Days = {};
    // Cria chaves para todos os 30 dias para garantir que dias sem vendas apareçam com 0
    for (let i = 29; i >= 0; i--) {
        const d = new Date();
        d.setDate(d.getDate() - i);
        const key = `${d.getDate().toString().padStart(2, '0')}/${(d.getMonth() + 1).toString().padStart(2, '0')}`;
        salesLast30Days[key] = 0;
    }
    // Soma as vendas nos dias correspondentes
    vendas.forEach(v => {
        const d = new Date(v.dataHora);
        const diffDays = Math.floor((hoje - d) / (1000 * 60 * 60 * 24)); // Diferença em dias inteiros
        if (diffDays >= 0 && diffDays < 30) { // Garante que está nos últimos 30 dias
            const key = `${d.getDate().toString().padStart(2, '0')}/${(d.getMonth() + 1).toString().padStart(2, '0')}`;
            // Verifica se a chave existe (deveria existir sempre por causa do loop anterior)
            if(salesLast30Days.hasOwnProperty(key)) { 
                salesLast30Days[key] += v.valor;
            }
        }
    });
    // Garante a ordem correta dos dias antes de retornar
    const salesData = Object.entries(salesLast30Days)
                          .sort((a, b) => {
                              const [dayA, monthA] = a[0].split('/');
                              const [dayB, monthB] = b[0].split('/');
                              // Cria datas (ano irrelevante, só para ordenar mês/dia corretamente)
                              const dateA = new Date(2000, parseInt(monthA)-1, parseInt(dayA)); 
                              const dateB = new Date(2000, parseInt(monthB)-1, parseInt(dayB));
                              return dateA - dateB;
                           })
                           .map(([day, total]) => [day, total]);

    // Top Produtos e Vendas por Vendedor (Mês Atual)
    const vendasMesAtual = vendas.filter(v => {
        const dataVenda = new Date(v.dataHora);
        return dataVenda.getMonth() === hoje.getMonth() && dataVenda.getFullYear() === hoje.getFullYear();
    });

    const topProducts = vendasMesAtual.reduce((acc, v) => {
        acc[v.produto] = (acc[v.produto] || 0) + v.quantidade;
        return acc;
    }, {});
    const topProductsData = Object.entries(topProducts)
        .sort(([, a], [, b]) => b - a)
        .slice(0, 5) // Top 5
        .map(([name, qty]) => [name, qty]);

    const salesBySeller = vendasMesAtual.reduce((acc, v) => {
        // Ignora vendedores vazios ou não definidos
        if (v.vendedor) { 
           acc[v.vendedor] = (acc[v.vendedor] || 0) + v.valor;
        }
        return acc;
    }, {});
    const salesBySellerData = Object.entries(salesBySeller).map(([name, total]) => [name, total]);

    // Retorna dados formatados ou placeholders se vazios
    return {
        salesLast30Days: salesData.length > 0 ? salesData : [['Nenhum', 0]],
        topProducts: topProductsData.length > 0 ? topProductsData : [['Nenhum', 0]],
        salesBySeller: salesBySellerData.length > 0 ? salesBySellerData : [['Nenhum', 0]]
    };
}


function getOrdemProducao_(lotes) {
    const localKeywords = loadKeywords_();
    const ordemProducaoMap = new Map(); // Usar Map para agrupar e somar

    const lotesPendentes = lotes.filter(lote => lote.status === 'Aguardando Produção');

    lotesPendentes.forEach(lote => {
        if (lote.itens && Array.isArray(lote.itens)) {
            lote.itens.forEach(item => {
                const produtoFinal = item.produto;
                const quantidade = parseInt(item.quantidade, 10) || 0;
                
                if (produtoFinal && quantidade > 0) {
                    const components = extractProductDetails_(produtoFinal, localKeywords);
                    // Constrói nome da peça lisa de forma consistente
                    const pecaLisa = [
                        components.modelo,
                        components.genero,
                        components.cor,
                        components.cor_manga ? 'Manga ' + components.cor_manga : '',
                        components.tamanho
                    ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim(); // filter(Boolean) remove nulos/vazios

                    if (pecaLisa && components.modelo && components.cor && components.tamanho) { // Garante dados mínimos
                        const estampa = components.estampa || 'N/A';
                        const key = `${pecaLisa}|${estampa}`; // Chave combinada para agrupar
                        
                        // Soma a quantidade ao item existente no Map ou cria um novo
                        const current = ordemProducaoMap.get(key) || { pecaLisa: pecaLisa, estampa: estampa, quantidade: 0 };
                        current.quantidade += quantidade;
                        ordemProducaoMap.set(key, current);
                    } else {
                         Logger.log(`AVISO: Não foi possível determinar a peça lisa para "${produtoFinal}" no lote ${lote.id}`);
                    }
                }
            });
        }
    });

    // Converter Map para Array e ordenar
    return Array.from(ordemProducaoMap.values())
           .sort((a, b) => a.pecaLisa.localeCompare(b.pecaLisa) || a.estampa.localeCompare(b.estampa));
}


// =================================================================
// FUNÇÕES DE AÇÃO (Chamadas pelo google.script.run do Frontend)
// =================================================================

function getAcertoData(payload) {
    try {
        const { revendedor, dataFim } = payload;
        if (!revendedor) throw new Error("Revendedor não especificado.");
        // Garante que dataFimObj é uma data válida, usando hoje se não for especificada
        const dataFimObj = dataFim ? new Date(dataFim + "T23:59:59") : new Date(); 

        const consignacaoData = readSheetData(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
        const headers = getHeaders(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
        const revendedorCol = headers.indexOf('Revendedor');
        const dataEnvioCol = headers.indexOf('Data_Envio');
        const produtoCol = headers.indexOf('Produto');
        const vendidaCol = headers.indexOf('Qtd_Vendida');
        const acertadaCol = headers.indexOf('Qtd_Acertada');
        const precoCol = headers.indexOf('Preco_Venda');

        // Valida se colunas essenciais existem
        if ([revendedorCol, dataEnvioCol, produtoCol, vendidaCol, acertadaCol, precoCol].includes(-1)) {
            throw new Error("Colunas faltando na aba ESTOQUE_CONSIGNACAO (Revendedor, Data_Envio, Produto, Qtd_Vendida, Qtd_Acertada, Preco_Venda).");
        }

        const vendedoresData = readSheetData(SHEET_NAMES.VENDEDORES);
        const vendedoresHeaders = getHeaders(SHEET_NAMES.VENDEDORES);
        const nomeCol = vendedoresHeaders.indexOf('Nome');
        const comissaoCol = vendedoresHeaders.indexOf('Comissao') !== -1 ? vendedoresHeaders.indexOf('Comissao') : vendedoresHeaders.indexOf('Comissão'); // Handle variations

        if (nomeCol === -1 || comissaoCol === -1) {
             throw new Error("Colunas 'Nome' ou 'Comissao'/'Comissão' faltando na aba Vendedores_Revendedores.");
        }
        
        const revendedorInfo = vendedoresData.find(row => row[nomeCol] === revendedor);
        if (!revendedorInfo) throw new Error(`Não foi possível encontrar o revendedor: ${revendedor}`);
        const comissaoPercent = parseFloat(String(revendedorInfo[comissaoCol]).replace(",", ".")) || 0.4; 

        let totalPecasVendidas = 0;
        let valorTotalVendido = 0;
        const itensParaAcerto = [];

        consignacaoData.forEach((row, index) => {
            try {
                const dataEnvioValue = row[dataEnvioCol];
                const dataEnvio = dataEnvioValue instanceof Date ? dataEnvioValue : (dataEnvioValue ? new Date(dataEnvioValue) : null);

                if (!dataEnvio || isNaN(dataEnvio.getTime())) {
                    Logger.log(`AVISO: Data de envio inválida na linha ${index + 2} da consignação.`);
                    return; 
                }

                const revendedorPlanilha = row[revendedorCol];
                const qtdVendida = parseInt(row[vendidaCol], 10) || 0;
                const qtdAcertada = parseInt(row[acertadaCol], 10) || 0;

                if (revendedorPlanilha === revendedor && dataEnvio <= dataFimObj && qtdVendida > qtdAcertada) {
                    const qtdPendente = qtdVendida - qtdAcertada;
                    const precoVenda = parseFloat(String(row[precoCol]).replace(",", ".")) || 0;
                    const valorPendente = qtdPendente * precoVenda;

                    totalPecasVendidas += qtdPendente;
                    valorTotalVendido += valorPendente;
                    itensParaAcerto.push({
                        produto: row[produtoCol],
                        qtdVendida: qtdPendente, // Mostra apenas a qtd pendente
                        valorTotal: valorPendente,
                        rowIndex: index + 2 // Linha na planilha (1-based + header)
                    });
                }
            } catch (dateError) {
                 Logger.log(`Erro ao processar data na linha ${index + 2} da consignação: ${dateError.message}`);
            }
        });

        const valorComissao = valorTotalVendido * comissaoPercent;
        const valorAReceber = valorTotalVendido - valorComissao;

        return {
            revendedor,
            totalPecasVendidas,
            valorTotalVendido,
            valorComissao,
            valorAReceber,
            itens: itensParaAcerto,
            comissaoPercent: comissaoPercent // Retorna o percentual usado
        };

    } catch (e) {
        Logger.log(`Erro em getAcertoData: ${e.stack}`);
        return { error: `Erro ao buscar dados para o acerto: ${e.message}`};
    }
}


function realizarAcertoPeloPainel(acertoInfo) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) { // Tenta travar por 30s
        return { success: false, message: 'Não foi possível executar a ação. O sistema está ocupado. Tente novamente.' };
    }
    try {
        const { revendedor, dataFim } = acertoInfo;
        
        // Recalcula os dados do acerto para garantir consistência e pegar a comissão correta
        const acertoData = getAcertoData({ revendedor, dataFim }); 

        if (acertoData.error) {
            return { success: false, message: acertoData.error };
        }
        if (acertoData.totalPecasVendidas === 0) {
            return { success: false, message: "Nenhuma venda pendente para acertar neste período." };
        }

        const consignacaoSheet = getSheet(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
        const headers = getHeaders(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
        const qtdAcertadaIndex = headers.indexOf('Qtd_Acertada');
        const statusIndex = headers.indexOf('Status');
        const qtdRestanteIndex = headers.indexOf('Qtd_Restante');
        const qtdVendidaIndex = headers.indexOf('Qtd_Vendida'); // Necessário para checar status final

        if (qtdAcertadaIndex === -1 || statusIndex === -1 || qtdRestanteIndex === -1 || qtdVendidaIndex === -1) {
            throw new Error("Colunas faltando na aba ESTOQUE_CONSIGNACAO (Qtd_Acertada, Status, Qtd_Restante, Qtd_Vendida).");
        }

        // --- Atualiza a planilha de consignação ---
        const rangesToUpdate = [];
        const valuesToUpdate = [];

        acertoData.itens.forEach(item => {
            const currentAcertada = consignacaoSheet.getRange(item.rowIndex, qtdAcertadaIndex + 1).getValue() || 0;
          	const newAcertada = currentAcertada + item.qtdVendida;
            
          	// Adiciona atualização da Qtd_Acertada
          	rangesToUpdate.push(consignacaoSheet.getRange(item.rowIndex, qtdAcertadaIndex + 1).getA1Notation());
          	valuesToUpdate.push(newAcertada);

          	// Verifica se o item foi totalmente vendido e acertado para mudar status
          	const qtdRestanteAtual = consignacaoSheet.getRange(item.rowIndex, qtdRestanteIndex + 1).getValue() || 0;
          	const qtdVendidaTotal = consignacaoSheet.getRange(item.rowIndex, qtdVendidaIndex + 1).getValue() || 0; 
            
          	if (qtdRestanteAtual === 0 && newAcertada === qtdVendidaTotal) {
                // Adiciona atualização do Status
                rangesToUpdate.push(consignacaoSheet.getRange(item.rowIndex, statusIndex + 1).getA1Notation());
                valuesToUpdate.push('Acertado e Finalizado');
          	} else if (newAcertada > 0 && newAcertada < qtdVendidaTotal) {
                // Se foi parcialmente acertado, mas ainda falta
                rangesToUpdate.push(consignacaoSheet.getRange(item.rowIndex, statusIndex + 1).getA1Notation());
                valuesToUpdate.push('Parcialmente Acertado');
          	} else if (newAcertada === qtdVendidaTotal && qtdRestanteAtual > 0) {
                // Totalmente acertado, mas ainda tem peça física com o revendedor
                rangesToUpdate.push(consignacaoSheet.getRange(item.rowIndex, statusIndex + 1).getA1Notation());
                valuesToUpdate.push('Acertado (Aguard. Retorno)');
          	}
        });
        
        // Aplica atualizações em lote se houverem
        if (rangesToUpdate.length > 0) {
          const sheet = consignacaoSheet.getParent().getSheetByName(SHEET_NAMES.ESTOQUE_CONSIGNACAO); // Garante a aba correta
          rangesToUpdate.forEach((rangeA1, index) => {
              sheet.getRange(rangeA1).setValue(valuesToUpdate[index]);
          });
          SpreadsheetApp.flush(); // Garante que as escritas sejam aplicadas
        }

        // --- Lançamento no fluxo de caixa ---
        const fluxoSheet = getSheet(SHEET_NAMES.FLUXO_CAIXA);
        const fluxoHeaders = getHeaders(SHEET_NAMES.FLUXO_CAIXA); // Pega cabeçalhos do fluxo
        const dataFimFormatada = dataFim ? new Date(dataFim+'T00:00:00').toLocaleDateString('pt-BR') : 'data atual';
        const dataLancamento = new Date();
        
        const { valorAReceber, valorComissao, totalPecasVendidas, valorTotalVendido, comissaoPercent } = acertoData;
        
        // Cria linhas com base nos cabeçalhos encontrados
        const entradaRow = new Array(fluxoHeaders.length).fill('');
        entradaRow[fluxoHeaders.indexOf('Data')] = dataLancamento;
        entradaRow[fluxoHeaders.indexOf('Tipo')] = 'Entrada';
        entradaRow[fluxoHeaders.indexOf('Descricao') !== -1 ? fluxoHeaders.indexOf('Descricao') : fluxoHeaders.indexOf('Descrição')] = `Acerto Consignação - ${revendedor} (até ${dataFimFormatada})`;
        entradaRow[fluxoHeaders.indexOf('Entrada')] = valorAReceber;
        const categoriaEntradaIndex = fluxoHeaders.indexOf('Categoria');
        if (categoriaEntradaIndex > -1) entradaRow[categoriaEntradaIndex] = 'Consignação';
        const respEntradaIndex = fluxoHeaders.indexOf('Responsavel') !== -1 ? fluxoHeaders.indexOf('Responsavel') : fluxoHeaders.indexOf('Responsável');
        if (respEntradaIndex > -1) entradaRow[respEntradaIndex] = 'Painel Admin';

        const saidaRow = new Array(fluxoHeaders.length).fill('');
        saidaRow[fluxoHeaders.indexOf('Data')] = dataLancamento;
        saidaRow[fluxoHeaders.indexOf('Tipo')] = 'Saída';
        saidaRow[fluxoHeaders.indexOf('Descricao') !== -1 ? fluxoHeaders.indexOf('Descricao') : fluxoHeaders.indexOf('Descrição')] = `Comissão Consignação (${(comissaoPercent*100).toFixed(0)}%) - ${revendedor} (até ${dataFimFormatada})`;
        saidaRow[fluxoHeaders.indexOf('Saida') !== -1 ? fluxoHeaders.indexOf('Saida') : fluxoHeaders.indexOf('Saída')] = valorComissao;
        const categoriaSaidaIndex = fluxoHeaders.indexOf('Categoria');
        if (categoriaSaidaIndex > -1) saidaRow[categoriaSaidaIndex] = 'Comissão';
        const respSaidaIndex = fluxoHeaders.indexOf('Responsavel') !== -1 ? fluxoHeaders.indexOf('Responsavel') : fluxoHeaders.indexOf('Responsável');
        if (respSaidaIndex > -1) saidaRow[respSaidaIndex] = 'Painel Admin';

        fluxoSheet.appendRow(entradaRow);
        fluxoSheet.appendRow(saidaRow);
        
        // Mensagem de sucesso formatada
        const mensagemSucesso = `Acerto para ${revendedor} realizado!\n\n` +
                              `Peças Acertadas: ${totalPecasVendidas}\n` +
                              `Valor Total Vendido: ${valorTotalVendido.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}\n` +
                              `Comissão (${(comissaoPercent * 100).toFixed(0)}%): ${valorComissao.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}\n` +
                              `Valor Recebido: ${valorAReceber.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}\n\n` +
                              `Lançamentos no caixa efetuados.`;

        return { success: true, message: mensagemSucesso };

    } catch (e) {
        Logger.log(`Erro em realizarAcertoPeloPainel: ${e.stack}`);
        // Tenta dar uma mensagem de erro mais específica
        const specificError = e.message.includes("Colunas faltando") ? e.message : `Ocorreu um erro inesperado. Verifique os logs.`;
        return { success: false, message: `Erro ao realizar o acerto: ${specificError}` };
    } finally {
        lock.releaseLock();
    }
}


function registrarVendaPeloPainel(vendaInfo) {
  const { produto, quantidade, vendedor } = vendaInfo;
  
  if (!produto || !quantidade || !vendedor || quantidade <= 0) {
    return { success: false, message: 'Produto, Quantidade e Vendedor são obrigatórios.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Sistema ocupado, tente novamente.' };
  }
  
  try {
    const estoqueSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
    const estoqueData = estoqueSheet.getDataRange().getValues(); // Lê com header
    const headers = estoqueData[0];
    
    // Índices das colunas
    const modeloIdx = headers.indexOf('Modelo');
    const generoIdx = headers.indexOf('Gênero');
    const corIdx = headers.indexOf('Cor');
    const corMangaIdx = headers.indexOf('Cor_Manga');
    const tamanhoIdx = headers.indexOf('Tamanho');
    const estampaIdx = headers.indexOf('Estampa');
    const qtdIdx = headers.indexOf('Quantidade');
    const precoIdx = headers.indexOf('Preço');
    const custoIdx = headers.indexOf('Custo_Unitario');
    const dataAttIdx = headers.indexOf('Data_Atualização');
    const statusIdx = headers.indexOf('Status');
    
    // Normaliza o nome do produto vindo do frontend
    const produtoNomeLower = produto.toLowerCase().trim();
    let produtoEncontradoRow = -1;
    let rowData = null;

    // Procura pelo produto no estoque
    for (let i = 1; i < estoqueData.length; i++) {
        const row = estoqueData[i];
        const nomeNaPlanilha = [
            row[modeloIdx], row[generoIdx], row[corIdx],
            row[corMangaIdx] ? 'Manga ' + row[corMangaIdx] : '',
            row[tamanhoIdx], row[estampaIdx]
        ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();

        if (nomeNaPlanilha === produtoNomeLower) {
            produtoEncontradoRow = i + 1; // Linha real na planilha
            rowData = row;
            break;
        }
    }
    
    if (produtoEncontradoRow === -1) {
       return { success: false, message: `Produto "${produto}" não encontrado no estoque.` };
    }
    
    // Verifica estoque
    const qtdAtual = parseInt(rowData[qtdIdx], 10) || 0;
    if (qtdAtual < quantidade) {
        return { success: false, message: `Estoque insuficiente para "${produto}". Disponível: ${qtdAtual}.` };
    }
    
    // 1. Abate do Estoque
    const novaQtd = qtdAtual - quantidade;
    estoqueSheet.getRange(produtoEncontradoRow, qtdIdx + 1).setValue(novaQtd);
    // Atualiza data e status
    if (dataAttIdx > -1) estoqueSheet.getRange(produtoEncontradoRow, dataAttIdx + 1).setValue(new Date());
    if (statusIdx > -1) estoqueSheet.getRange(produtoEncontradoRow, statusIdx + 1).setValue(novaQtd > 0 ? 'disponivel' : 'esgotado');

    // 2. Registra no VENDAS_LOG
    const vendasLogSheet = getSheet(SHEET_NAMES.VENDAS_LOG);
    const logHeaders = getHeaders(SHEET_NAMES.VENDAS_LOG);
    const newLogRow = new Array(logHeaders.length).fill('');
    
    const custoUnitario = parseFloat(String(rowData[custoIdx]).replace(",", ".")) || 0;
    const precoVendaUnitario = parseFloat(String(rowData[precoIdx]).replace(",", ".")) || 0; // Assume preço da tabela
    const custoTotalVenda = custoUnitario * quantidade;
    const valorTotalVenda = precoVendaUnitario * quantidade;
    const lucro = valorTotalVenda - custoTotalVenda;

    newLogRow[logHeaders.indexOf('ID')] = `VENDA-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss')}`;
    newLogRow[logHeaders.indexOf('Data_Hora')] = new Date();
    newLogRow[logHeaders.indexOf('Vendedor')] = vendedor;
    newLogRow[logHeaders.indexOf('Produto_Completo')] = produto.toUpperCase();
    newLogRow[logHeaders.indexOf('Quantidade')] = quantidade;
    newLogRow[logHeaders.indexOf('Tipo_Venda')] = 'varejo'; // Assume varejo, pode precisar de lógica de atacado
    newLogRow[logHeaders.indexOf('Canal')] = 'Painel Admin';
    newLogRow[logHeaders.indexOf('Valor')] = parseFloat(valorTotalVenda.toFixed(2));
    newLogRow[logHeaders.indexOf('Custo_Total')] = parseFloat(custoTotalVenda.toFixed(2));
    newLogRow[logHeaders.indexOf('Lucro')] = parseFloat(lucro.toFixed(2));
    newLogRow[logHeaders.indexOf('Status')] = 'concluída';
    // Forma_Pagamento pode ser adicionada ao formulário de Vendas
    
    vendasLogSheet.appendRow(newLogRow);
    SpreadsheetApp.flush(); // Garante a escrita

    return { success: true, message: `Venda registrada! Restam ${novaQtd} unidades de "${produto}".` };
    
  } catch (e) {
    Logger.log(`Erro em registrarVendaPeloPainel: ${e.stack}`);
    return { success: false, message: `Erro ao registrar venda: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function criarLotePeloPainel(loteInfo) {
    Logger.log("Chamando criarLotePeloPainel (simples)... redirecionando para criarLoteMassaPeloPainel");
    return criarLoteMassaPeloPainel(loteInfo); 
}

function enviarConsignacaoPeloPainel(consignacaoInfo) {
    Logger.log("Chamando enviarConsignacaoPeloPainel (simples)... redirecionando para realizarSaidaMassaPeloPainel");
    const saidaInfo = {
        revendedor: consignacaoInfo.revendedor,
        tipoSaida: "Consignação", // Define o tipo
        itens: consignacaoInfo.itens 
    };
    return realizarSaidaMassaPeloPainel(saidaInfo);
}


// =================================================================
// FUNÇÕES DE PRODUÇÃO E SAÍDA EM MASSA
// =================================================================

function criarLoteMassaPeloPainel(loteInfo) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) { // Tenta travar por 30s
        return { success: false, message: 'Não foi possível executar a ação. O sistema está ocupado. Tente novamente.' };
    }
    
    try {
        const { descricao, solicitante, itens } = loteInfo;
        if (!descricao || !solicitante || !itens || itens.length === 0) {
            return { success: false, message: 'Dados incompletos. É preciso descrição, solicitante e pelo menos um item.' };
        }

        const lotesSheet = getSheet(SHEET_NAMES.LOTES_PRODUCAO);
        const itensLoteSheet = getSheet(SHEET_NAMES.ITENS_LOTE);
        const materiaPrimaSheet = getSheet(SHEET_NAMES.ESTOQUE_MATERIA);
        const materiaData = materiaPrimaSheet.getDataRange().getValues(); // Lê com cabeçalho
        const materiaHeaders = materiaData[0];
        
        // Cria mapa de matéria-prima para acesso rápido e eficiente
        const materiaMap = new Map();
        // Índices das colunas relevantes
        const modeloIdx = materiaHeaders.indexOf('Modelo');
        const generoIdx = materiaHeaders.indexOf('Gênero');
        const corIdx = materiaHeaders.indexOf('Cor');
        const corMangaIdx = materiaHeaders.indexOf('Cor_Manga');
        const tamanhoIdx = materiaHeaders.indexOf('Tamanho');
        const qtdAtualIdx = materiaHeaders.indexOf('Qtd_Atual');
        const qtdReservadaIdx = materiaHeaders.indexOf('Qtd_Reservada');

        materiaData.slice(1).forEach((row, index) => { // Pula cabeçalho
            // Constrói a chave do mapa de forma consistente
            let nomeMateriaCompleto = [
                row[modeloIdx], row[generoIdx], row[corIdx], 
                row[corMangaIdx] ? 'Manga ' + row[corMangaIdx] : '', 
                row[tamanhoIdx]
            ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
            
            const qtdAtual = parseInt(row[qtdAtualIdx], 10) || 0;
            const qtdReservada = parseInt(row[qtdReservadaIdx], 10) || 0;
            
            materiaMap.set(nomeMateriaCompleto, { 
                rowIndex: index + 2, // Linha real na planilha (1-based + header)
                qtdDisponivel: qtdAtual - qtdReservada,
                // Referência direta à célula para atualização posterior
                cellReservada: materiaPrimaSheet.getRange(index + 2, qtdReservadaIdx + 1) 
            });
        });

        const KEYWORDS = loadKeywords_(); // Carrega palavras-chave locais
        const itensParaReservar = [];
        const errosEstoque = [];
        
        // --- Validação de Estoque ---
        for (const item of itens) {
            const { produto, quantidade } = item;
            if (!produto || !quantidade || quantidade <= 0) {
                 errosEstoque.push(`Item inválido encontrado: ${JSON.stringify(item)}`);
                 continue; // Pula item inválido
            }

            const detalhesItem = extractProductDetails_(produto, KEYWORDS);
            // Constrói nome da matéria-prima necessária de forma consistente
            const materiaPrimaNecessaria = [
                detalhesItem.modelo, detalhesItem.genero, detalhesItem.cor, 
                detalhesItem.cor_manga ? 'Manga ' + detalhesItem.cor_manga : '', 
                detalhesItem.tamanho
            ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();

            // Valida se a matéria-prima foi identificada corretamente
            if (!detalhesItem.modelo || !detalhesItem.cor || !detalhesItem.tamanho) {
                 errosEstoque.push(`Não foi possível identificar a matéria-prima base para "${produto}". Verifique o nome ou as palavras-chave.`);
                 continue; 
            }

            const materiaInfo = materiaMap.get(materiaPrimaNecessaria);
            if (!materiaInfo) {
                // Tenta encontrar sem gênero/cor de manga se não achar exato
                 const materiaBase = [detalhesItem.modelo, detalhesItem.cor, detalhesItem.tamanho]
                                      .filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
                 const materiaInfoBase = materiaMap.get(materiaBase);

                 if(materiaInfoBase) {
                     errosEstoque.push(`Matéria-prima "${materiaPrimaNecessaria}" não encontrada no estoque.`);
                 } else {
                    errosEstoque.push(`Matéria-prima "${materiaPrimaNecessaria}" não encontrada no estoque.`);
                 }
                 continue; // Pula para o próximo item
            } 
            
            if (materiaInfo.qtdDisponivel < quantidade) {
                errosEstoque.push(`Estoque insuficiente para "${materiaPrimaNecessaria}": Pedido ${quantidade}, Disponível ${materiaInfo.qtdDisponivel}.`);
            } else {
                // Adiciona à lista para reserva e atualiza a disponibilidade *no mapa* para a validação dos próximos itens
                itensParaReservar.push({ materiaInfo, quantidade, produtoFinal: produto });
                materiaInfo.qtdDisponivel -= quantidade; 
            }
        }
        
        // Se houveram erros, retorna sem modificar a planilha
        if (errosEstoque.length > 0) {
            return { success: false, message: `Erro de estoque:\n- ${errosEstoque.join('\n- ')}` };
        }
        
        // --- Criação do Lote e Reserva ---
        // Gera ID único para o lote
        const loteId = `LOTE-${new Date().toISOString().slice(2, 10).replace(/-/g, '')}-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss')}`; 
        
        // Adiciona o lote à planilha LOTES_PRODUCAO
        const lotesHeaders = getHeaders(SHEET_NAMES.LOTES_PRODUCAO);
        const newLoteRow = new Array(lotesHeaders.length).fill('');
        newLoteRow[lotesHeaders.indexOf('ID_Lote')] = loteId;
        newLoteRow[lotesHeaders.indexOf('Descricao')] = descricao;
        newLoteRow[lotesHeaders.indexOf('Solicitante')] = solicitante;
        newLoteRow[lotesHeaders.indexOf('Data_Criacao')] = new Date();
        newLoteRow[lotesHeaders.indexOf('Status')] = 'Aguardando Produção';
        lotesSheet.appendRow(newLoteRow);

        // Adiciona os itens à planilha ITENS_LOTE
        const itensParaPlanilha = itensParaReservar.map(item => [loteId, item.produtoFinal, item.quantidade]);
        if (itensParaPlanilha.length > 0) {
           // Adiciona em lote para performance
           itensLoteSheet.getRange(itensLoteSheet.getLastRow() + 1, 1, itensParaPlanilha.length, 3).setValues(itensParaPlanilha);
        }

        // Atualiza a quantidade reservada na planilha ESTOQUE_MATERIA em lote
        const updatesReservasRanges = [];
        const updatesReservasValues = [];
        for (const item of itensParaReservar) {
             updatesReservasRanges.push(item.materiaInfo.cellReservada.getA1Notation());
             updatesReservasValues.push((item.materiaInfo.cellReservada.getValue() || 0) + item.quantidade);
        }
        if(updatesReservasRanges.length > 0) {
            const sheet = materiaPrimaSheet.getParent().getSheetByName(SHEET_NAMES.ESTOQUE_MATERIA); // Pega a aba de novo
            updatesReservasRanges.forEach((rangeA1, index) => {
               sheet.getRange(rangeA1).setValue(updatesReservasValues[index]);
            });
            SpreadsheetApp.flush(); // Garante a escrita
        }

        return { success: true, message: `Lote "${loteId}" (${itens.length} tipo(s) de item) criado com sucesso e matéria-prima reservada.` };

    } catch(e) {
        Logger.log(`Erro em criarLoteMassaPeloPainel: ${e.stack}`);
        return { success: false, message: `Ocorreu um erro inesperado ao criar o lote: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}


function realizarSaidaMassaPeloPainel(saidaInfo) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) { // Tenta travar por 30s
        return { success: false, message: 'Não foi possível executar a ação. O sistema está ocupado. Tente novamente.' };
    }
    
    try {
        const { revendedor, tipoSaida, itens } = saidaInfo;
         if (!revendedor || !tipoSaida || !itens || itens.length === 0) {
            return { success: false, message: 'Dados incompletos. É preciso revendedor/destino, tipo de saída e pelo menos um item.' };
        }
        
        const estoqueProntoSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
        const estoqueData = estoqueProntoSheet.getDataRange().getValues(); // Lê com header
        const estoqueHeaders = estoqueData[0];
        
        // Cria mapa de estoque pronto para acesso rápido e eficiente
        const estoqueMap = new Map();
        // Indices das colunas
        const modeloIdx = estoqueHeaders.indexOf('Modelo');
        const generoIdx = estoqueHeaders.indexOf('Gênero');
        const corIdx = estoqueHeaders.indexOf('Cor');
        const corMangaIdx = estoqueHeaders.indexOf('Cor_Manga');
        const tamanhoIdx = estoqueHeaders.indexOf('Tamanho');
        const estampaIdx = estoqueHeaders.indexOf('Estampa');
        const qtdIdx = estoqueHeaders.indexOf('Quantidade');
        const precoIdx = estoqueHeaders.indexOf('Preço');

        estoqueData.slice(1).forEach((row, index) => { // Pula header
            // Constrói chave do mapa consistentemente
            const produtoNome = [
                row[modeloIdx], row[generoIdx], row[corIdx], 
                row[corMangaIdx] ? 'Manga ' + row[corMangaIdx] : '', 
                row[tamanhoIdx], row[estampaIdx]
            ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
            
            const quantidade = parseInt(row[qtdIdx], 10) || 0;
            
            estoqueMap.set(produtoNome.toLowerCase(), {
                rowIndex: index + 2, // Linha real na planilha
                quantidade: quantidade,
                preco: parseFloat(String(row[precoIdx]).replace(",", ".")) || 0,
                // Referência direta à célula para atualização
                cellQuantidade: estoqueProntoSheet.getRange(index + 2, qtdIdx + 1), 
                produtoNomeOriginal: produtoNome // Preserva nome original para logs/consignação
            });
        });
        
        const errosEstoque = [];
        const itensParaAbater = [];
        const itensParaConsignacao = [];
        
        // --- Validação de Estoque ---
        for (const item of itens) {
            const { produto, quantidade } = item;
             if (!produto || !quantidade || quantidade <= 0) {
                 errosEstoque.push(`Item inválido encontrado no pedido: ${JSON.stringify(item)}`);
                 continue; // Pula item inválido
            }

            const produtoInfo = estoqueMap.get(produto.toLowerCase());
            
            if (!produtoInfo) {
                errosEstoque.push(`Produto "${produto}" não encontrado no estoque pronto.`);
            } else if (produtoInfo.quantidade < quantidade) {
                errosEstoque.push(`Estoque insuficiente para "${produto}": Pedido ${quantidade}, Disponível ${produtoInfo.quantidade}.`);
            } else {
                // Adiciona à lista para abater e atualiza quantidade *no mapa* para validações futuras
                itensParaAbater.push({ produtoInfo, quantidade });
                produtoInfo.quantidade -= quantidade; 
                
                // Prepara dados para consignação, se aplicável
                if (tipoSaida === "Consignação") {
                    // Adiciona Qtd_Acertada inicializada como 0 e outras colunas padrão
                    const consignacaoHeaders = getHeaders(SHEET_NAMES.ESTOQUE_CONSIGNACAO); // Pega headers da consignação
                    const newRow = new Array(consignacaoHeaders.length).fill('');
                    newRow[consignacaoHeaders.indexOf('Revendedor')] = revendedor;
                    // newRow[consignacaoHeaders.indexOf('Ponto_Venda')] = ''; // Deixa vazio por padrão
                    newRow[consignacaoHeaders.indexOf('Produto')] = produtoInfo.produtoNomeOriginal;
                    newRow[consignacaoHeaders.indexOf('Data_Envio')] = new Date();
                    newRow[consignacaoHeaders.indexOf('Qtd_Enviada')] = quantidade;
                    newRow[consignacaoHeaders.indexOf('Qtd_Vendida')] = 0;
                    newRow[consignacaoHeaders.indexOf('Qtd_Acertada')] = 0; 
                    newRow[consignacaoHeaders.indexOf('Qtd_Retornada')] = 0;
                    newRow[consignacaoHeaders.indexOf('Qtd_Restante')] = quantidade;
                    newRow[consignacaoHeaders.indexOf('Preco_Venda')] = produtoInfo.preco;
                    newRow[consignacaoHeaders.indexOf('Status')] = 'Em consignação';
                    itensParaConsignacao.push(newRow);
                }
            }
        }
        
        // Se houveram erros, retorna sem modificar
        if (errosEstoque.length > 0) {
            return { success: false, message: `Erro de estoque:\n- ${errosEstoque.join('\n- ')}` };
        }
        
        // --- Aplica Abates no Estoque Pronto em lote ---
        const updatesEstoqueRanges = [];
        const updatesEstoqueValues = [];
        for (const item of itensParaAbater) {
             updatesEstoqueRanges.push(item.produtoInfo.cellQuantidade.getA1Notation());
             updatesEstoqueValues.push(Math.max(0, (item.produtoInfo.cellQuantidade.getValue() || 0) - item.quantidade));
        }
         if(updatesEstoqueRanges.length > 0) {
            const sheet = estoqueProntoSheet.getParent().getSheetByName(SHEET_NAMES.ESTOQUE_PRONTO); // Pega a aba de novo
            updatesEstoqueRanges.forEach((rangeA1, index) => {
               sheet.getRange(rangeA1).setValue(updatesEstoqueValues[index]);
            });
            SpreadsheetApp.flush(); // Garante a escrita
         }
        
        // --- Registra Saída em Consignação (se aplicável) em lote ---
        if (tipoSaida === "Consignação" && itensParaConsignacao.length > 0) {
            const consignacaoSheet = getSheet(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
            const numColsConsignacao = getHeaders(SHEET_NAMES.ESTOQUE_CONSIGNACAO).length; 
            // Garante que o número de colunas está correto
            const rowsToAdd = itensParaConsignacao.map(row => row.slice(0, numColsConsignacao)); 
            consignacaoSheet.getRange(consignacaoSheet.getLastRow() + 1, 1, rowsToAdd.length, numColsConsignacao).setValues(rowsToAdd);
        }
        
        // --- Lógica Adicional para outros Tipos de Saída ---
        if (tipoSaida === "Venda Atacado") {
            // IMPLEMENTAR: Registrar em VENDAS_LOG, FLUXO_CAIXA, gerar nota/recibo?
            Logger.log(`Saída de Venda Atacado para ${revendedor} (${itensParaAbater.length} tipo(s) de item) registrada (SIMULADO).`);
        }
        
         if (tipoSaida === "Ajuste de Estoque") {
            // IMPLEMENTAR: Registrar um log específico de ajustes? Motivo do ajuste?
            Logger.log(`Ajuste de Estoque para ${revendedor} (${itensParaAbater.length} tipo(s) de item) realizado (SIMULADO).`);
        }

        const totalItens = itensParaAbater.reduce((sum, item) => sum + item.quantidade, 0);
        return { success: true, message: `Saída em massa (${tipoSaida}) de ${totalItens} item(ns) para ${revendedor} realizada com sucesso.` };

    } catch(e) {
        Logger.log(`Erro em realizarSaidaMassaPeloPainel: ${e.stack}`);
        return { success: false, message: `Ocorreu um erro inesperado ao realizar a saída: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}


function concluirLotePeloPainel(loteId) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) { // Tenta travar por 30s
        return { success: false, message: 'Não foi possível executar a ação. O sistema está ocupado. Tente novamente.' };
    }
    try {
        const lotesSheet = getSheet(SHEET_NAMES.LOTES_PRODUCAO);
        const lotesData = lotesSheet.getDataRange().getValues();
        const lotesHeaders = lotesData[0];
        // Encontra a linha 0-based
        const loteRowIndex_0 = lotesData.findIndex(row => row[lotesHeaders.indexOf('ID_Lote')] === loteId);

        // Verifica se o lote existe e está pendente
        if (loteRowIndex_0 === -1 || lotesData[loteRowIndex_0][lotesHeaders.indexOf('Status')] !== 'Aguardando Produção') {
            const currentStatus = lotesData[loteRowIndex_0] ? lotesData[loteRowIndex_0][lotesHeaders.indexOf('Status')] : 'Não encontrado';
            return { success: false, message: `Lote "${loteId}" não encontrado ou o status atual é "${currentStatus}".` };
        }
        const loteRowIndex_1 = loteRowIndex_0 + 1; // 1-based index

        // Busca os itens do lote
        const itensLoteSheet = getSheet(SHEET_NAMES.ITENS_LOTE);
        const itensLoteHeaders = getHeaders(SHEET_NAMES.ITENS_LOTE);
        const itensDoLote = itensLoteSheet.getDataRange().getValues().slice(1).filter(row => row[itensLoteHeaders.indexOf('ID_Lote')] === loteId);

        // Se o lote estiver vazio, apenas marca como concluído
        if (itensDoLote.length === 0) {
             lotesSheet.getRange(loteRowIndex_1, lotesHeaders.indexOf('Status') + 1).setValue('Concluído (Vazio)');
             lotesSheet.getRange(loteRowIndex_1, lotesHeaders.indexOf('Data_Conclusao') + 1).setValue(new Date());
             SpreadsheetApp.flush();
             return { success: true, message: `Lote ${loteId} (vazio) marcado como concluído.` };
        }

        // --- Lógica para adicionar/atualizar Estoque Pronto ---
        const estoqueProntoSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
        const estoqueProntoData = estoqueProntoSheet.getDataRange().getValues(); // Lê com header
        const estoqueProntoHeaders = estoqueProntoData[0];
        
        // Cria mapa para acesso rápido
        const estoqueProntoMap = new Map();
        const epModeloIdx = estoqueProntoHeaders.indexOf('Modelo');
        const epGeneroIdx = estoqueProntoHeaders.indexOf('Gênero');
        const epCorIdx = estoqueProntoHeaders.indexOf('Cor');
        const epCorMangaIdx = estoqueProntoHeaders.indexOf('Cor_Manga');
        const epTamanhoIdx = estoqueProntoHeaders.indexOf('Tamanho');
        const epEstampaIdx = estoqueProntoHeaders.indexOf('Estampa');
        const epQtdIdx = estoqueProntoHeaders.indexOf('Quantidade');
        const epPrecoIdx = estoqueProntoHeaders.indexOf('Preço');
        const epStatusIdx = estoqueProntoHeaders.indexOf('Status');
        const epDataAttIdx = estoqueProntoHeaders.indexOf('Data_Atualização');

        estoqueProntoData.slice(1).forEach((row, index) => { // Pula header
             const nomeCompleto = [
                row[epModeloIdx], row[epGeneroIdx], row[epCorIdx],
                row[epCorMangaIdx] ? 'Manga ' + row[epCorMangaIdx] : '',
                row[epTamanhoIdx], row[epEstampaIdx]
             ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
             
             estoqueProntoMap.set(nomeCompleto, { 
                rowIndex: index + 2, // Linha real
                cellQtd: estoqueProntoSheet.getRange(index + 2, epQtdIdx + 1)
             });
        });
        
        const KEYWORDS = loadKeywords_(); // Carrega keywords locais
        const updatesEstoqueProntoRanges = [];
        const updatesEstoqueProntoValues = [];
        const novasLinhasEstoquePronto = [];

        for (const item of itensDoLote) {
            const produtoFinalCompleto = item[itensLoteHeaders.indexOf('Produto_Final_Completo')];
            const quantidadeProduzida = parseInt(item[itensLoteHeaders.indexOf('Quantidade')], 10) || 0;
            
            if (!produtoFinalCompleto || quantidadeProduzida <= 0) continue; // Pula item inválido

            const produtoLower = produtoFinalCompleto.toLowerCase();
            const estoqueInfo = estoqueProntoMap.get(produtoLower);
            
            if (estoqueInfo) {
                // Produto já existe, agenda atualização da quantidade
                updatesEstoqueProntoRanges.push(estoqueInfo.cellQtd.getA1Notation());
                updatesEstoqueProntoValues.push((estoqueInfo.cellQtd.getValue() || 0) + quantidadeProduzida);
                // Agenda atualização da data
                if (epDataAttIdx > -1) {
                  updatesEstoqueProntoRanges.push(estoqueProntoSheet.getRange(estoqueInfo.rowIndex, epDataAttIdx + 1).getA1Notation());
                  updatesEstoqueProntoValues.push(new Date());
                }
                 // Agenda atualização do status para 'disponivel'
                if (epStatusIdx > -1) {
                  updatesEstoqueProntoRanges.push(estoqueProntoSheet.getRange(estoqueInfo.rowIndex, epStatusIdx + 1).getA1Notation());
                  updatesEstoqueProntoValues.push('disponivel');
                }

            } else {
                // Produto não existe, prepara nova linha
                const detalhes = extractProductDetails_(produtoFinalCompleto, KEYWORDS);
                // Tenta encontrar um preço similar ou usa um padrão
                const precoPadrao = 99.9; // Defina um preço padrão razoável
                let precoEncontrado = precoPadrao;
                // Procura por um item com mesmo Modelo/Tamanho/Estampa para pegar o preço
                 const similar = estoqueProntoData.slice(1).find(row => 
                    row[epModeloIdx] === detalhes.modelo &&
                    row[epTamanhoIdx] === detalhes.tamanho &&
                    row[epEstampaIdx] === detalhes.estampa
                );
                if (similar && similar[epPrecoIdx]) precoEncontrado = parseFloat(String(similar[epPrecoIdx]).replace(",", ".")) || precoPadrao;
                
                const novaLinha = new Array(estoqueProntoHeaders.length).fill('');
                // Gera ID único
                const novoId = `PROD-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss')}-${Math.floor(Math.random()*100)}`;
                const idIndex = estoqueProntoHeaders.indexOf('ID'); 
                novaLinha[idIndex > -1 ? idIndex : 0] = novoId; // Usa coluna ID se existir
                
                // Preenche colunas com base nos detalhes
                novaLinha[epModeloIdx] = detalhes.modelo || '';
                novaLinha[epGeneroIdx] = detalhes.genero || '';
                novaLinha[epCorIdx] = detalhes.cor || '';
                novaLinha[epCorMangaIdx] = detalhes.cor_manga || '';
                novaLinha[epTamanhoIdx] = detalhes.tamanho || '';
                novaLinha[epEstampaIdx] = detalhes.estampa || '';
                novaLinha[epQtdIdx] = quantidadeProduzida;
                novaLinha[epPrecoIdx] = precoEncontrado; 
                if (epDataAttIdx > -1) novaLinha[epDataAttIdx] = new Date();
                if (epStatusIdx > -1) novaLinha[epStatusIdx] = 'disponivel'; // Status inicial
                // Adicionar Custo_Unitario se existir (pode buscar da matéria prima?)
                
                novasLinhasEstoquePronto.push(novaLinha);
            }
        }

        // Aplica atualizações e adiciona novas linhas em lote
        if(updatesEstoqueProntoRanges.length > 0) {
            const sheet = estoqueProntoSheet.getParent().getSheetByName(SHEET_NAMES.ESTOQUE_PRONTO); // Pega a aba de novo
            updatesEstoqueProntoRanges.forEach((rangeA1, index) => {
               sheet.getRange(rangeA1).setValue(updatesEstoqueProntoValues[index]);
            });
        }
        if (novasLinhasEstoquePronto.length > 0) {
            estoqueProntoSheet.getRange(estoqueProntoSheet.getLastRow() + 1, 1, novasLinhasEstoquePronto.length, novasLinhasEstoquePronto[0].length)
                             .setValues(novasLinhasEstoquePronto);
        }

        // --- Atualiza Status e Data de Conclusão do Lote ---
        const statusColIndex = lotesHeaders.indexOf('Status');
        const dataConclusaoColIndex = lotesHeaders.indexOf('Data_Conclusao');
        // Atualiza em lote (mesmo sendo uma linha só)
        const loteUpdateRange = lotesSheet.getRange(loteRowIndex_1, statusColIndex + 1, 1, dataConclusaoColIndex - statusColIndex + 1);
        const loteUpdateValues = [new Array(dataConclusaoColIndex - statusColIndex + 1).fill('')];
        loteUpdateValues[0][0] = 'Concluído'; // Índice 0 relativo ao range (Status)
        loteUpdateValues[0][dataConclusaoColIndex - statusColIndex] = new Date(); // Índice relativo (Data_Conclusao)
        loteUpdateRange.setValues(loteUpdateValues);
            
        SpreadsheetApp.flush(); // Garante a escrita antes de retornar

        return { success: true, message: `Lote ${loteId} concluído com sucesso! Estoque pronto atualizado.` };
        
    } catch(err) {
        Logger.log(`Erro crítico ao concluir lote ${loteId}: ${err.stack}`);
        // Tenta reverter o status se falhar no meio? (Complexo)
        // Por segurança, retorna erro claro.
        return { success: false, message: `Erro crítico ao concluir lote: ${err.message}. Verifique os logs e o estado das planilhas.` };
    } finally {
        lock.releaseLock();
    }
}


// Função de cancelar lote (já existente e ok)
function cancelarLotePeloPainel(loteId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Não foi possível executar a ação. O sistema está ocupado. Tente novamente.' };
  }
  try {
    if (!loteId) {
      return { success: false, message: 'ID do lote não fornecido.' };
    }
    
    const lotesSheet = getSheet(SHEET_NAMES.LOTES_PRODUCAO);
    const lotesData = lotesSheet.getDataRange().getValues();
    const lotesHeaders = lotesData[0];
    let loteRowIndex = -1;
    let loteRowData = null;

    for (let i = 1; i < lotesData.length; i++) {
      if (lotesData[i][lotesHeaders.indexOf('ID_Lote')] === loteId) {
        loteRowData = lotesData[i];
        if (String(loteRowData[lotesHeaders.indexOf('Status')]).trim() === 'Aguardando Produção') {
          loteRowIndex = i + 1; // 1-based index
        } else {
          // Permite cancelar mesmo se já cancelado (idempotente)
          if (String(loteRowData[lotesHeaders.indexOf('Status')]).trim() === 'Cancelado') {
             return { success: true, message: `Lote ${loteId} já estava cancelado.` };
          }
          return { success: false, message: `Não é possível cancelar o lote "${loteId}", pois o seu status é "${loteRowData[lotesHeaders.indexOf('Status')]}".` };
        }
        break;
      }
    }

    if (loteRowIndex === -1) {
       return { success: false, message: `Lote "${loteId}" não encontrado com status 'Aguardando Produção'.` };
    }

    // --- Devolve Matéria Prima ---
    const itensLoteSheet = getSheet(SHEET_NAMES.ITENS_LOTE);
    const itensLoteHeaders = itensLoteSheet.getRange(1, 1, 1, itensLoteSheet.getLastColumn()).getValues()[0];
    const itensDoLote = itensLoteSheet.getDataRange().getValues().slice(1).filter(row => row[itensLoteHeaders.indexOf('ID_Lote')] === loteId);
    
    if (itensDoLote.length > 0) {
        const materiaPrimaSheet = getSheet(SHEET_NAMES.ESTOQUE_MATERIA);
        const materiaData = materiaPrimaSheet.getDataRange().getValues(); // com header
        const materiaHeaders = materiaData[0];
        // Cria mapa para acesso rápido
        const materiaMap = new Map(); 
        const qtdReservadaIdx = materiaHeaders.indexOf('Qtd_Reservada');
        materiaData.slice(1).forEach((row, index) => {
            let nomeMateriaCompleto = [
                row[materiaHeaders.indexOf('Modelo')], row[materiaHeaders.indexOf('Gênero')], row[materiaHeaders.indexOf('Cor')], 
                row[materiaHeaders.indexOf('Cor_Manga')] ? 'Manga ' + row[materiaHeaders.indexOf('Cor_Manga')] : '', 
                row[materiaHeaders.indexOf('Tamanho')]
             ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
            materiaMap.set(nomeMateriaCompleto, { 
              rowIndex: index + 2, 
              cellReservada: materiaPrimaSheet.getRange(index + 2, qtdReservadaIdx + 1)
            });
        });

        const KEYWORDS = loadKeywords_();
        const updatesReservasRanges = [];
        const updatesReservasValues = [];

        for (const item of itensDoLote) {
            const produtoFinalCompleto = item[itensLoteHeaders.indexOf('Produto_Final_Completo')];
            const quantidadeReservada = parseInt(item[itensLoteHeaders.indexOf('Quantidade')], 10) || 0;
            
            if (quantidadeReservada <= 0) continue; // Pula se quantidade for 0

            const detalhesItem = extractProductDetails_(produtoFinalCompleto, KEYWORDS);
            const materiaPrimaNecessaria = [
                detalhesItem.modelo, detalhesItem.genero, detalhesItem.cor, 
                detalhesItem.cor_manga ? 'Manga ' + detalhesItem.cor_manga : '', 
                detalhesItem.tamanho
            ].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
            
            const materiaInfo = materiaMap.get(materiaPrimaNecessaria);
            if (materiaInfo) {
                updatesReservasRanges.push(materiaInfo.cellReservada.getA1Notation());
                updatesReservasValues.push(Math.max(0, (materiaInfo.cellReservada.getValue() || 0) - quantidadeReservada));
            } else {
                Logger.log(`AVISO (Cancelamento): Matéria-prima "${materiaPrimaNecessaria}" do lote ${loteId} não foi encontrada para devolver ao estoque.`);
            }
        }
        // Aplica as atualizações de reserva em lote
        if(updatesReservasRanges.length > 0) {
            const sheet = materiaPrimaSheet.getParent().getSheetByName(SHEET_NAMES.ESTOQUE_MATERIA); // Pega a aba de novo
            updatesReservasRanges.forEach((rangeA1, index) => {
               sheet.getRange(rangeA1).setValue(updatesReservasValues[index]);
            });
            SpreadsheetApp.flush(); // Garante escrita
        }
    }
    
    // --- Atualiza Status do Lote ---
    lotesSheet.getRange(loteRowIndex, lotesHeaders.indexOf('Status') + 1).setValue('Cancelado');
    // Opcional: Limpar Data_Conclusao se existir
    const dataConclusaoIdx = lotesHeaders.indexOf('Data_Conclusao');
    if (dataConclusaoIdx > -1) {
       lotesSheet.getRange(loteRowIndex, dataConclusaoIdx + 1).setValue(''); 
    }
    SpreadsheetApp.flush();

    return { success: true, message: `Lote ${loteId} cancelado com sucesso! A matéria-prima foi devolvida ao estoque.` };

  } catch(e) {
    Logger.log(`Erro em cancelarLotePeloPainel: ${e.stack}`);
    return { success: false, message: `Ocorreu um erro ao cancelar o lote: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}



function registrarMovimentoCaixaPeloPainel(movimentoInfo) {
    try {
        const { tipo, valor, descricao, responsavel } = movimentoInfo;
         if (!tipo || !valor || !descricao || !responsavel) {
            return { success: false, message: 'Tipo, Valor, Descrição e Responsável são obrigatórios.' };
        }
        const valorNum = parseFloat(valor);
        if (isNaN(valorNum) || valorNum <= 0) {
            return { success: false, message: 'O valor deve ser um número positivo.' };
        }
        
        const fluxoSheet = getSheet(SHEET_NAMES.FLUXO_CAIXA);
        const headers = getHeaders(SHEET_NAMES.FLUXO_CAIXA); 
        const newRow = new Array(headers.length).fill(''); 
        
        // Mapeia os dados para as colunas corretas pelos cabeçalhos
        newRow[headers.indexOf('Data')] = new Date();
        newRow[headers.indexOf('Tipo')] = tipo;
        // Usa 'Descricao' ou 'Descrição'
        const descIndex = headers.indexOf('Descricao') !== -1 ? headers.indexOf('Descricao') : headers.indexOf('Descrição');
        if(descIndex > -1) newRow[descIndex] = descricao; 

        newRow[headers.indexOf('Entrada')] = tipo === 'Entrada' ? valorNum : '';
        // Usa 'Saida' ou 'Saída'
        const saidaIndex = headers.indexOf('Saida') !== -1 ? headers.indexOf('Saida') : headers.indexOf('Saída');
        if(saidaIndex > -1) newRow[saidaIndex] = tipo === 'Saída' ? valorNum : ''; 
        
        // Opcional: Categoria, Evento_Relacionado
        const catIndex = headers.indexOf('Categoria');
        if(catIndex > -1) newRow[catIndex] = ''; // Pode ser preenchido se necessário
        const eventoIndex = headers.indexOf('Evento_Relacionado');
         if(eventoIndex > -1) newRow[eventoIndex] = ''; 

        // Usa 'Responsavel' ou 'Responsável'
        const respIndex = headers.indexOf('Responsavel') !== -1 ? headers.indexOf('Responsavel') : headers.indexOf('Responsável');
        if(respIndex > -1) newRow[respIndex] = responsavel; 
        
        fluxoSheet.appendRow(newRow);
        SpreadsheetApp.flush(); // Garante a escrita
        return { success: true, message: `${tipo} registrada com sucesso.` };
  } catch (e) {
        Logger.log(`Erro em registrarMovimentoCaixaPeloPainel: ${e.stack}`);
        return { success: false, message: `Erro ao registrar no caixa: ${e.message}` };
    }
}


function adicionarMateriaPrimaPeloPainel(materiaInfo) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
        return { success: false, message: 'Sistema ocupado, tente novamente.' };
    }
    try {
        const { modelo, genero, cor, corManga, tamanho, quantidade } = materiaInfo;
        const qtdNum = parseInt(quantidade, 10); // Converte para número

      	if (!modelo || !cor || !tamanho || (isNaN(qtdNum) || qtdNum < 0)) { 
            return { success: false, message: 'Modelo, Cor, Tamanho e Quantidade (>=0) são obrigatórios.' };
        }

        const materiaSheet = getSheet(SHEET_NAMES.ESTOQUE_MATERIA);
        const headers = getHeaders(SHEET_NAMES.ESTOQUE_MATERIA);
        const data = materiaSheet.getDataRange().getValues(); // Lê tudo incluindo header

        const modeloIndex = headers.indexOf('Modelo');
        const generoIndex = headers.indexOf('Gênero');
        const corIndex = headers.indexOf('Cor');
        const corMangaIndex = headers.indexOf('Cor_Manga');
        const tamanhoIndex = headers.indexOf('Tamanho');
        const qtdAtualIndex = headers.indexOf('Qtd_Atual');
        const dataMovIndex = headers.indexOf('Data_Ultima_Movimentacao'); // Opcional

        // Normaliza entradas para comparação
        const inputModelo = (modelo || '').trim();
        const inputGenero = (genero || '').trim();
        const inputCor = (cor || '').trim();
        const inputCorManga = (corManga || '').trim();
        const inputTamanho = (tamanho || '').trim();

      	let rowIndexToUpdate = -1;

      	// Procura por correspondência exata
      	for (let i = 1; i < data.length; i++) { // Começa do 1 para pular header
          	const row = data[i];
          	if ((row[modeloIndex] || '').trim() === inputModelo &&
              	(row[generoIndex] || '').trim() === inputGenero &&
              	(row[corIndex] || '').trim() === inputCor &&
              	(row[corMangaIndex] || '').trim() === inputCorManga &&
              	(row[tamanhoIndex] || '').trim() === inputTamanho)
          	{
              	rowIndexToUpdate = i + 1; // 1-based index da linha na planilha
              	break;
          	}
      	}

      	if (rowIndexToUpdate !== -1) {
          	// --- Atualiza item existente ---
          	const qtdAtualCell = materiaSheet.getRange(rowIndexToUpdate, qtdAtualIndex + 1);
          	const qtdAtual = parseInt(qtdAtualCell.getValue(), 10) || 0;
          	const novaQtd = qtdAtual + qtdNum;
          	
          	// Atualiza quantidade e data da movimentação
          	qtdAtualCell.setValue(novaQtd);
          	if (dataMovIndex > -1) {
            		materiaSheet.getRange(rowIndexToUpdate, dataMovIndex + 1).setValue(new Date());
          	}
          	SpreadsheetApp.flush(); // Garante escrita

          	return { success: true, message: `Estoque de matéria-prima atualizado! Novo total: ${novaQtd}.` };
        
      	} else {
          	// --- Adiciona novo item ---
          	const novaLinha = new Array(headers.length).fill('');
          	novaLinha[modeloIndex] = inputModelo;
          	novaLinha[generoIndex] = inputGenero;
          	novaLinha[corIndex] = inputCor;
          	novaLinha[corMangaIndex] = inputCorManga;
          	novaLinha[tamanhoIndex] = inputTamanho;
          	novaLinha[qtdAtualIndex] = qtdNum;
          	
          	// Define padrões para outras colunas importantes
          	const qtdReservadaIdx = headers.indexOf('Qtd_Reservada');
          	if (qtdReservadaIdx > -1) novaLinha[qtdReservadaIdx] = 0;
          	
          	const dataEntradaIdx = headers.indexOf('Data_Entrada');
          	if (dataEntradaIdx > -1) novaLinha[dataEntradaIdx] = new Date();
          	
          	const fornecedorIdx = headers.indexOf('Fornecedor');
          	if (fornecedorIdx > -1) novaLinha[fornecedorIdx] = 'Painel Admin'; // Ou um campo para isso
          	
          	const statusIndex = headers.indexOf('Status');
          	if (statusIndex > -1) novaLinha[statusIndex] = qtdNum > 0 ? 'Disponível' : 'Zerado'; // Status inicial
            	
          	const locIndex = headers.indexOf('Localização');
          	if (locIndex > -1) novaLinha[locIndex] = 'Estoque A'; // Localização Padrão
            	
          	if (dataMovIndex > -1) novaLinha[dataMovIndex] = new Date(); // Data da última movimentação

          	materiaSheet.appendRow(novaLinha);
          	SpreadsheetApp.flush(); // Garante escrita
          	return { success: true, message: 'Nova matéria-prima adicionada com sucesso!' };
      	}

  	} catch (e) {
      	Logger.log(`Erro em adicionarMateriaPrimaPeloPainel: ${e.stack}`);
      	return { success: false, message: `Ocorreu um erro ao adicionar matéria-prima: ${e.message}` };
  	} finally {
      	lock.releaseLock();
  	}
}


function adicionarProdutoProntoPeloPainel(produtoInfo) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
         return { success: false, message: 'Sistema ocupado, tente novamente.' };
    }
    try {
        const { modelo, genero, cor, corManga, tamanho, estampa, quantidade, preco } = produtoInfo;
        const qtdNum = parseInt(quantidade, 10);
        const precoNum = parseFloat(preco);

        // Validação mais robusta
      	if (!modelo || !cor || !tamanho || !estampa || (isNaN(qtdNum) || qtdNum < 0) || (isNaN(precoNum) || precoNum <= 0) ) { 
            return { success: false, message: 'Modelo, Cor, Tamanho, Estampa, Quantidade (>=0) e Preço (>0) são obrigatórios.' };
        }

        // Constrói nome consistentemente para busca
        const nomeCompleto = [modelo, genero, cor, corManga ? 'Manga ' + corManga : '', tamanho, estampa]
            .filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
        const nomeCompletoLower = nomeCompleto.toLowerCase();

        const estoqueSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
        const headers = getHeaders(SHEET_NAMES.ESTOQUE_PRONTO);
        const data = estoqueSheet.getDataRange().getValues(); // Lê com header

        // Indices das colunas
        const modeloIndex = headers.indexOf('Modelo');
        const generoIndex = headers.indexOf('Gênero');
        const corIndex = headers.indexOf('Cor');
        const corMangaIndex = headers.indexOf('Cor_Manga');
        const tamanhoIndex = headers.indexOf('Tamanho');
        const estampaIndex = headers.indexOf('Estampa');
      	const qtdIndex = headers.indexOf('Quantidade');
      	const precoIndex = headers.indexOf('Preço');
      	const dataAtualizacaoIndex = headers.indexOf('Data_Atualização');
      	const statusIndex = headers.indexOf('Status'); // Para atualizar status

      	let rowIndexToUpdate = -1;

      	// Procura por correspondência exata
      	for (let i = 1; i < data.length; i++) { // Pula header
          	const row = data[i];
          	const nomeNaPlanilha = [
              	row[modeloIndex], row[generoIndex], row[corIndex],
              	row[corMangaIndex] ? 'Manga ' + row[corMangaIndex] : '',
              	row[tamanhoIndex], row[estampaIndex]
          	].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();

          	if (nomeNaPlanilha === nomeCompletoLower) {
              	rowIndexToUpdate = i + 1; // Linha real na planilha
              	break;
          	}
      	}

      	if (rowIndexToUpdate !== -1) {
          	// --- Atualiza item existente ---
          	const qtdAtualCell = estoqueSheet.getRange(rowIndexToUpdate, qtdIndex + 1);
          	const qtdAtual = parseInt(qtdAtualCell.getValue(), 10) || 0;
          	const novaQtd = qtdAtual + qtdNum;
          	
          	// Atualiza quantidade, preço, data e status
          	qtdAtualCell.setValue(novaQtd);
          	estoqueSheet.getRange(rowIndexToUpdate, precoIndex + 1).setValue(precoNum); // Atualiza preço
          	if (dataAtualizacaoIndex > -1) {
            		estoqueSheet.getRange(rowIndexToUpdate, dataAtualizacaoIndex + 1).setValue(new Date());
          	}
          	if (statusIndex > -1 && novaQtd > 0) {
            	 	estoqueSheet.getRange(rowIndexToUpdate, statusIndex + 1).setValue('disponivel');
          	} else if (statusIndex > -1 && novaQtd <= 0) {
              	estoqueSheet.getRange(rowIndexToUpdate, statusIndex + 1).setValue('esgotado');
          	}
          	SpreadsheetApp.flush(); // Garante escrita

          	return { success: true, message: `Estoque de "${nomeCompleto}" atualizado. Nova quantidade: ${novaQtd}.` };
      	} else {
          	// --- Adiciona novo item ---
          	const novoId = `PROD-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss')}-${Math.floor(Math.random()*100)}`;
          	
          	const newRow = new Array(headers.length).fill('');
          	
          	const idIndex = headers.indexOf('ID');
          	newRow[idIndex > -1 ? idIndex : 0] = novoId; // Usa coluna ID se existir
          	
          	newRow[modeloIndex] = modelo;
          	newRow[generoIndex] = genero || '';
          	newRow[corIndex] = cor;
          	newRow[corMangaIndex] = corManga || '';
          	newRow[tamanhoIndex] = tamanho;
  	        newRow[estampaIndex] = estampa;
  	        newRow[qtdIndex] = qtdNum;
  	        newRow[precoIndex] = precoNum;
  	        if (dataAtualizacaoIndex > -1) newRow[dataAtualizacaoIndex] = new Date();
  	        if (statusIndex > -1) newRow[statusIndex] = qtdNum > 0 ? 'disponivel' : 'esgotado'; // Status inicial
  	        // Adicionar Custo_Unitario se existir (pode buscar da matéria prima?)

  	        estoqueSheet.appendRow(newRow);
  	        SpreadsheetApp.flush(); // Garante escrita
  	        return { success: true, message: `Novo produto "${nomeCompleto}" adicionado com ${qtdNum} unidade(s)!` };
  	   	}

  	} catch(e) {
  	   	Logger.log(`Erro em adicionarProdutoProntoPeloPainel: ${e.stack}`);
  	   	return { success: false, message: `Erro ao adicionar produto pronto: ${e.message}` };
  	} finally {
  	   	lock.releaseLock();
  	}
}


function adicionarRevendedorPeloPainel(nomeRevendedor) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
       return { success: false, message: 'Sistema ocupado, tente novamente.' };
    }
  	try {
      	const nomeLimpo = (nomeRevendedor || "").trim();
      	if (!nomeLimpo) {
          	return { success: false, message: "O nome do vendedor/revendedor não pode estar vazio." };
      	}

      	const revendedorSheet = getSheet(SHEET_NAMES.VENDEDORES);
      	const headers = getHeaders(SHEET_NAMES.VENDEDORES);
      	const data = readSheetData(SHEET_NAMES.VENDEDORES); // Lê sem header
  	   	const nomeColIndex = headers.indexOf('Nome');

  	   	if (nomeColIndex === -1) {
          		throw new Error("A coluna 'Nome' não foi encontrada na aba Vendedores_Revendedores.");
      	}

      	const nomeNormalizado = nomeLimpo.toLowerCase();
      	// Verifica se já existe (case-insensitive)
      	const jaExiste = data.some(row => (row[nomeColIndex] || "").toString().trim().toLowerCase() === nomeNormalizado);

  	   	if (jaExiste) {
          	return { success: false, message: `O vendedor/revendedor "${nomeLimpo}" já está cadastrado.` };
      	}
      	
      	// Cria nova linha com valores padrão
      	const newRow = new Array(headers.length).fill('');
      	newRow[nomeColIndex] = nomeLimpo; // Nome como fornecido
      	const statusColIndex = headers.indexOf('Status');
      	if (statusColIndex > -1) newRow[statusColIndex] = 'ativo'; // Padrão 'ativo'
      	
      	// Adiciona comissão padrão se a coluna existir
      	const comissaoColIndex = headers.indexOf('Comissao') !== -1 ? headers.indexOf('Comissao') : headers.indexOf('Comissão');
      	if (comissaoColIndex > -1) newRow[comissaoColIndex] = 0.4; // Comissão padrão 40%
      	
      	// Adiciona permissão padrão se a coluna existir
      	const permissoesColIndex = headers.indexOf('Permissões');
      	if (permissoesColIndex > -1) newRow[permissoesColIndex] = 'revendedor'; // Permissão padrão

      	revendedorSheet.appendRow(newRow);
      	SpreadsheetApp.flush(); // Garante escrita
      	
      	return { success: true, message: `Vendedor/Revendedor "${nomeLimpo}" adicionado com sucesso!` };

  	} catch(e) {
      	Logger.log(`Erro ao adicionar revendedor: ${e.stack}`);
      	return { success: false, message: `Erro ao adicionar revendedor: ${e.message}` };
  	} finally {
      	lock.releaseLock();
  	}
}


// =================================================================
// (NOVA FUNÇÃO ADICIONADA)
// Função para lançar Venda ou Retorno de Consignação pelo painel
// =================================================================

function registrarLancamentoConsignado(lancamentoInfo) {
    const { tipo, revendedor, produto, quantidade } = lancamentoInfo;
    
    if (!tipo || !revendedor || !produto || !quantidade || quantidade <= 0) {
        return { success: false, message: 'Tipo, Revendedor, Produto e Quantidade (>0) são obrigatórios.' };
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
        return { success: false, message: 'Sistema ocupado, tente novamente.' };
    }
    
    try {
      	const consignacaoSheet = getSheet(SHEET_NAMES.ESTOQUE_CONSIGNACAO);
      	const consignacaoData = consignacaoSheet.getDataRange().getValues(); // Lê com header
      	const headers = consignacaoData[0];

      	// Mapeia colunas
    	   	const revendedorCol = headers.indexOf('Revendedor');
  	   	const produtoCol = headers.indexOf('Produto');
  	   	const qtdRestanteCol = headers.indexOf('Qtd_Restante');
  	   	const qtdVendidaCol = headers.indexOf('Qtd_Vendida');
  	   	const qtdRetornadaCol = headers.indexOf('Qtd_Retornada');
  	   	const precoVendaCol = headers.indexOf('Preco_Venda');

  	   	if (qtdRestanteCol === -1 || qtdVendidaCol === -1 || qtdRetornadaCol === -1) {
  	       	throw new Error("Colunas essenciais (Qtd_Restante, Qtd_Vendida, Qtd_Retornada) não encontradas em ESTOQUE_CONSIGNACAO.");
  	   	}

  	   	// Normaliza nomes para busca
  	   	const revendedorLower = revendedor.toLowerCase().trim();
  	   	const produtoLower = produto.toLowerCase().trim();
  	   	let rowIndexToUpdate = -1;
  	   	let rowData = null;
  	   	let totalRestanteAgregado = 0; // O total de estoque em TODAS as linhas
      	let produtoFoiEncontrado = false; // Flag para saber se o produto existe

      	// Procura pelo item
      	for (let i = 1; i < consignacaoData.length; i++) {
          	const row = consignacaoData[i];
          	const revendedorPlanilha = (row[revendedorCol] || '').toLowerCase().trim();
          	const produtoPlanilha = (row[produtoCol] || '').toLowerCase().trim();
          	
          	if (revendedorPlanilha === revendedorLower && produtoPlanilha === produtoLower) {
              	produtoFoiEncontrado = true; // Marcamos que o produto existe
              	const qtdRestanteNestaLinha = parseInt(row[qtdRestanteCol], 10) || 0;
              	totalRestanteAgregado += qtdRestanteNestaLinha; // Soma o total
              	
              	// Se esta linha TEM estoque E AINDA não escolhemos uma linha para dar baixa...
              	if (qtdRestanteNestaLinha >= quantidade && rowIndexToUpdate === -1) { 
                  	rowIndexToUpdate = i + 1; // Linha real na planilha
                  	rowData = row; // Guarda os dados desta linha específica
                  	// Não damos 'break' para continuar a somar o 'totalRestanteAgregado'
              	}
          	}
      	}

      	// --- Nova Lógica de Erro ---
      	if (rowIndexToUpdate === -1) { // Nenhuma linha individual tinha estoque suficiente
          	if (!produtoFoiEncontrado) {
              	// Erro 1: Produto nunca existiu para este revendedor
              	return { success: false, message: `Produto "${produto}" não encontrado no estoque consignado de ${revendedor}.` };
          	} else {
              	// Erro 2: Produto existe, mas o estoque é insuficiente
              	// A mensagem de erro agora mostra o estoque TOTAL (soma de todas as linhas)
              	return { success: false, message: `Estoque consignado insuficiente para "${produto}". Pedido: ${quantidade}, Disponível (total): ${totalRestanteAgregado} un.` };
          	}
      	}

      	// --- Lógica de Venda Consignada ---
      	if (tipo === 'Venda Consignada') {
          	// Atualiza Qtd_Vendida e Qtd_Restante
          	const qtdVendidaAtual = parseInt(rowData[qtdVendidaCol], 10) || 0;
          	const qtdRestanteAtual = parseInt(rowData[qtdRestanteCol], 10) || 0;
          	const novaQtdVendida = qtdVendidaAtual + quantidade;
          	const novaQtdRestante = qtdRestanteAtual - quantidade;

          	consignacaoSheet.getRange(rowIndexToUpdate, qtdVendidaCol + 1).setValue(novaQtdVendida);
          	consignacaoSheet.getRange(rowIndexToUpdate, qtdRestanteCol + 1).setValue(novaQtdRestante);
          	
          	// Registra a venda no VENDAS_LOG
          	const vendasLogSheet = getSheet(SHEET_NAMES.VENDAS_LOG);
          	const logHeaders = getHeaders(SHEET_NAMES.VENDAS_LOG);
          	const newLogRow = new Array(logHeaders.length).fill('');
          	
          	const precoVenda = parseFloat(String(rowData[precoVendaCol]).replace(",", ".")) || 0;
          	const valorTotalVenda = precoVenda * quantidade;
          	// Custo e Lucro podem não ser calculados aqui, mas sim no acerto
          	
          	newLogRow[logHeaders.indexOf('ID')] = `VENDA-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss')}`;
          	newLogRow[logHeaders.indexOf('Data_Hora')] = new Date();
          	newLogRow[logHeaders.indexOf('Vendedor')] = revendedor; // O revendedor é o vendedor
          	newLogRow[logHeaders.indexOf('Produto_Completo')] = produto.toUpperCase();
          	newLogRow[logHeaders.indexOf('Quantidade')] = quantidade;
          	newLogRow[logHeaders.indexOf('Tipo_Venda')] = 'consignado';
          	newLogRow[logHeaders.indexOf('Canal')] = 'Painel Admin';
          	newLogRow[logHeaders.indexOf('Valor')] = parseFloat(valorTotalVenda.toFixed(2));
          	newLogRow[logHeaders.indexOf('Status')] = 'concluída'; // A venda está concluída (pendente de acerto)

          	vendasLogSheet.appendRow(newLogRow);
          	SpreadsheetApp.flush(); // Garante a escrita

          	return { success: true, message: `Venda consignada de ${quantidade}x "${produto}" registrada para ${revendedor}.` };
      	}
      	
      	// --- Lógica de Retorno Consignado ---
      	if (tipo === 'Retorno Consignado') {
          	// Atualiza Qtd_Retornada e Qtd_Restante
          	const qtdRetornadaAtual = parseInt(rowData[qtdRetornadaCol], 10) || 0;
          	const qtdRestanteAtual = parseInt(rowData[qtdRestanteCol], 10) || 0;
          	const novaQtdRetornada = qtdRetornadaAtual + quantidade;
          	const novaQtdRestante = qtdRestanteAtual - quantidade;

          	consignacaoSheet.getRange(rowIndexToUpdate, qtdRetornadaCol + 1).setValue(novaQtdRetornada);
          	consignacaoSheet.getRange(rowIndexToUpdate, qtdRestanteCol + 1).setValue(novaQtdRestante);
          	
          	// Devolve o item ao Estoque Pronto
          	const estoqueProntoSheet = getSheet(SHEET_NAMES.ESTOQUE_PRONTO);
          	const estoqueProntoData = estoqueProntoSheet.getDataRange().getValues();
          	const estoqueProntoHeaders = estoqueProntoData[0];
          	const epQtdIdx = estoqueProntoHeaders.indexOf('Quantidade');

          	let produtoProntoRowIndex = -1;
          	// Procura o produto no estoque pronto para devolver
          	for (let i = 1; i < estoqueProntoData.length; i++) {
              	const row = estoqueProntoData[i];
              	const nomeNaPlanilha = [
                  	row[estoqueProntoHeaders.indexOf('Modelo')], row[estoqueProntoHeaders.indexOf('Gênero')], row[estoqueProntoHeaders.indexOf('Cor')],
                  	row[estoqueProntoHeaders.indexOf('Cor_Manga')] ? 'Manga ' + row[estoqueProntoHeaders.indexOf('Cor_Manga')] : '',
                  	row[estoqueProntoHeaders.indexOf('Tamanho')], row[estoqueProntoHeaders.indexOf('Estampa')]
  	           	].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
  		        
              	if (nomeNaPlanilha === produtoLower) {
              		produtoProntoRowIndex = i + 1; // Linha real
                  	break;
              	}
          	}

          	if (produtoProntoRowIndex !== -1) {
              	const qtdAtualCell = estoqueProntoSheet.getRange(produtoProntoRowIndex, epQtdIdx + 1);
              	const qtdAtual = parseInt(qtdAtualCell.getValue(), 10) || 0;
              	qtdAtualCell.setValue(qtdAtual + quantidade);
          	} else {
              	// Se o produto não existir mais no estoque pronto (improvável), ele precisa ser recriado
              	Logger.log(`AVISO: Produto "${produto}" retornado da consignação não foi encontrado no Estoque Pronto. Adicionando como nova linha (sem preço/custo).`);
              	// Tentar recriar com dados mínimos
              	const KEYWORDS = loadKeywords_();
              	const detalhes = extractProductDetails_(produto, KEYWORDS);
              	const newRow = new Array(estoqueProntoHeaders.length).fill('');
              	newRow[estoqueProntoHeaders.indexOf('Modelo')] = detalhes.modelo || '';
              	newRow[estoqueProntoHeaders.indexOf('Gênero')] = detalhes.genero || '';
            	 	newRow[estoqueProntoHeaders.indexOf('Cor')] = detalhes.cor || '';
            	 	newRow[estoqueProntoHeaders.indexOf('Cor_Manga')] = detalhes.cor_manga || '';
              	newRow[estoqueProntoHeaders.indexOf('Tamanho')] = detalhes.tamanho || '';
            	 	newRow[estoqueProntoHeaders.indexOf('Estampa')] = detalhes.estampa || '';
            	 	newRow[epQtdIdx] = quantidade;
            	 	newRow[epPrecoIdx] = parseFloat(String(rowData[precoVendaCol]).replace(",", ".")) || 0; // Pega preço da consignação
      	       	newRow[estoqueProntoHeaders.indexOf('Status')] = 'disponivel';
            	 	newRow[estoqueProntoHeaders.indexOf('Data_Atualização')] = new Date();
            	 	estoqueProntoSheet.appendRow(newRow);
          	}

          	SpreadsheetApp.flush(); // Garante a escrita

          	return { success: true, message: `Retorno de ${quantidade}x "${produto}" registrado. Estoque principal atualizado.` };
      	}
      	
      	return { success: false, message: 'Tipo de lançamento desconhecido.' };

    } catch (e) {
        Logger.log(`Erro em registrarLancamentoConsignado: ${e.stack}`);
        return { success: false, message: `Erro ao processar lançamento: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}



// =================================================================
// FUNÇÕES HELPERS LOCAIS (Keywords, Extração de Detalhes)
// =================================================================

// Cache para Keywords (simples, válido por 10 minutos)
const CACHE_EXPIRATION = 600; // 10 minutos
const SCRIPT_CACHE = CacheService.getScriptCache();

function loadKeywords_() {
  const cacheKey = 'KEYWORDS_CACHE_DASHBOARD'; // Chave específica para o dashboard
  const cached = SCRIPT_CACHE.get(cacheKey);
  if (cached) {
    try { 
      // Logger.log("Usando keywords do cache.");
      return JSON.parse(cached);
    } catch(e) { 
      Logger.log("Erro ao parsear cache de keywords, recarregando.");
      /* Ignora cache inválido e recarrega */ 
    }
  }
  Logger.log("Carregando keywords da planilha...");
  const keywordSheet = getSheet(SHEET_NAMES.PALAVRAS_CHAVE);
  const data = keywordSheet.getDataRange().getValues().slice(1); // Pula header
  const keywords = {};
  data.forEach(row => {
    const palavraUsuario = String(row[0]).toLowerCase().trim();
    if (!palavraUsuario) return; // Pula linhas vazias
    const palavraSistema = String(row[1]).trim();
    const categoria = String(row[2]).trim();
    const prioridade = parseInt(row[3], 10) || 99; // Padrão 99 se não for número
    
    if (!palavraSistema || !categoria) return; // Pula linhas sem dados essenciais

    if (!keywords[categoria]) {
      keywords[categoria] = [];
    }
    keywords[categoria].push({ palavraUsuario, palavraSistema, prioridade });
  });

  // Ordena por prioridade e depois por comprimento (mais específico primeiro)
  for (const categoria in keywords) {
    keywords[categoria].sort((a, b) => (a.prioridade - b.prioridade) || (b.palavraUsuario.length - a.palavraUsuario.length));
  }
  
  try {
     SCRIPT_CACHE.put(cacheKey, JSON.stringify(keywords), CACHE_EXPIRATION);
     Logger.log("Keywords salvas no cache.");
  } catch (e) {
     Logger.log(`Erro ao salvar keywords no cache: ${e.message}`);
  }
  return keywords;
}

// Função de extração adaptada para usar aqui
function extractProductDetails_(text, keywords) {
    const lowerText = String(text || '').toLowerCase(); // Garante que é string
    // Inicializa todos os detalhes como null
    const details = { modelo: null, cor: null, tamanho: null, estampa: null, tipo_tecido: null, cor_manga: null, genero: null, quantidade: 1, forma_pagamento: null };
    let cleanText = ` ${lowerText} `; // Add espaços para limites de palavra (\b)

    // Ordem de extração: do mais específico/menos ambíguo para o mais geral
    const orderedCategories = ['tamanho', 'genero', 'cor_manga', 'cor', 'modelo', 'tipo_tecido', 'forma_pagamento'];

    for (const categoria of orderedCategories) {
        if (!keywords[categoria]) continue; // Pula se não houver keywords para a categoria
        
        // Itera sobre as keywords já ordenadas (prioridade, comprimento)
        for (const kw of keywords[categoria]) {
            // Cria regex para encontrar a palavra exata (limites \b)
          	// Escapa caracteres especiais na palavra do usuário se necessário
          	const escapedKw = kw.palavraUsuario.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    	       const regex = new RegExp(`\\b${escapedKw}\\b`, 'i'); 
          	 
          	if (regex.test(cleanText)) {
          	     details[categoria] = kw.palavraSistema; // Atribui o valor padrão
          	     // Remove a palavra encontrada do texto para evitar re-matches
          	     cleanText = cleanText.replace(regex, ' ').trim(); 
        	        cleanText = ` ${cleanText} `; // Re-adiciona espaços para próximos matches
        	       break; // Encontrou a melhor correspondência para esta categoria, vai para a próxima
        	   }
        }
         cleanText = cleanText.trim(); // Remove espaços extras entre categorias
    }
    
    // O que sobrou no cleanText é considerado a estampa
    cleanText = cleanText.trim();
    if (cleanText) {
  	   // Capitaliza cada palavra da estampa (ex: "key of death" -> "Key Of Death")
  	   details.estampa = cleanText.split(' ')
    	                      .filter(Boolean) // Remove palavras vazias se houver espaços múltiplos
   	                       .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    	                     	.join(' ');
    } else {
        details.estampa = null; // Garante que estampa seja null se nada sobrou
    }

    // Tenta extrair quantidade do início do texto original se não foi pego antes
  	const qtyMatchStart = String(text || '').trim().match(/^(\d+)\s+/);
  	if (qtyMatchStart) {
  	     details.quantidade = parseInt(qtyMatchStart[1], 10);
  }

    return details;
}

