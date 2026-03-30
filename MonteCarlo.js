/**
 * Cria um menu customizado na barra superior do Google Sheets.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎲 Simulação Monte Carlo')
      .addItem('1. Configurar Planilha Padrão', 'setupTemplate')
      .addItem('2. Rodar Simulação e Gerar Gráficos', 'runMonteCarlo')
      .addToUi();
}

/**
 * Prepara o layout da planilha aplicando larguras personalizadas, 
 * cores de destaque para títulos e inputs padronizados.
 */
function setupTemplate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  
  // Títulos e Cabeçalhos (Linha 1)
  sheet.getRange("A1").setValue("Histórico:");
  sheet.getRange("B1").setValue("Estimativa 3 Pontos");
  sheet.getRange("C1").setValue("Estimativa inicial");
  sheet.getRange("D1").setValue("Cenários gerados");
  sheet.getRange("E1").setValue("Unidade:");
  sheet.getRange("F1").setValue("R$");
  sheet.getRange("G1").setValue("Resultados");
  
  // Inputs de 3 Pontos e Rótulos
  sheet.getRange("B2").setValue("Mínimo:");
  sheet.getRange("C2").setValue("");
  sheet.getRange("B3").setValue("Mais Provável:");
  sheet.getRange("C3").setValue("");
  sheet.getRange("B4").setValue("Máximo:");
  sheet.getRange("C4").setValue("");
  
  // Percentis e Alvo
  sheet.getRange("F2").setValue("Percentil 10");
  sheet.getRange("F3").setValue("Percentil 50");
  sheet.getRange("F4").setValue("Percentil 90");
  sheet.getRange("E5").setValue("Valor alvo:");
  sheet.getRange("F5").setValue("");
  sheet.getRange("G5").setValue("Confiabilidade");
  
  // ==========================================
  // FORMATAÇÃO VISUAL E ESTÉTICA
  // ==========================================
  
  // Fonte 13 para toda a área de trabalho (Dashboard + Colunas de simulação)
  sheet.getRange("A1:G1005").setFontSize(13);
  
  // Negrito nos cabeçalhos e rótulos
  sheet.getRange("A1:G1").setFontWeight("bold");
  sheet.getRange("B2:B4").setFontWeight("bold");
  sheet.getRange("F2:F4").setFontWeight("bold");
  sheet.getRange("E5").setFontWeight("bold");
  sheet.getRange("G5").setFontWeight("bold");
  
  // Cores da Paleta
  var greenColor = "#d9ead3"; // Verde claro para inputs
  var blueColor  = "#cfe2f3"; // Azul claro para títulos e rótulos
  
  // Aplicando Azul Claro (Títulos e Rótulos)
  sheet.getRange("A1:E1").setBackground(blueColor);
  sheet.getRange("G1").setBackground(blueColor);
  sheet.getRange("B2:B4").setBackground(blueColor);
  sheet.getRange("F2:F4").setBackground(blueColor);
  sheet.getRange("E5").setBackground(blueColor);
  sheet.getRange("G5").setBackground(blueColor);
  
  // Aplicando Verde Claro (Inputs do Usuário)
  sheet.getRange("A2:A1000").setBackground(greenColor); // Coluna de Histórico
  sheet.getRange("C2:C4").setBackground(greenColor);    // Inputs 3 pontos
  sheet.getRange("F1").setBackground(greenColor);       // Unidade
  sheet.getRange("F5").setBackground(greenColor);       // Valor Alvo
  
  // Alinhamento centralizado para colunas de dados para ficar mais limpo
  sheet.getRange("C1:G5").setHorizontalAlignment("center");
  
  // Largura das Colunas (Substituindo o autoResize para dar mais "respiro" visual)
  sheet.setColumnWidth(1, 160); // A - Histórico
  sheet.setColumnWidth(2, 180); // B - Rótulos 3 Pontos
  sheet.setColumnWidth(3, 140); // C - Estimativas Iniciais
  sheet.setColumnWidth(4, 150); // D - Cenários Gerados
  sheet.setColumnWidth(5, 120); // E - Unidade / Rótulo Alvo
  sheet.setColumnWidth(6, 120); // F - Percentis / Input Alvo
  sheet.setColumnWidth(7, 140); // G - Resultados

  // Formatação de números (Milhar e 2 casas decimais)
  sheet.getRange("C2:C4").setNumberFormat("#,##0.00"); // Inputs 3 pontos
  sheet.getRange("F5").setNumberFormat("#,##0.00");    // Valor Alvo
}

/**
 * Gera um número aleatório em uma Distribuição Beta (Motor PERT).
 */
function randomBeta(alpha, beta) {
  var u1, u2, v1, v2;
  do {
    u1 = Math.random();
    u2 = Math.random();
    v1 = Math.pow(u1, 1 / alpha);
    v2 = Math.pow(u2, 1 / beta);
  } while (v1 + v2 > 1);
  return v1 / (v1 + v2);
}

/**
 * Executa a Simulação, cola os cenários e atualiza os dados gráficos (Versão Otimizada).
 */
function runMonteCarlo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var numSims = 10000; 
  var unit = sheet.getRange("F1").getValue() || "";
  
  // Leitura em bloco para ganhar performance
  var inputValues = sheet.getRange("C2:C4").getValues();
  var min = Number(inputValues[0][0]);
  var mode = Number(inputValues[1][0]);
  var max = Number(inputValues[2][0]);
  var chute = sheet.getRange("F5").getValue();
  
  if (isNaN(min) || isNaN(mode) || isNaN(max) || min > mode || mode > max) {
    ui.alert("Erro de Parâmetros: Verifique se a regra Mínimo <= Mais Provável <= Máximo está sendo respeitada.");
    return;
  }

  var results = [];
  var scenariosOutput = [];
  
  var alpha = 1 + 4 * ((mode - min) / (max - min));
  var betaValue = 1 + 4 * ((max - mode) / (max - min));
  
  // Motor Matemático (Instantâneo)
  for (var i = 0; i < numSims; i++) {
    var betaRandom = randomBeta(alpha, betaValue);
    var simValue = min + (betaRandom * (max - min));
    results.push(simValue);
    scenariosOutput.push([simValue]); 
  }
  
  // Limpa e cola os novos cenários de uma vez
  sheet.getRange("D2:D").clearContent();
  var rangeCenarios = sheet.getRange(2, 4, numSims, 1);
  rangeCenarios.setValues(scenariosOutput);
  rangeCenarios.setNumberFormat("#,##0.00");
  
  var sortedResults = results.slice().sort(function(a, b) { return a - b; });
  
  var p10 = sortedResults[Math.floor(numSims * 0.10)];
  var p50 = sortedResults[Math.floor(numSims * 0.50)];
  var p90 = sortedResults[Math.floor(numSims * 0.90)];
  
  var formatValue = function(val) {
    var formattedNum = val.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return unit ? formattedNum + " " + unit : formattedNum;
  };
  
  // ENVIO EM BLOCO (Batch) para os Percentis: 3x mais rápido que enviar um por um
  var percentisOutput = [
    [formatValue(p10)],
    [formatValue(p50)],
    [formatValue(p90)]
  ];
  sheet.getRange("G2:G4").setValues(percentisOutput);
  
  // Calcula a confiabilidade do chute
  if (chute !== "" && !isNaN(chute)) {
    var count = 0;
    for (var j = 0; j < sortedResults.length; j++) { if (sortedResults[j] <= Number(chute)) count++; }
    var prob = (count / numSims);
    sheet.getRange("G6").setValue((prob * 100).toFixed(2) + "%").setHorizontalAlignment("center").setFontWeight("bold").setFontSize(13);
  } else {
    sheet.getRange("G6").setValue("Insira número").setFontSize(13);
  }

  // ==========================================
  // ATUALIZAÇÃO DOS GRÁFICOS (Gargalo resolvido)
  // ==========================================
  
  var dataSheetName = "Dados_Simulacao_Oculta";
  var dataSheet = spreadsheet.getSheetByName(dataSheetName);
  
  if (!dataSheet) {
    dataSheet = spreadsheet.insertSheet(dataSheetName);
  }
  
  // Prepara os dados do gráfico
  var chartData = [["Valores Simulados", "Probabilidade Acumulada"]];
  for (var k = 0; k < numSims; k++) {
    var cumulProb = (k + 1) / numSims; 
    chartData.push([sortedResults[k], cumulProb]);
  }
  
  // Sobrescreve a aba oculta (os gráficos vão ler isso e piscar sozinhos)
  dataSheet.getRange(1, 1, chartData.length, 2).setValues(chartData);
  dataSheet.hideSheet(); 
  
  // LÓGICA DE PERFORMANCE: Só desenha os gráficos se eles não existirem!
  var existingCharts = sheet.getCharts();
  
  if (existingCharts.length < 2) {
    // Se apagou os gráficos sem querer, ele recria. Caso contrário, ignora essa parte demorada.
    var sCurveChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, numSims + 1, 2))
      .setPosition(1, 9, 0, 0) 
      .setOption('title', 'Curva S (Probabilidade Acumulada)')
      .setOption('hAxis', {title: 'Valores Estimados (' + unit + ')'})
      .setOption('vAxis', {title: 'Certeza / Probabilidade', format: '#%'})
      .setOption('legend', {position: 'none'})
      .setOption('curveType', 'function')
      .build();
      
    sheet.insertChart(sCurveChart);
    
    var densityChart = sheet.newChart()
      .setChartType(Charts.ChartType.HISTOGRAM)
      .addRange(dataSheet.getRange(1, 1, numSims + 1, 1))
      .setPosition(16, 9, 0, 0) 
      .setOption('title', 'Densidade do Risco (Frequência das Iterações)')
      .setOption('hAxis', {title: 'Valores Estimados (' + unit + ')'})
      .setOption('vAxis', {title: 'Frequência (Nº de Cenários)'})
      .setOption('legend', {position: 'none'})
      .build();
      
    sheet.insertChart(densityChart);
  }
}
