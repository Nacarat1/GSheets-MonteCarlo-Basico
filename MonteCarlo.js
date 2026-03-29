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
  sheet.getRange("C2").setValue(80);
  sheet.getRange("B3").setValue("Mais Provável:");
  sheet.getRange("C3").setValue(100);
  sheet.getRange("B4").setValue("Máximo:");
  sheet.getRange("C4").setValue(150);
  
  // Percentis e Alvo
  sheet.getRange("F2").setValue("Percentil 10");
  sheet.getRange("F3").setValue("Percentil 50");
  sheet.getRange("F4").setValue("Percentil 90");
  sheet.getRange("E5").setValue("Valor alvo:");
  sheet.getRange("F5").setValue(120);
  sheet.getRange("G5").setValue("Probabilidade");
  
  // ==========================================
  // FORMATAÇÃO VISUAL E ESTÉTICA
  // ==========================================
  
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
 * Executa a Simulação, cola os cenários e gera os gráficos executivos.
 */
function runMonteCarlo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var numSims = 1000; 
  var unit = sheet.getRange("F1").getValue() || "";
  
  var histValues = sheet.getRange("A2:A").getValues().filter(function(r) { 
    return r[0] !== "" && !isNaN(r[0]); 
  });
  
  var min, mode, max;
  
  if (histValues.length > 0) {
    var flatHist = histValues.map(function(r) { return Number(r[0]); });
    flatHist.sort(function(a, b) { return a - b; });
    min = flatHist[0];
    max = flatHist[flatHist.length - 1];
    var mid = Math.floor(flatHist.length / 2);
    mode = flatHist.length % 2 !== 0 ? flatHist[mid] : (flatHist[mid - 1] + flatHist[mid]) / 2;
  } else {
    min = Number(sheet.getRange("C2").getValue());
    mode = Number(sheet.getRange("C3").getValue());
    max = Number(sheet.getRange("C4").getValue());
  }
  
  if (isNaN(min) || isNaN(mode) || isNaN(max) || min > mode || mode > max) {
    ui.alert("Erro de Parâmetros: Verifique se a regra Mínimo <= Mais Provável (ou Mediana) <= Máximo está sendo respeitada.");
    return;
  }

  var results = [];
  var scenariosOutput = [];
  
  if (min === max) {
    for (var i = 0; i < numSims; i++) { 
      results.push(min);
      scenariosOutput.push([min]);
    }
  } else {
    var alpha = 1 + 4 * ((mode - min) / (max - min));
    var betaValue = 1 + 4 * ((max - mode) / (max - min));
    
    for (var i = 0; i < numSims; i++) {
      var betaRandom = randomBeta(alpha, betaValue);
      var simValue = min + (betaRandom * (max - min));
      results.push(simValue);
      scenariosOutput.push([simValue]); 
    }
  }
  
  sheet.getRange("D2:D").clearContent();
  sheet.getRange(2, 4, numSims, 1).setValues(scenariosOutput);
  
  var sortedResults = results.slice().sort(function(a, b) { return a - b; });
  
  var p10 = sortedResults[Math.floor(numSims * 0.10)];
  var p50 = sortedResults[Math.floor(numSims * 0.50)];
  var p90 = sortedResults[Math.floor(numSims * 0.90)];
  
  var formatValue = function(val) {
    return unit ? val.toFixed(2) + " " + unit : val.toFixed(2);
  };
  
  sheet.getRange("G2").setValue(formatValue(p10));
  sheet.getRange("G3").setValue(formatValue(p50));
  sheet.getRange("G4").setValue(formatValue(p90));
  
  var chute = sheet.getRange("F5").getValue();
  if (chute !== "" && !isNaN(chute)) {
    var count = 0;
    for (var j = 0; j < sortedResults.length; j++) { if (sortedResults[j] <= Number(chute)) count++; }
    var prob = (count / numSims);
    sheet.getRange("G6").setValue((prob * 100).toFixed(2) + "%"); // Corrigido para G6 caso queira deixar abaixo do título, ou manter na mesma célula formatada
    // OBS: Como o título "Probabilidade" ficou em G5, o valor da probabilidade agora cairá em G6 para não apagar o título azul.
    sheet.getRange("G6").setValue((prob * 100).toFixed(2) + "%").setHorizontalAlignment("center").setFontWeight("bold");
  } else {
    sheet.getRange("G6").setValue("Insira número");
  }

  // ==========================================
  // GERAÇÃO DE GRÁFICOS
  // ==========================================
  
  var dataSheetName = "Dados_Simulacao_Oculta";
  var dataSheet = spreadsheet.getSheetByName(dataSheetName);
  
  if (!dataSheet) {
    dataSheet = spreadsheet.insertSheet(dataSheetName);
  }
  dataSheet.clear();
  
  var chartData = [["Valores Simulados", "Probabilidade Acumulada"]];
  for (var k = 0; k < numSims; k++) {
    var cumulProb = (k + 1) / numSims; 
    chartData.push([sortedResults[k], cumulProb]);
  }
  
  dataSheet.getRange(1, 1, chartData.length, 2).setValues(chartData);
  dataSheet.hideSheet(); 
  
  var existingCharts = sheet.getCharts();
  for (var c = 0; c < existingCharts.length; c++) {
    sheet.removeChart(existingCharts[c]);
  }
  
  var sCurveChart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataSheet.getRange(1, 1, numSims + 1, 2))
    .setPosition(2, 9, 0, 0) 
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
    .setPosition(24, 9, 0, 0) 
    .setOption('title', 'Densidade do Risco (Frequência das Iterações)')
    .setOption('hAxis', {title: 'Valores Estimados (' + unit + ')'})
    .setOption('vAxis', {title: 'Frequência (Nº de Cenários)'})
    .setOption('legend', {position: 'none'})
    .build();
    
  sheet.insertChart(densityChart);
}
