// =================================================================
// --- BLOCO 14: RELATÓRIOS DE INTELIGÊNCIA (BI) ---
// =================================================================

function gerarRelatorioPerformanceFornecedores() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const abaComp = ss.getSheetByName(CONFIG.destino.nomeAba); 
    if (!abaComp) throw new Error("Aba Compilados não encontrada.");

    const dados = abaComp.getRange(2, 1, abaComp.getLastRow() - 1, 19).getValues();
    const stats = new Map();

    dados.forEach(r => {
      const fornecedor = r[4] ? String(r[4]).trim().toUpperCase() : "NÃO INFORMADO";
      const status = r[18] ? String(r[18]).trim().toUpperCase() : "";

      if (status === "ELIMINADA" || status === "SOLICITAR ASSOCIAÇÃO NO EMS") return;

      if (!stats.has(fornecedor)) {
        stats.set(fornecedor, { total: 0, pendentes: 0, concluidos: 0 });
      }

      const s = stats.get(fornecedor);
      s.total++;

      if (status.includes("PENDENTE")) s.pendentes++;
      else if (status === "CONCLUÍDO") s.concluidos++;
    });

    const relatorio = [];
    stats.forEach((val, key) => {
      if (val.total > 1) {
        const taxaProblema = val.total > 0 ? (val.pendentes / val.total) : 0;
        relatorio.push([
          key, 
          val.total, 
          val.concluidos, 
          val.pendentes, 
          taxaProblema
        ]);
      }
    });

    relatorio.sort((a, b) => b[3] - a[3] || b[4] - a[4]);

    let abaRel = ss.getSheetByName("BI_Fornecedores");
    if (!abaRel) abaRel = ss.insertSheet("BI_Fornecedores");
    abaRel.clear();

    const cabecalho = ["Fornecedor", "Total de Itens", "Entregues", "Pendentes (Atraso)", "% de Pendência"];
    abaRel.getRange(1, 1, 1, 5).setValues([cabecalho])
      .setFontWeight("bold")
      .setBackground("#4c1130") 
      .setFontColor("white");

    if (relatorio.length > 0) {
      abaRel.getRange(2, 1, relatorio.length, 5).setValues(relatorio);
      abaRel.getRange(2, 5, relatorio.length, 1).setNumberFormat("0.0%");
      abaRel.getRange(2, 1, relatorio.length, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    }

    abaRel.autoResizeColumns(1, 5);
    abaRel.activate();
    ui.alert("Relatório de Performance Gerado com Sucesso!");

  } catch (e) {
    ui.alert("Erro BI Fornecedores: " + e.message);
  }
}

function gerarRelatorioFinanceiroExecutivo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const dadosEntradas = obterDadosEntradasGlobal();

    const respIni = ui.prompt("Filtro de Período", "Digite o ANO INICIAL (ex: 2024):", ui.ButtonSet.OK_CANCEL);
    if (respIni.getSelectedButton() !== ui.Button.OK) return;
    const anoInicial = parseInt(respIni.getResponseText().trim());

    const respFim = ui.prompt("Filtro de Período", "Digite o ANO FINAL (ex: 2025):", ui.ButtonSet.OK_CANCEL);
    if (respFim.getSelectedButton() !== ui.Button.OK) return;
    const anoFinal = parseInt(respFim.getResponseText().trim());

    if (isNaN(anoInicial) || isNaN(anoFinal)) {
      ui.alert("Por favor, digite anos válidos (apenas números).");
      return;
    }

    let totalEmpenhadoPeriodo = 0;
    let totalEntreguePeriodo = 0;
    let totalPendentePeriodo = 0;
    let totalResiduoPeriodo = 0;
    let countItens = 0;

    dadosEntradas.forEach(r => {
      const empenhoStr = String(r[0]).trim(); 
      const anoEmpenho = parseInt(empenhoStr.substring(0, 4));

      if (!anoEmpenho || anoEmpenho < anoInicial || anoEmpenho > anoFinal) return;

      const valorEmpenhado = parseFloat(r[10]) || 0; 
      const valorEntregue = parseFloat(r[12]) || 0;  
      
      const saldo = valorEmpenhado - valorEntregue; 
      
      countItens++;
      totalEmpenhadoPeriodo += valorEmpenhado;
      totalEntreguePeriodo += valorEntregue;

      if (saldo > 0.01) { 
        if (valorEmpenhado > 0 && saldo <= (valorEmpenhado * 0.10)) {
           totalResiduoPeriodo += saldo;
        } else {
           totalPendentePeriodo += saldo;
        }
      }
    });

    let abaFin = ss.getSheetByName("BI_Financeiro");
    if (!abaFin) abaFin = ss.insertSheet("BI_Financeiro");
    abaFin.clear();

    const titulo = [`PANORAMA FINANCEIRO (${anoInicial} a ${anoFinal})`, "", ""];
    const output = [
      titulo,
      ["MÉTRICA", "VALOR (R$)", "DETALHE"],
      ["Total Empenhado no Período", totalEmpenhadoPeriodo, `Fonte: Dados Globais (Col K) | Filtro: ${anoInicial}-${anoFinal}`],
      ["Total Efetivamente Entregue", totalEntreguePeriodo, "Fonte: Dados Globais (Col M)"],
      ["Valor Pendente (A Receber)", totalPendentePeriodo, "Saldo > 10% do Empenho (K - M)"],
      ["Valor em Resíduo", totalResiduoPeriodo, "Saldo <= 10% do Empenho (K - M)"]
    ];

    abaFin.getRange(1, 1, 6, 3).setValues(output);
    
    abaFin.getRange("A1:C1").merge().setBackground("#274e13").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    abaFin.getRange("A2:C2").setFontWeight("bold").setBackground("#b6d7a8");
    
    const rangeValores = abaFin.getRange("B3:B6");
    rangeValores.setNumberFormat("R$ #,##0.00");
    
    const charts = abaFin.getCharts();
    charts.forEach(c => abaFin.removeChart(c));
    
    const chart = abaFin.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(abaFin.getRange("A3:B5")) 
      .setPosition(8, 1, 0, 0)
      .setOption('title', `Finanças ${anoInicial}-${anoFinal} (Itens Analisados: ${countItens})`)
      .setOption('legend', {position: 'bottom'})
      .setOption('colors', ['#1c4587', '#38761d', '#cc0000']) 
      .build();
      
    abaFin.insertChart(chart);
    abaFin.autoResizeColumns(1, 3);
    abaFin.activate();
    
    if (countItens === 0) {
      ui.alert("Atenção", `Nenhum empenho encontrado iniciando entre ${anoInicial} e ${anoFinal}.`, ui.ButtonSet.OK);
    } else {
      ui.alert("Sucesso", `Relatório gerado com base em ${countItens} registros da memória.`, ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert("Erro BI Financeiro: " + e.message);
  }
}