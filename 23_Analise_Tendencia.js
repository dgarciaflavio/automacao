// =================================================================
// --- BLOCO 23: ANÁLISE DE TENDÊNCIA (CORRIGIDO: COLUNAS CERTAS) ---
// =================================================================

function gerarRelatorioTendencia() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const toast = (msg) => ss.toast(msg, "Analise de Tendência", 5);
    toast("Carregando dados globais...");

    // 1. DADOS GLOBAIS (ENTRADAS/SAIDAS REAIS)
    const dados = obterDadosEntradasGlobal(); 
    
    // 2. DADOS LOCAIS (PARA PEGAR O CMM E DESCRIÇÃO)
    const abaDados = ss.getSheetByName("dados");
    if (!abaDados) throw new Error("Aba 'dados' não encontrada.");
    
    const mapaCMM = new Map();
    const mapaDescricoes = new Map();
    
    const lastRow = abaDados.getLastRow();
    if (lastRow >= 5) {
      const v = abaDados.getRange(5, 1, lastRow - 4, 8).getValues(); 
      v.forEach(r => {
        const cod = _norm(r[1]); // Coluna B
        const desc = String(r[2]).trim(); // Coluna C (DESCRIÇÃO LOCAL)
        const cmm = parseFloat(r[7]) || 0; // Coluna H
        
        if (cod) {
            mapaCMM.set(cod, cmm);
            mapaDescricoes.set(cod, desc);
        }
      });
    }

    // 3. CALCULAR CONSUMO RECENTE (ÚLTIMOS 30 e 60 DIAS)
    const hoje = new Date();
    const data30dias = new Date(); data30dias.setDate(hoje.getDate() - 30);
    const data60dias = new Date(); data60dias.setDate(hoje.getDate() - 60);

    const consumo30 = new Map();
    const consumo60 = new Map();

    dados.forEach(r => {
      // --- CORREÇÃO DE COLUNAS AQUI ---
      // Coluna O (Índice 14) = Data de Recebimento (Movimentação Real)
      // Coluna L (Índice 11) = Quantidade Entregue (Consumo/Entrada)
      // Coluna C (Índice 2)  = Código
      
      const dataMov = r[14]; // Antes estava r[0] (Errado)
      const cod = _norm(r[2]); 
      const qtdMov = parseFloat(r[11]) || 0; // Antes estava r[12] (Errado)
      
      // Verifica se a data é válida
      let dataValida = null;
      if (dataMov instanceof Date) {
        dataValida = dataMov;
      } else if (typeof dataMov === 'string') {
        dataValida = _parseDataSegura(dataMov);
      }

      if (cod && dataValida) {
        if (dataValida >= data30dias) {
          consumo30.set(cod, (consumo30.get(cod) || 0) + qtdMov);
        }
        if (dataValida >= data60dias) {
          consumo60.set(cod, (consumo60.get(cod) || 0) + qtdMov);
        }
      }
    });

    // 4. ANÁLISE DE DESVIO
    const relatorio = [];
    
    mapaCMM.forEach((cmm, cod) => {
      const qtd30 = consumo30.get(cod) || 0;
      
      // Regra de Ignorar itens muito pequenos para evitar ruído
      if (cmm < 5 && qtd30 < 5) return;

      const desvio = qtd30 - cmm;
      const percentual = cmm > 0 ? (desvio / cmm) : (qtd30 > 0 ? 1 : 0); 

      let status = "Estável";
      let cor = null;

      if (percentual > 0.30) { 
        status = "🔥 Aceleração Alta";
        cor = "#ea9999"; // Vermelho
      } else if (percentual < -0.30) { 
        status = "❄️ Desaceleração";
        cor = "#cfe2f3"; // Azul
      }

      if (status !== "Estável") {
        const descricaoFinal = mapaDescricoes.get(cod) || "Descrição não encontrada";
        
        relatorio.push([
          cod,
          descricaoFinal,
          cmm,
          qtd30,
          percentual,
          status
        ]);
      }
    });

    relatorio.sort((a, b) => b[4] - a[4]);

    // 5. GERAR ABA DE RELATÓRIO
    let abaRel = ss.getSheetByName("BI_Tendencia");
    if (!abaRel) abaRel = ss.insertSheet("BI_Tendencia");
    abaRel.clear();

    const header = ["Código", "Descrição", "CMM (Histórico)", "Consumo (30d)", "Variação %", "Diagnóstico"];
    abaRel.getRange(1, 1, 1, 6).setValues([header])
      .setFontWeight("bold").setBackground("#134f5c").setFontColor("white");

    if (relatorio.length > 0) {
      abaRel.getRange(2, 1, relatorio.length, 6).setValues(relatorio);
      abaRel.getRange(2, 5, relatorio.length, 1).setNumberFormat("+0%"); 
      
      const cores = relatorio.map(r => {
        const st = r[5];
        if (st.includes("Aceleração")) return ["#ea9999"];
        if (st.includes("Desaceleração")) return ["#cfe2f3"];
        return [null];
      });
      abaRel.getRange(2, 6, relatorio.length, 1).setBackgrounds(cores);
    }

    abaRel.autoResizeColumns(1, 6);
    abaRel.setColumnWidth(2, 350); 
    
    ui.alert(`Análise concluída!\n${relatorio.length} itens com anomalia de consumo detectados.`);

  } catch (e) {
    ui.alert("Erro na Análise de Tendência: " + e.message);
  }
}
