// =================================================================
// --- BLOCO 23: ANÁLISE DE TENDÊNCIA (CORRIGIDO: DESCRIÇÃO DA GUIA DADOS) ---
// =================================================================

function gerarRelatorioTendencia() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const toast = (msg) => ss.toast(msg, "Analise de Tendência", 5);
    toast("Carregando dados globais...");

    // 1. DADOS GLOBAIS (ENTRADAS/SAIDAS REAIS)
    const dados = obterDadosEntradasGlobal(); // Usa o helper existente
    
    // 2. DADOS LOCAIS (PARA PEGAR O CMM E DESCRIÇÃO)
    const abaDados = ss.getSheetByName("dados");
    if (!abaDados) throw new Error("Aba 'dados' não encontrada.");
    
    // Mapeia o CMM e a DESCRIÇÃO de cada item (Coluna B=Cod, Coluna C=Desc, Coluna H=CMM)
    const mapaCMM = new Map();
    const mapaDescricoes = new Map(); // Mapa para guardar as descrições corretas
    
    const lastRow = abaDados.getLastRow();
    if (lastRow >= 5) {
      // Pega até a coluna H (índice 8)
      const v = abaDados.getRange(5, 1, lastRow - 4, 8).getValues(); 
      v.forEach(r => {
        const cod = _norm(r[1]); // Coluna B
        const desc = String(r[2]).trim(); // Coluna C (DESCRIÇÃO LOCAL)
        const cmm = parseFloat(r[7]) || 0; // Coluna H
        
        if (cod) {
            mapaCMM.set(cod, cmm);
            mapaDescricoes.set(cod, desc); // Salva a descrição local
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
      // Estrutura do Global: Col A=Data, Col C=Cod, Col M=Qtd Entregue (Consumo)
      const dataMov = r[0]; // Coluna A
      const cod = _norm(r[2]); // Coluna C
      const qtdMov = parseFloat(r[12]) || 0; // Coluna M (Qtd Entregue - Proxy de Consumo/Giro)
      
      if (cod && dataMov instanceof Date) {
        if (dataMov >= data30dias) {
          consumo30.set(cod, (consumo30.get(cod) || 0) + qtdMov);
        }
        if (dataMov >= data60dias) {
          consumo60.set(cod, (consumo60.get(cod) || 0) + qtdMov);
        }
      }
    });

    // 4. ANÁLISE DE DESVIO
    const relatorio = [];
    
    mapaCMM.forEach((cmm, cod) => {
      // Só analisa itens que têm movimentação ou CMM relevante
      const qtd30 = consumo30.get(cod) || 0;
      
      // Regra de Ignorar itens muito pequenos para evitar ruído
      if (cmm < 5 && qtd30 < 5) return;

      const desvio = qtd30 - cmm;
      const percentual = cmm > 0 ? (desvio / cmm) : (qtd30 > 0 ? 1 : 0); // 100% se CMM 0 e teve consumo

      let status = "Estável";
      let cor = null;

      if (percentual > 0.30) { // +30%
        status = "🔥 Aceleração Alta";
        cor = "#ea9999"; // Vermelho
      } else if (percentual < -0.30) { // -30%
        status = "❄️ Desaceleração";
        cor = "#cfe2f3"; // Azul
      }

      if (status !== "Estável") {
        // Agora pega a descrição do mapa local (Coluna C da guia dados)
        const descricaoFinal = mapaDescricoes.get(cod) || "Descrição não encontrada";
        
        relatorio.push([
          cod,
          descricaoFinal, // Usa a descrição correta
          cmm,
          qtd30,
          percentual,
          status
        ]);
      }
    });

    // Ordenar pelos maiores desvios percentuais
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
      abaRel.getRange(2, 5, relatorio.length, 1).setNumberFormat("+0%"); // Formato +30%, -10%
      
      // Pintar diagnósticos
      const cores = relatorio.map(r => {
        const st = r[5];
        if (st.includes("Aceleração")) return ["#ea9999"];
        if (st.includes("Desaceleração")) return ["#cfe2f3"];
        return [null];
      });
      abaRel.getRange(2, 6, relatorio.length, 1).setBackgrounds(cores);
    }

    abaRel.autoResizeColumns(1, 6);
    // Força largura maior para a descrição
    abaRel.setColumnWidth(2, 350); 
    
    ui.alert(`Análise concluída!\n${relatorio.length} itens com anomalia de consumo detectados.`);

  } catch (e) {
    ui.alert("Erro na Análise de Tendência: " + e.message);
  }
}
