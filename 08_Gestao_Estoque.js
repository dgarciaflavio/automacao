// =================================================================
// --- BLOCO 8: GESTÃO DE ESTOQUE (ATUALIZADO: META 6 MESES) ---
// =================================================================

function sincronizarControleEstoque() {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // 1. MAPEAMENTO DE EMPENHOS ATIVOS
    const destinos = [
      { id: CONFIG.ids.materiais, aba: CONFIG.abas.materiais },
      { id: CONFIG.ids.medicamentos, aba: CONFIG.abas.medicamentos }
    ];
    
    const mapaEmpenhosAtivos = new Map();
    destinos.forEach(dest => {
      try {
        const ssRemoto = SpreadsheetApp.openById(dest.id);
        const wsAlvo = ssRemoto.getSheetByName(dest.aba);
        if (wsAlvo && wsAlvo.getLastRow() > 1) {
          const dadosR = wsAlvo.getRange(2, 1, wsAlvo.getLastRow() - 1, 6).getValues();
          dadosR.forEach(r => {
            const cod = _norm(r[5]); 
            if (cod) {
              if (!mapaEmpenhosAtivos.has(cod)) {
                mapaEmpenhosAtivos.set(cod, { empenho: r[0], fornecedor: r[4] });
              }
            }
          });
        }
      } catch (e) { 
        console.log("Aviso: Falha ao acessar planilha remota para estoque: " + e.message); 
      }
    });

    // 2. PROCESSAMENTO DA BASE CENTRAL (Guia DADOS)
    const abaDados = ssLocal.getSheetByName(CONFIG.abas.fonteDadosNome); 
    if (!abaDados) throw new Error("Aba 'dados' não encontrada.");
    
    const ultLinha = abaDados.getLastRow();
    if (ultLinha < 5) throw new Error("Guia 'dados' não possui registros suficientes a partir da linha 5.");
    
    const vDados = abaDados.getRange(5, 1, ultLinha - 4, 39).getValues();
    const dadosConsolidadosLocal = [];

    vDados.forEach(r => {
      const itemCod = _norm(r[1]); 
      if (!itemCod) return; 

      const localRef = _norm(r[0]); 
      const itemDesc = r[2];       
      
      const infoEstoque = parseFloat(r[6]) || 0; // Coluna G (Saldo)
      const infoCMM = parseFloat(r[7]) || 0;     // Coluna H (CMM)
      
      const processoSEI = r[38];   
      
      // CLASSIFICAÇÃO
      let tipo = "MATERIAL"; 
      if (localRef === "FAR") {
          tipo = "MEDICAMENTO";
      } else if (localRef === "ALM") {
          tipo = "MATERIAL";
      } else {
          if (/^\d/.test(itemCod)) {
             tipo = "MEDICAMENTO";
          }
      }
      
      const infoRemota = mapaEmpenhosAtivos.get(itemCod) || { empenho: "---", fornecedor: "---" };
      const alerta = _calcularAlertaEstoque(infoEstoque, infoCMM);

      // CÁLCULOS DE COBERTURA (ATUAL)
      let diasCobertura = 0;
      let dataEsgotamentoAtual = "";
      
      if (infoCMM > 0) {
         diasCobertura = Math.floor(infoEstoque / (infoCMM / 30));
         const hoje = new Date();
         hoje.setDate(hoje.getDate() + diasCobertura);
         dataEsgotamentoAtual = diasCobertura > 365 ? "Estável (>1 ano)" : hoje;
      } else {
         diasCobertura = 9999; 
         dataEsgotamentoAtual = (infoEstoque > 0) ? "Sem Consumo" : "Zerado";
      }

      // --- CÁLCULOS: SUGESTÃO DE PEDIDO (AGORA 6 MESES) ---
      let sugestaoPedido = 0;
      let previsaoComSugestao = "";

      if (infoCMM > 0) {
        // META ALTERADA: 6 meses
        const metaAlvo = infoCMM * 6;
        
        // Se o estoque atual for menor que a meta, sugere a diferença
        if (infoEstoque < metaAlvo) {
           sugestaoPedido = Math.round(metaAlvo - infoEstoque);
        } else {
           sugestaoPedido = 0; // Já tem estoque suficiente
        }

        // Calcula quando esse estoque (Atual + Sugestão) vai acabar
        const estoqueFuturo = infoEstoque + sugestaoPedido;
        const diasFuturos = Math.floor(estoqueFuturo / (infoCMM / 30));
        
        const dataFutura = new Date();
        dataFutura.setDate(dataFutura.getDate() + diasFuturos);
        previsaoComSugestao = dataFutura;

      } else {
        sugestaoPedido = 0;
        previsaoComSugestao = "Sem Consumo";
      }

      dadosConsolidadosLocal.push([
        tipo,
        itemCod,
        itemDesc,
        infoRemota.fornecedor,
        infoRemota.empenho,
        infoEstoque,
        infoCMM,
        alerta,
        diasCobertura === 9999 ? "-" : diasCobertura,
        dataEsgotamentoAtual,
        processoSEI,
        sugestaoPedido,      // Coluna L (Calculado para 6 meses)
        previsaoComSugestao  // Coluna M
      ]);
    });

    // 3. ATUALIZAÇÃO DA GUIA "Cont.Estoque"
    let abaLocalEstoque = ssLocal.getSheetByName(CONFIG.abas.estoqueRemoto);
    if (!abaLocalEstoque) abaLocalEstoque = ssLocal.insertSheet(CONFIG.abas.estoqueRemoto);
    
    abaLocalEstoque.clear();
    
    // CABEÇALHO ATUALIZADO
    const cabecalho = [
      "Tipo", "Item", "Descrição", "Fornecedor", "Empenho", 
      "Estoque", "CMM", "Status", "Cobertura (Dias)", "Previsão Esgotamento (Atual)", 
      "Processo SEI", "Sugestão Pedido (6 Meses)", "Prev. Esgotamento (Sugestão)"
    ];
    
    abaLocalEstoque.getRange(1, 1, 1, 13).setValues([cabecalho])
      .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");

    if (dadosConsolidadosLocal.length > 0) {
      abaLocalEstoque.getRange(2, 1, dadosConsolidadosLocal.length, 13).setValues(dadosConsolidadosLocal);
      
      // Formatações
      abaLocalEstoque.getRange(2, 10, dadosConsolidadosLocal.length, 1).setNumberFormat("dd/mm/yyyy"); 
      abaLocalEstoque.getRange(2, 11, dadosConsolidadosLocal.length, 1).setNumberFormat("@"); 
      
      abaLocalEstoque.getRange(2, 12, dadosConsolidadosLocal.length, 1).setNumberFormat("#,##0"); 
      abaLocalEstoque.getRange(2, 13, dadosConsolidadosLocal.length, 1).setNumberFormat("dd/mm/yyyy");

      // Cores de Status
      const coresLocalStatus = dadosConsolidadosLocal.map(l => {
          const st = l[7];
          if (st === "Crítico") return [CONFIG.cores.ALERTA_CRITICO];
          if (st === "Atenção") return [CONFIG.cores.ALERTA_ATENCAO];
          if (st === "Suprimento Ok") return [CONFIG.cores.ALERTA_OK];
          return [null];
      });
      abaLocalEstoque.getRange(2, 8, dadosConsolidadosLocal.length, 1).setBackgrounds(coresLocalStatus);
      
      abaLocalEstoque.autoResizeColumns(1, 13);
      
      if (abaLocalEstoque.getFilter()) abaLocalEstoque.getFilter().remove();
      abaLocalEstoque.getDataRange().createFilter();
    }
    
    ui.alert("Sucesso", "Controle de Estoque atualizado! Sugestão recalculada para 6 meses.", ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("Erro no Estoque", e.message, ui.ButtonSet.OK);
  }
}

/**
 * Função Auxiliar: Calcula o status visual do estoque com base no CMM
 */
function _calcularAlertaEstoque(estoque, cmm) {
  if (cmm <= 0) return estoque > 0 ? "Suprimento Ok" : "Crítico";
  const meses = estoque / cmm;
  if (meses < 2) return "Crítico";
  if (meses >= 3 && meses <= 5) return "Atenção";
  return "Suprimento Ok";
}