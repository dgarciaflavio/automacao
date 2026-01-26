// =================================================================
// --- BLOCO 7: COMPILAÇÃO LOCAL (COM VALIDAÇÃO OFICIAL) ---
// =================================================================

function compilarDadosLocal() { 
  try {
    const dados = obterDadosEntradasGlobal();
    compilarDados(dados); 
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao buscar dados para compilação: " + e.message);
  }
}

function compilarDados(dadosGlobais) {
  var ui = SpreadsheetApp.getUi();
  try {
    if (!dadosGlobais) {
      dadosGlobais = obterDadosEntradasGlobal();
    }

    var ssDestino = SpreadsheetApp.getActiveSpreadsheet();
    var abaDestino = ssDestino.getSheetByName(CONFIG.destino.nomeAba);
    const abaRecProvisorio = ssDestino.getSheetByName("Rec.Provisorio");
    
    if (!abaDestino || !abaRecProvisorio) {
      throw new Error("Abas de destino ('Compilados' ou 'Rec.Provisorio') não encontradas.");
    }

    // 1. Mapa de Qtd Oficial (Dados Globais)
    // Essencial para saber se houve entrada oficial (L > 0)
    const mapaQtdOficial = new Map();
    dadosGlobais.forEach(r => {
       const emp = String(r[0]).trim();
       const cod = _norm(r[2]);
       const qS = Math.round(parseFloat(r[11]) || 0);
       if (emp && cod) {
          const k = `${emp}|${cod}`;
          const atual = mapaQtdOficial.get(k) || 0;
          mapaQtdOficial.set(k, atual + qS);
       }
    });

    // Mapeamento Rec. Provisório
    const recProvisorioMap = new Map();
    const lastRowRec = abaRecProvisorio.getLastRow();
    if (lastRowRec >= 2) {
      const recData = abaRecProvisorio.getRange(2, 1, lastRowRec - 1, 6).getValues();
      recData.forEach(row => {
        const itemCode = _norm(row[0]);
        const qtd = Math.round(parseFloat(row[2]) || 0); 
        const empenhoStr = String(row[5]).trim();
        if (itemCode && empenhoStr.includes('/')) {
          const parts = empenhoStr.split('/');
          if (parts.length === 2) {
            const numero = parts[0].padStart(4, '0');
            const ano = (parts[1].length === 2) ? '20' + parts[1] : parts[1];
            recProvisorioMap.set(`${ano}${numero}|${itemCode}`, qtd); 
          }
        }
      });
    }

    // Mapeamento Eliminadas
    const mapaStatusEliminados = new Map();
    dadosGlobais.forEach(linha => {
        if (_norm(linha[13]) === "ELIMINADA") {
            const empenho = parseInt(linha[0], 10);
            const codigo = _norm(linha[2]); 
            if (empenho && codigo) {
              mapaStatusEliminados.set(`${empenho}|${codigo}`, "Eliminada");
            }
        }
    });

    var dadosCompilados = [];
    const fontes = [
        {id: CONFIG.ids.materiais, nomeAba: CONFIG.abas.materiais},
        {id: CONFIG.ids.medicamentos, nomeAba: CONFIG.abas.medicamentos}
    ];

    fontes.forEach(fonte => {
      try {
        var ssFonte = SpreadsheetApp.openById(fonte.id);
        var abaFonte = ssFonte.getSheetByName(fonte.nomeAba);
        if (abaFonte) {
          var dadosFonte = abaFonte.getDataRange().getValues();
          if (dadosFonte.length > 1) dadosCompilados.push(...dadosFonte.slice(1));
        }
      } catch (e) { 
        console.warn(`Aviso: Não foi possível ler a fonte ${fonte.id} (${e.message})`);
      }
    });

    // --- PROCESSAMENTO FINAL ---
    dadosCompilados.forEach(linha => {
      const empenho = linha[0] ? parseInt(linha[0], 10) : null;
      const codigo = _norm(linha[5]); 
      const chave = (empenho && codigo) ? `${empenho}|${codigo}` : null;
      const isProvisorio = (chave && recProvisorioMap.has(chave));

      if (chave && mapaStatusEliminados.has(chave)) {
          linha[18] = "Eliminada";
      } else {
          var vI = Math.round(parseFloat(linha[8]) || 0);  // Qtd Empenho
          
          // AQUI ESTÁ O SEGREDO:
          // vP_Visual -> É o valor que vem da planilha Materiais/Meds (que pode ser o provisório 960)
          var vP_Visual = Math.round(parseFloat(linha[15]) || 0); 
          
          // qS_Real_Oficial -> É o valor REAL do sistema (Dados Globais). Ex: 14
          var qS_Real_Oficial = 0;
          if (chave) qS_Real_Oficial = mapaQtdOficial.get(chave) || 0;

          // O Saldo deve bater com o Visual (Físico)
          var vQ = vI - vP_Visual; 
          
          linha[8] = vI;
          linha[15] = vP_Visual; // Mantém o valor visual (Físico)
          linha[16] = vQ;        // Mantém o saldo Físico

          var statusOrig = linha[18] ? linha[18].toString().trim() : '';
          
          // O status é decidido pelo Oficial (para quebrar o provisório) e pelo Saldo Físico (para definir resíduo)
          linha[18] = _calcularStatusUnificado(vI, qS_Real_Oficial, vQ, isProvisorio, statusOrig);
      }
    });

    // Escrita na Aba Compilados
    if (dadosCompilados.length > 0) {
      abaDestino.unhideColumn(abaDestino.getRange("A:U"));
      if (abaDestino.getLastRow() > 1) {
        abaDestino.getRange(2, 1, abaDestino.getLastRow()-1, abaDestino.getMaxColumns()).clearContent().clearFormat();
      }
      
      abaDestino.getRange(2, 1, dadosCompilados.length, dadosCompilados[0].length).setValues(dadosCompilados);
      
      var numRows = dadosCompilados.length;
      abaDestino.getRange(2, 9, numRows, 1).setNumberFormat("0");
      abaDestino.getRange(2, 16, numRows, 1).setNumberFormat("0");
      abaDestino.getRange(2, 17, numRows, 1).setNumberFormat("0");

      var valores = abaDestino.getRange(2, 1, numRows, 21).getValues();
      var backgrounds = valores.map(l => {
        return new Array(21).fill(CONFIG.cores[_norm(l[18])] || null);
      });
      abaDestino.getRange(2, 1, numRows, 21).setBackgrounds(backgrounds);
      
      abaDestino.hideColumns(2, 3); 
      abaDestino.hideColumns(8, 1); 
      abaDestino.hideColumns(10, 6); 
      abaDestino.hideColumns(18, 1);
    }
  } catch (e) { 
    ui.alert("Erro Compilados Local", e.message, ui.ButtonSet.OK); 
  }
}