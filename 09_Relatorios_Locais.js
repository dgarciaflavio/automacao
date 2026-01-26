// =================================================================
// --- BLOCO 9: RELATÓRIOS LOCAIS (ATUALIZADO) ---
// =================================================================

function gerarRelatorioValidadeAtas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const abaDestino = ss.getSheetByName("Validade De Atas");
    const abaDados = ss.getSheetByName("dados");
    if (!abaDados) throw new Error("Aba 'dados' não encontrada.");
    if (!abaDestino) throw new Error("Aba 'Validade De Atas' não encontrada.");
    const rawInicio = abaDestino.getRange("L1").getValue();
    const rawFim = abaDestino.getRange("M1").getValue();
    
    const dataInicial = _parseDataSegura(rawInicio);
    const dataFinal = _parseDataSegura(rawFim);
    const usarFiltro = (dataInicial && dataFinal);

    const lastRow = abaDados.getLastRow();
    if (lastRow < 5) { ui.alert("Aba 'dados' parece vazia."); return;
    }
    
    const valores = abaDados.getRange(5, 1, lastRow - 4, 39).getValues();
    const dadosFiltrados = [];

    for (let i = 0; i < valores.length; i++) {
      const linha = valores[i];
      const item = linha[1];
      if (!item) continue;

      const vencimentoAta = _parseDataSegura(linha[19]);
      
      let incluir = true;
      if (usarFiltro) {
        if (!vencimentoAta) {
          incluir = false;
        } else {
          const v = new Date(vencimentoAta); v.setHours(0,0,0,0);
          const ini = new Date(dataInicial); ini.setHours(0,0,0,0);
          const fim = new Date(dataFinal); fim.setHours(0,0,0,0);
          if (v < ini || v > fim) incluir = false;
        }
      }

      if (incluir) {
        dadosFiltrados.push([
          item,
          linha[2],
          linha[6],
          linha[8],
          linha[33],
          vencimentoAta,
          linha[38],
          linha[18]
        ]);
      }
    }

    if (abaDestino.getLastRow() > 1) {
      abaDestino.getRange(2, 1, abaDestino.getLastRow() - 1, 8).clearContent();
    }

    const cabecalho = ["Item", "Descrição", "Estoque", "Saldo em dias", "Saldo em Ata", "Vencimento da Ata", "Processo SEI", "Preço Unitário"];
    abaDestino.getRange(1, 1, 1, 8).setValues([cabecalho]).setFontWeight("bold").setBackground("#cfe2f3");

    if (dadosFiltrados.length > 0) {
      abaDestino.getRange(2, 1, dadosFiltrados.length, 8).setValues(dadosFiltrados);
      abaDestino.getRange(2, 6, dadosFiltrados.length, 1).setNumberFormat("dd/mm/yyyy");
      abaDestino.getRange(2, 8, dadosFiltrados.length, 1).setNumberFormat("R$ #,##0.00");
      
      ui.alert("Sucesso", `${dadosFiltrados.length} registros encontrados para o período.`, ui.ButtonSet.OK);
    } else {
      ui.alert("Atenção", "Nenhum dado encontrado para este período.", ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert("Erro: " + e.message);
  }
}

function gerarRelatorioResiduo10() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const abaComp = ss.getSheetByName(CONFIG.destino.nomeAba);
    let abaRes = ss.getSheetByName("Residuo10");
    if (!abaComp) throw new Error("Aba Compilados faltando.");
    if (!abaRes) abaRes = ss.insertSheet("Residuo10");

    const dadosEnt = obterDadosEntradasGlobal();
    const mapaVal = new Map();
    dadosEnt.forEach(l => {
      const valUnit = parseFloat(l[17]) || 0; 
      if (_norm(l[0]) && _norm(l[2])) mapaVal.set(`${_norm(l[0])}|${_norm(l[2])}`, valUnit);
    });
    const lrComp = abaComp.getLastRow();
    const dadosSaida = [];
    let totalGeral = 0;
    if (lrComp >= 2) {
      abaComp.getRange(2, 1, lrComp - 1, 20).getValues().forEach(l => {
        if (_norm(l[18]) === "RESÍDUO 10%") {
           const vUnit = mapaVal.get(`${_norm(l[0])}|${_norm(l[5])}`) || 0;
           const vTot = (parseFloat(l[16])||0) * vUnit;
           totalGeral += vTot;
           dadosSaida.push([l[0], l[6], l[5], l[19], "Resíduo 10%", l[16], vUnit, vTot]);
        }
      });
    }

    abaRes.clear();
    const cabecalho = ["Empenho", "Descrição", "Código", "Processo", "Status", "Qtd (Saldo)", "Vlr Unit.", "Vlr Total"];
    abaRes.getRange(1, 1, 1, 8).setValues([cabecalho]).setFontWeight("bold").setBackground("#cfe2f3");
    if (dadosSaida.length > 0) {
      abaRes.getRange(2, 1, dadosSaida.length, 8).setValues(dadosSaida);
      abaRes.getRange(2, 7, dadosSaida.length, 2).setNumberFormat("R$ #,##0.00");
      const linTot = dadosSaida.length + 2;
      abaRes.getRange(linTot, 1, 1, 7).merge().setValue("TOTAL GERAL").setHorizontalAlignment("right").setFontWeight("bold");
      abaRes.getRange(linTot, 8).setValue(totalGeral).setNumberFormat("R$ #,##0.00").setFontWeight("bold").setBackground("#d9ead3");
      abaRes.autoResizeColumns(1, 8);
      ui.alert("Sucesso", `${dadosSaida.length} itens gerados.`, ui.ButtonSet.OK);
    } else ui.alert("Info", "Nenhum resíduo encontrado.", ui.ButtonSet.OK);
  } catch (e) { ui.alert("Erro", e.message, ui.ButtonSet.OK);
  }
}

function gerarRelatorioAtrasos() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaComp = ss.getSheetByName("Compilados");
    let abaAtraso = ss.getSheetByName("Atraso>10");
    if (!abaComp) throw new Error("Aba Compilados não encontrada.");
    if (!abaAtraso) abaAtraso = ss.insertSheet("Atraso>10");
    abaAtraso.clear();

    const dados = abaComp.getRange("A1:U" + abaComp.getLastRow()).getValues();
    const head = dados[0], body = dados.slice(1);
    
    // Filtra e já mapeia para evitar erro de escopo do 'st'
    const dadosFilt = [];
    body.forEach(l => {
      const st = _norm(l[18]);
      let incluir = true;

      if (['RESÍDUO 10%', 'CONCLUÍDO', 'ELIMINADA'].includes(st)) incluir = false;
      
      const dias = parseAtrasoParaDias(l[17] ? String(l[17]) : "");
      if (dias < 11) incluir = false;
      
      const code = _norm(l[5]);
      if (['A3', 'A5', 'A7', 'C', 'D', 'P', 'S'].some(p => code.startsWith(p))) incluir = false;
      
      if (_norm(l[1]) === 'MAI') incluir = false;
      if (['A+', 'NÃO COBRAR', 'CÓDIGO MAI'].includes(_norm(l[12]))) incluir = false;
      if (_norm(l[6]).match(/REAGENTE|COMPRESSIVA|^PRÓTESE/)) incluir = false;
      
      if (incluir) {
        // Monta a linha com a observação tratada
        dadosFilt.push([
          l[0], l[4], l[5], l[6], l[8], l[15], l[16], l[17], l[19], l[20], 
          (st === 'RECEBIMENTO PROVISÓRIO' ? 'Recebimento Provisório' : '')
        ]);
      }
    });

    const headerNovo = [head[0], head[4], head[5], head[6], head[8], head[15], head[16], head[17], head[19], head[20], "Observação"];
    abaAtraso.getRange(1, 1).setValues([headerNovo]).setFontWeight('bold');
    
    if (dadosFilt.length > 0) {
      abaAtraso.getRange(2, 1, dadosFilt.length, 11).setValues(dadosFilt);
    }
    
    ui.alert("Relatório Gerado", `${dadosFilt.length} itens encontrados.`, ui.ButtonSet.OK);
  } catch (e) { ui.alert("Erro", e.message, ui.ButtonSet.OK);
  }
}

function processarRestosAPagar() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaComp = ss.getSheetByName("Compilados");
    const abaRestos = ss.getSheetByName("Restos a pagar");
    if (!abaComp || !abaRestos) throw new Error("Abas não encontradas.");
    const dados = abaComp.getRange("A2:T" + abaComp.getLastRow()).getValues();
    const empenhos = abaRestos.getRange("A2:A" + abaRestos.getLastRow()).getValues().flat().filter(String);
    if (!empenhos.length) throw new Error("Lista vazia.");
    const mapa = new Map();
    dados.forEach(l => {
      const e = _norm(l[0]);
      if(e) { if(!mapa.has(e)) mapa.set(e, []); mapa.get(e).push([l[4], l[5], l[6], l[16], l[18], l[19]]); }
    });
    let saida = [['Nº EMPENHO', 'FORNECEDOR', 'ITEM', 'DESCRIÇÃO', 'QTD RESIDUAL', 'STATUS', 'PROCESSO']];
    empenhos.forEach(eLongo => {
      let eCurto = null;
      const eStr = String(eLongo).trim();
      if(eStr.toUpperCase().includes('NE')) { const p = eStr.split(/NE/i); if(p.length>=2) eCurto = p[0].slice(-4)+p[1].slice(-4); }
      if(eCurto && mapa.has(eCurto)) mapa.get(eCurto).forEach(i => saida.push([eLongo, ...i]));
      else saida.push([eLongo, 'Nenhum item encontrado', '', '', '', '', '']);
    });
    abaRestos.clear();
    abaRestos.getRange(1, 1, saida.length, saida[0].length).setValues(saida);
    ui.alert("Concluído", `Geradas ${saida.length-1} linhas.`, ui.ButtonSet.OK);
  } catch (e) { ui.alert("Erro", e.message, ui.ButtonSet.OK);
  }
}

function atualizarResumo() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaComp = ss.getSheetByName("Compilados");
    var abaRes = ss.getSheetByName("Resumo");
    if(!abaComp || !abaRes) throw new Error("Abas faltando.");

    const dadosEnt = obterDadosEntradasGlobal();
    var mapV = new Map();
    
    dadosEnt.forEach(l => { 
      if(l[0] && l[2]) mapV.set(_norm(l[0])+"|"+_norm(l[2]), parseFloat(l[17])||0); 
    });
    var dRes=[], sRes=0, dPen=[], sPen=0;
    abaComp.getRange("A2:S"+abaComp.getLastRow()).getValues().forEach(l => {
      var st = _norm(l[18]);
      if(st==="RESÍDUO 10%" || st==="PENDENTE") {
        var v = mapV.get(_norm(l[0])+"|"+_norm(l[5]))||0;
        var q = parseFloat(l[16])||0;
        var tot = q*v;
        var lin = [l[0], l[5], l[6], v, q, tot];
        if(st==="RESÍDUO 10%") { dRes.push(lin); sRes+=tot; } else { dPen.push(lin); sPen+=tot; }
      }
    });

    abaRes.clear();
    var h = ["Empenho", "Código", "Descrição", "Valor Unitário", "Quantidade", "Valor Total"];
    gerarBlocoDeRelatorio(abaRes, 1, h, dRes, sRes, "Total Resíduo");
    gerarBlocoDeRelatorio(abaRes, 8, h, dPen, sPen, "Total Pendente");
    ui.alert("Sucesso", ui.ButtonSet.OK);
  } catch(e) { ui.alert("Erro", e.message, ui.ButtonSet.OK); }
}

// =================================================================================
// FUNÇÃO ATUALIZADA: Mapeamento de Colunas C, D, E, H e K
// =================================================================================
function buscarDadosLista() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const abaLista = ss.getSheetByName("lista");
    const abaDados = ss.getSheetByName("dados");
    const abaCompilados = ss.getSheetByName("Compilados");
    const abaEstoque = ss.getSheetByName("Cont.Estoque");
    
    if (!abaLista || !abaDados) throw new Error("Abas 'lista' e/ou 'dados' não encontradas.");
    
    // Leitura dos códigos de entrada
    const ultLinhaLista = abaLista.getLastRow();
    if (ultLinhaLista < 2) { 
      ui.alert("Atenção", "Digite códigos na coluna A da guia 'lista'.", ui.ButtonSet.OK); 
      return;
    }

    const codigosInput = abaLista.getRange(2, 1, ultLinhaLista - 1, 1).getValues().flat();
    const setMeusCodigos = new Set(codigosInput.map(c => _norm(c)).filter(c => c !== ""));

    // Mapeamentos Auxiliares
    // Leitura de Compilados expandida até a Coluna T (índice 20)
    const mapaCompilados = new Map();
    if (abaCompilados) {
      const lastRowComp = abaCompilados.getLastRow();
      if (lastRowComp >= 2) {
        // Lê A(1) até T(20)
        const dadosComp = abaCompilados.getRange(2, 1, lastRowComp - 1, 20).getValues();
        dadosComp.forEach(row => {
          let empenho = _norm(row[0]); 
          let codigo = _norm(row[5]);
          if (empenho && codigo) {
            mapaCompilados.set(`${empenho}|${codigo}`, { 
              colunaI_Comp: row[8],   // Coluna I (índice 8)
              saldo: row[16],         // Coluna Q (índice 16)
              status: row[18],        // Coluna S (índice 18)
              colunaT_Comp: row[19]   // Coluna T (índice 19)
            });
          }
        });
      }
    }
    
    const mapaEstoque = new Map();
    if (abaEstoque && abaEstoque.getLastRow() > 1) {
      const dadosEst = abaEstoque.getRange(2, 1, abaEstoque.getLastRow() - 1, 10).getValues();
      dadosEst.forEach(r => {
        const cod = _norm(r[1]); 
        if (cod) {
          mapaEstoque.set(cod, {
            estoque: r[5],   
            status: r[7],    
            previsao: r[9]   
          });
        }
      });
    }

    // Leitura: Da linha 5 até o fim, Pegando 44 colunas (A até AR)
    const ultLinhaDados = abaDados.getLastRow();
    if (ultLinhaDados < 5) {
       ui.alert("Aba 'dados' não possui registros suficientes (começando da linha 5).");
       return;
    }
    const valoresDados = abaDados.getRange(5, 1, ultLinhaDados - 4, 44).getValues();
    
    let dadosSaida = [];

    for (let i = 0; i < valoresDados.length; i++) {
      const linha = valoresDados[i];
      // [0]=Col A, [1]=Col B, ...
      const codAtual = _norm(linha[1]); // Coluna B
      
      if (!codAtual) continue; 

      if (setMeusCodigos.has(codAtual)) {
        
        // --- MAPEAMENTO SOLICITADO ---
        
        // 1. Coluna C da Lista -> Coluna G de Dados (Índice 6)
        let colC_Lista = linha[6];

        // 2. Coluna D da Lista -> Coluna U de Dados (Índice 20) com Regra "Começa com 1"
        let rawU = String(linha[20]).trim();
        let colD_Lista = "";
        if (rawU.startsWith("1")) {
            colD_Lista = rawU;
        }

        // 3. Coluna E da Lista -> Coluna V de Dados (Índice 21)
        let colE_Lista = linha[21];

        // --- DADOS COMPILADOS ---
        let empenhoDados = _norm(linha[10]); // Coluna K (Índice 10)
        let infoExtra = mapaCompilados.get(`${empenhoDados}|${codAtual}`) || { 
            colunaI_Comp: "", 
            saldo: "", 
            status: "", 
            colunaT_Comp: "" 
        };

        // 4. Coluna H da Lista -> Coluna I de Compilados
        let colH_Lista = infoExtra.colunaI_Comp;

        // 5. Coluna K da Lista -> Coluna T de Compilados
        let colK_Lista = infoExtra.colunaT_Comp;

        // --- CÁLCULOS EXTRAS MANTIDOS ---
        let infoEst = mapaEstoque.get(codAtual) || { estoque: "ND", status: "-", previsao: "-" };
        let cmm = parseFloat(linha[7]) || 0; // Coluna H
        const estoqueAtual = parseFloat(linha[6]) || 0; 
        const saldoPendente = parseFloat(infoExtra.saldo) || 0; 
        const qtdAE = parseFloat(colE_Lista) || 0;
        
        const estoqueTotalProjetado = estoqueAtual + saldoPendente + qtdAE;
        let previsaoProjetada = "-";
        if (cmm > 0) {
           const consumoDiario = cmm / 30;
           const diasCobertura = Math.floor(estoqueTotalProjetado / consumoDiario);
           const hoje = new Date();
           hoje.setDate(hoje.getDate() + diasCobertura);
           if (diasCobertura > 1000) previsaoProjetada = "Longo Prazo"; 
           else previsaoProjetada = hoje;
        } else {
           previsaoProjetada = (estoqueTotalProjetado > 0) ? "Sem Consumo" : "Zerado";
        }

        let textoColN = "";
        if (infoEst.status) textoColN += infoEst.status;
        if (infoEst.previsao) {
          if (infoEst.previsao instanceof Date) {
            textoColN += " | " + Utilities.formatDate(infoEst.previsao, "GMT-3", "dd/MM/yyyy");
          } else {
            textoColN += " | " + infoEst.previsao;
          }
        }
        
        // Montagem da Linha de Saída
        dadosSaida.push([
          codAtual,      // A (Cod)
          linha[2],      // B (Desc)
          colC_Lista,    // C (Dados Col G)
          colD_Lista,    // D (Dados Col U filtrado)
          colE_Lista,    // E (Dados Col V)
          linha[17],     // F (Preço Unitário - Mantido)
          linha[10],     // G (Empenho - Mantido)
          colH_Lista,    // H (Compilados Col I)
          infoExtra.saldo,  // I (Saldo - Mantido)
          infoExtra.status, // J (Status - Mantido)
          colK_Lista,    // K (Compilados Col T)
          cmm,           // L
          previsaoProjetada, // M 
          textoColN          // N 
        ]);
      }
    }
    
    abaLista.getRange("A2:N").clearContent();
    if (dadosSaida.length > 0) {
      abaLista.getRange(2, 1, dadosSaida.length, 14).setValues(dadosSaida);
      
      // Formatações
      abaLista.getRange(2, 9, dadosSaida.length, 1).setNumberFormat("#,##0");
      abaLista.getRange(2, 12, dadosSaida.length, 1).setNumberFormat("#,##0.00"); 
      abaLista.getRange(2, 13, dadosSaida.length, 1).setNumberFormat("dd/mm/yyyy"); 
      
      abaLista.getRange("M1").setValue("Previsão Projetada").setFontWeight("bold").setBackground("#d9ead3");
      abaLista.getRange("N1").setValue("Status Esgotamento (Atual)").setFontWeight("bold").setBackground("#f4cccc");
      ui.alert("Sucesso!", `${dadosSaida.length} linhas geradas.\nNovos mapeamentos aplicados.`, ui.ButtonSet.OK);
    } else ui.alert("Resultado", "Nenhum código correspondente encontrado na aba 'dados' (a partir da linha 5).", ui.ButtonSet.OK);
  } catch (e) { ui.alert("Erro", e.message, ui.ButtonSet.OK); }
}