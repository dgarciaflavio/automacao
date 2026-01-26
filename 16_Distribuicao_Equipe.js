// =================================================================
// --- BLOCO 16: DISTRIBUIÇÃO POR EQUIPE (COM COLUNA CLASSIFICAÇÃO SEPARADA) ---
// =================================================================

/**
 * 1. FUNÇÃO DE INTERFACE: Abre o pop-up com checkboxes
 */
function abrirSeletorPlanejadores() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Segoe UI', sans-serif; padding: 15px; background-color: #f3f3f3; }
      h3 { margin-top: 0; color: #333; font-size: 16px; }
      .container { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      .checkbox-group { display: flex; flex-direction: column; gap: 10px; margin-bottom: 15px; }
      label { cursor: pointer; font-size: 14px; color: #444; display: flex; align-items: center; }
      input[type="checkbox"] { transform: scale(1.2); margin-right: 10px; accent-color: #1c4587; }
      
      .actions { display: flex; gap: 10px; margin-bottom: 15px; }
      .small-btn { background: #ddd; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px; }
      .small-btn:hover { background: #ccc; }

      .btn { 
        background-color: #1c4587; color: white; border: none; 
        padding: 12px 0; width: 100%; border-radius: 5px; 
        font-size: 16px; font-weight: bold; cursor: pointer; 
        transition: background 0.3s;
      }
      .btn:hover { background-color: #0f2e5e; }
      .loading { display: none; color: #666; font-size: 13px; text-align: center; margin-top: 10px; }
    </style>

    <div class="container">
      <h3>Selecionar Planejadores</h3>
      
      <div class="actions">
        <button type="button" class="small-btn" onclick="toggle(true)">Marcar Todos</button>
        <button type="button" class="small-btn" onclick="toggle(false)">Desmarcar</button>
      </div>

      <div class="checkbox-group">
        <label><input type="checkbox" name="planner" value="Bianca" checked> 👤 Bianca</label>
        <label><input type="checkbox" name="planner" value="Katia" checked> 👤 Katia</label>
        <label><input type="checkbox" name="planner" value="Leonardo" checked> 👤 Leonardo</label>
        <label><input type="checkbox" name="planner" value="Moises" checked> 👤 Moises</label>
        <label><input type="checkbox" name="planner" value="Rafaelle" checked> 👤 Rafaelle</label>
        <label><input type="checkbox" name="planner" value="Luciana" checked> 👤 Luciana</label>
      </div>
      
      <button type="button" class="btn" onclick="enviar()">Atualizar Selecionados</button>
      <div id="msg" class="loading">🔄 Processando... aguarde.</div>
    </div>

    <script>
      function toggle(estado) {
        const boxes = document.getElementsByName('planner');
        for(let box of boxes) box.checked = estado;
      }

      function enviar() {
        const selecionados = [];
        const boxes = document.getElementsByName('planner');
        for (let box of boxes) {
          if (box.checked) selecionados.push(box.value);
        }

        if (selecionados.length === 0) {
          alert("Selecione pelo menos um planejador.");
          return;
        }

        document.getElementById('msg').style.display = 'block';
        
        google.script.run
          .withSuccessHandler(function() { google.script.host.close(); })
          .withFailureHandler(function(e) { alert("Erro: " + e); google.script.host.close(); })
          .processarSelecaoPlanejadores(selecionados);
      }
    </script>
  `).setWidth(300).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, 'Distribuição de Equipe');
}

/**
 * 2. FUNÇÃO INTERMEDIÁRIA
 */
function processarSelecaoPlanejadores(listaNomes) {
  distribuirDadosPorEquipe(listaNomes);
}

/**
 * 3. MOTOR PRINCIPAL (GLOBAL CHANGE: COLUNA CLASSIFICAÇÃO)
 * @param {Array|string|null} alvos - Lista de nomes, nome único ou null (todos)
 */
function distribuirDadosPorEquipe(alvos = null) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
  try {
    // 0. CARREGAMENTO E CONFIGURAÇÃO
    const regras = carregarRegrasDistribuicao(ss);
    // USE AS VARIÁVEIS GLOBAIS
    const idHistorico = CONFIG.ids.fonteDadosGeral; 
    const idDadosAtuais = CONFIG.ids.compiladosLocal;
    const idPainelDestino = CONFIG.ids.painelEquipe;

    // 1. CÁLCULO DO CMA HISTÓRICO
    const mapCMA = new Map();
    const anosValidos = [2022, 2023, 2025]; 
    const historico = {};

    try {
      const ssHist = SpreadsheetApp.openById(idHistorico);
      const abaHist = ssHist.getSheetByName("dados"); 
      if (abaHist && abaHist.getLastRow() >= 2) {
          abaHist.getRange(2, 1, abaHist.getLastRow() - 1, 19).getValues().forEach(r => {
            const cod = _norm(r[2]); 
            const statusN = String(r[13]).trim().toUpperCase(); 
            const rawData = r[16];   
            const qtd = parseFloat(r[18]) || 0; 
            
            if (cod && statusN === "RECEBIDA") {
              const ano = _obterAnoSeguro(rawData);
              if (ano && anosValidos.includes(ano)) {
                historico[cod] = (historico[cod] || 0) + qtd;
              }
            }
          });
      }
      
      const abaOutras = ssHist.getSheetByName("Outras Entradas");
      if (abaOutras && abaOutras.getLastRow() >= 2) {
          abaOutras.getRange(2, 1, abaOutras.getLastRow() - 1, 13).getValues().forEach(r => {
            const dataOutra = r[1]; 
            const codOutra = _norm(r[11]); 
            const qtdOutra = parseFloat(r[12]) || 0; 
            if (codOutra) {
              const ano = _obterAnoSeguro(dataOutra);
              if (ano && anosValidos.includes(ano)) {
                historico[codOutra] = (historico[codOutra] || 0) + qtdOutra;
              }
            }
          });
      }
    } catch (e) { console.warn("Erro Histórico: " + e.message); }

    for (const cod in historico) {
       const mediaAnual = historico[cod] / 3; 
       mapCMA.set(cod, (mediaAnual / 12) * 15);
    }

    // 2. DADOS ATUAIS
    const ssAtuais = SpreadsheetApp.openById(idDadosAtuais);
    const abaDadosAtuais = ssAtuais.getSheetByName("dados"); 
    const range = abaDadosAtuais.getRange(5, 1, abaDadosAtuais.getLastRow() - 4, 44);
    const vDadosRaw = range.getValues();
    const vDadosDisplay = range.getDisplayValues();

    const setOuvidorias = new Set();
    const abaOuv = ssAtuais.getSheetByName("Ouvidorias");
    if (abaOuv && abaOuv.getLastRow() > 1) {
       abaOuv.getRange(2, 1, abaOuv.getLastRow() - 1, 1).getValues().forEach(r => { if(r[0]) setOuvidorias.add(_norm(r[0])); });
    }

    const abaComp = ssAtuais.getSheetByName(CONFIG.destino.nomeAba); 
    const mapStatus = new Map();
    const mapQtdSaldo = new Map(); 
    if (abaComp && abaComp.getLastRow() > 1) {
      abaComp.getRange(2, 1, abaComp.getLastRow()-1, 21).getValues().forEach(r => {
        const cod = _norm(r[5]); 
        const empenho = _norm(r[0]); 
        if (cod) {
          mapStatus.set(cod, { empenho: r[0], forn: r[4], status: r[18] });
          mapQtdSaldo.set(`${empenho}-${cod}`, { qtd: parseFloat(r[8])||0, saldo: parseFloat(r[16])||0 });
        }
      });
    }

    let abaEst = ss.getSheetByName(CONFIG.abas.estoqueRemoto) || ssAtuais.getSheetByName(CONFIG.abas.estoqueRemoto);
    const mapAlerta = new Map();
    if (abaEst && abaEst.getLastRow() > 1) {
      abaEst.getRange(2, 1, abaEst.getLastRow()-1, 13).getValues().forEach(r => {
        if(r[1]) mapAlerta.set(_norm(r[1]), { alerta: r[7], previsao: r[9] });
      });
    }

    const ssDestino = SpreadsheetApp.openById(idPainelDestino);
    const mapCMMManual = new Map();
    const abaConfigCMM = ssDestino.getSheetByName("ConfigCMM");
    if (abaConfigCMM && abaConfigCMM.getLastRow() >= 2) {
       abaConfigCMM.getRange(2, 1, abaConfigCMM.getLastRow()-1, 2).getValues().forEach(r => {
          if(r[0]) mapCMMManual.set(_norm(r[0]), parseFloat(r[1]));
       });
    }

    // -----------------------------------------------------------------
    // 3. DISTRIBUIÇÃO E CÁLCULOS (COM CABECALHO ATUALIZADO)
    // -----------------------------------------------------------------
    const cabecalho = [
      "Tipo", "Item", "Descrição", "Fornecedor", "Empenho", "Status Empenho", 
      "Qtd Empenho", "Saldo Empenho", "Notes", "AE em andamento", "Quantidade AE", 
      "Saldo Ata", "Vencimento Ata", "Estoque", "Saldo Total F+V", "CMA Histórico (15 Meses)", 
      "CMM Atual", "Status Estoque", "Saldo em Dias", "Classificação", "Família", "Previsão Esgotamento", 
      "Processo SEI", "Sugestão Pedido (6 Meses)", "Prev. Esgotamento (Sugestão)", 
      "Ouvidoria", "Observações (Equipe)"
    ];
    // Nota: "Classificação" entrou no índice 19 (Coluna T)

    const listaFinal = []; 
    const hoje = new Date(); hoje.setHours(0,0,0,0); 

    for (let i = 0; i < vDadosDisplay.length; i++) {
      const rStr = vDadosDisplay[i];
      const rRaw = vDadosRaw[i];
      const cod = _norm(rStr[1]); 
      if (!cod) continue; 

      const itemDesc = rStr[2]; 
      const familiaLimpa = String(rStr[4]).replace(',', '.').trim(); 
      const isNumeric = /^\d/.test(cod);

      // Definição de Destinos
      const destinos = new Set();
      if (regras[cod]) destinos.add(regras[cod]);
      if (regras[familiaLimpa]) destinos.add(regras[familiaLimpa]);
      if (destinos.size === 0) destinos.add(isNumeric ? "Rafaelle" : "Triagem");
      
      const arrayDestinos = Array.from(destinos);
      const isConflito = (arrayDestinos.length > 1);

      // Cálculos
      const estoque = parseFloat(rRaw[6]) || 0; 
      let cmm = mapCMMManual.has(cod) ? mapCMMManual.get(cod) : (parseFloat(rRaw[7]) || 0);
      const cmaCalculado = mapCMA.get(cod) || 0;

      const notes = String(rStr[15]).replace(/[.,]/g, "").trim(); 
      const rawU = String(rStr[20]).trim();
      const aeAndamento = rawU.startsWith("1") ? rawU : "";
      const qtdAE = rRaw[21]; 
      const saldoAta = rRaw[33];       
      let vencAta = rStr[19];        
      const dataVencObj = _parseDataSegura(vencAta); 
      if (dataVencObj && dataVencObj < hoje) vencAta = ""; 
      const processoSEI = String(rStr[38]).trim(); 
      const statusEmp = mapStatus.get(cod) || { empenho: "---", forn: "---", status: "Sem Empenho" };
      
      const empenhoNum = _norm(statusEmp.empenho);
      const dadosQtdSaldo = mapQtdSaldo.get(`${empenhoNum}-${cod}`) || { qtd: 0, saldo: 0 };
      const alertaEst = mapAlerta.get(cod) || { alerta: "Não Calculado", previsao: "---" };
      const statusOuvidoria = setOuvidorias.has(cod) ? "OUVIDORIA" : "";

      const somaSaldos = (parseFloat(rRaw[31])||0) + (dadosQtdSaldo.saldo > 0 ? dadosQtdSaldo.saldo : 0);
      
      // NOVA LÓGICA DE SEPARAÇÃO (S e T)
      const saldoDiasValor = parseFloat(rRaw[8]) || 0; // Coluna I (Dados) -> Numérico
      const classificacao = String(rStr[9]).trim();    // Coluna J (Dados) -> Texto

      const consumoMensalHistorico = cmaCalculado / 15;
      const meta6Meses = consumoMensalHistorico * 6;
      let sugestaoPedido = Math.round(meta6Meses - estoque);
      if (sugestaoPedido < 0) sugestaoPedido = 0;

      let previsaoComSugestao = "Sem Consumo";
      if (consumoMensalHistorico > 0) {
        const diasFuturos = Math.floor((estoque + sugestaoPedido) / (consumoMensalHistorico / 30));
        const dataFutura = new Date();
        dataFutura.setDate(dataFutura.getDate() + diasFuturos);
        previsaoComSugestao = dataFutura;
      }
      const familiaExibicao = familiaLimpa || (isNumeric ? "Medicamento" : "SEM FAMÍLIA");

      const linhaDadosBase = [
          isNumeric ? "MEDICAMENTO" : "MATERIAL", cod, itemDesc, statusEmp.forn, statusEmp.empenho, statusEmp.status, 
          dadosQtdSaldo.qtd, dadosQtdSaldo.saldo, notes, aeAndamento, qtdAE, saldoAta, vencAta, estoque, somaSaldos, 
          cmaCalculado, cmm, alertaEst.alerta, saldoDiasValor, classificacao, familiaExibicao, alertaEst.previsao, 
          processoSEI, sugestaoPedido, previsaoComSugestao, statusOuvidoria
      ];

      arrayDestinos.forEach(dono => {
          let notaConflito = "";
          let corLinhaVermelha = false;

          if (isConflito) {
              corLinhaVermelha = true;
              const outros = arrayDestinos.filter(d => d !== dono).join(", ");
              notaConflito = `⚠️ COMPARTILHADO:\nItem também gerido por: ${outros}`;
          }

          listaFinal.push({
             responsavel: dono,
             codItem: cod,
             linhaDados: linhaDadosBase,
             isConflito: corLinhaVermelha,
             notaItem: notaConflito
          });
      });
    }

    // -----------------------------------------------------------------
    // 4. ESCRITA FILTRADA
    // -----------------------------------------------------------------
    const setResponsaveis = new Set(listaFinal.map(d => d.responsavel));
    let nomesAbas = [];
    let isFullUpdate = false;

    if (alvos === null) {
      nomesAbas = Array.from(setResponsaveis);
      isFullUpdate = true;
    } else if (Array.isArray(alvos)) {
      nomesAbas = alvos.filter(nome => setResponsaveis.has(nome));
      if (nomesAbas.length === 0) {
        ui.alert(`Atenção`, `Nenhum item encontrado para os planejadores selecionados.`, ui.ButtonSet.OK);
        return;
      }
    } else {
      if (setResponsaveis.has(alvos)) {
        nomesAbas = [alvos];
      }
    }

    const listaDadosCOAGE = []; 

    nomesAbas.forEach(nomeAba => {
      let aba = ssDestino.getSheetByName(nomeAba);
      if (!aba) aba = ssDestino.insertSheet(nomeAba);

      const mapObservacoes = new Map();
      if (aba.getLastRow() >= 2) {
        aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getValues().forEach(r => {
           if(r[1] && r[r.length-1]) mapObservacoes.set(_norm(r[1]), r[r.length-1]);
        });
      }

      const dadosParaEscrever = [];
      const metadadosLinha = []; 

      listaFinal.filter(d => d.responsavel === nomeAba).forEach(d => {
          const obsSalva = mapObservacoes.get(d.codItem) || "";
          const linha = [...d.linhaDados, obsSalva];
          dadosParaEscrever.push(linha);
          
          metadadosLinha.push({ conflito: d.isConflito, nota: d.notaItem });
      });
      
      if (isFullUpdate) {
         listaFinal.filter(d => d.responsavel === nomeAba).forEach(d => {
             const obsSalva = mapObservacoes.get(d.codItem) || "";
             listaDadosCOAGE.push([...d.linhaDados, nomeAba, obsSalva]);
         });
      }

      if (aba.getFilter()) aba.getFilter().remove();
      aba.clear(); 
      aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
         .setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");

      if (dadosParaEscrever.length > 0) {
        _formatarAbaDestino(aba, dadosParaEscrever, cabecalho, false, metadadosLinha);
      }
    });

    if (isFullUpdate) {
      let abaCoage = ssDestino.getSheetByName("COAGE");
      if (!abaCoage) abaCoage = ssDestino.insertSheet("COAGE");
      if (abaCoage.getFilter()) abaCoage.getFilter().remove();

      abaCoage.clear();
      const cabecalhoCOAGE = [...cabecalho]; cabecalhoCOAGE.splice(cabecalhoCOAGE.length - 1, 0, "Planejador"); 
      abaCoage.getRange(1, 1, 1, cabecalhoCOAGE.length).setValues([cabecalhoCOAGE]).setFontWeight("bold").setBackground("#4c1130").setFontColor("#ffffff");
      if (listaDadosCOAGE.length > 0) _formatarAbaDestino(abaCoage, listaDadosCOAGE, cabecalhoCOAGE, true);
    }

    const msgSucesso = isFullUpdate 
      ? "Painel de TODOS atualizado com sucesso!" 
      : `Painéis atualizados: ${nomesAbas.join(", ")}`;
      
    ui.alert("Sucesso", msgSucesso, ui.ButtonSet.OK);

  } catch (e) { ui.alert("Erro na Distribuição", e.message, ui.ButtonSet.OK); }
}

function _formatarAbaDestino(aba, dados, headers, isCoage = false, metadados = []) {
    const numLinhas = dados.length;
    
    aba.getRange(2, 1, numLinhas, headers.length).setValues(dados);

    // Formatações Atualizadas para o Novo Layout (Coluna T inserida)
    aba.getRange(2, 7, numLinhas, 2).setNumberFormat("#,##0"); // G, H
    aba.getRange(2, 9, numLinhas, 2).setNumberFormat("@");     // I, J
    aba.getRange(2, 11, numLinhas, 2).setNumberFormat("#,##0"); // K, L
    aba.getRange(2, 14, numLinhas, 2).setNumberFormat("#,##0"); // N, O
    aba.getRange(2, 16, numLinhas, 1).setNumberFormat("#,##0.00"); // P
    aba.getRange(2, 17, numLinhas, 1).setNumberFormat("#,##0"); // Q
    
    aba.getRange(2, 19, numLinhas, 1).setNumberFormat("#,##0.00"); // S (Saldo Dias Numérico)
    aba.getRange(2, 20, numLinhas, 1).setNumberFormat("@");        // T (Classificação Texto)
    aba.getRange(2, 21, numLinhas, 1).setNumberFormat("@");        // U (Família)

    aba.getRange(2, 13, numLinhas, 1).setNumberFormat("dd/mm/yyyy"); // M
    aba.getRange(2, 22, numLinhas, 1).setNumberFormat("dd/mm/yyyy"); // V (Previsão)
    aba.getRange(2, 25, numLinhas, 1).setNumberFormat("dd/mm/yyyy"); // Y (Prev Sugestão)
    
    aba.getRange(2, 24, numLinhas, 1).setNumberFormat("#,##0"); // X (Sugestão)
    
    aba.getRange(2, 26, numLinhas, 1).setNumberFormat("@").setHorizontalAlignment("center").setFontWeight("bold").setFontColor("#cc0000"); // Z (Ouvidoria)

    const backgroundsQ = []; const backgroundsW = []; 
    dados.forEach(l => {
        const statusEstoque = l[17]; 
        const isOuvidoria = l[25] === "OUVIDORIA"; // Índice mudou de 24 para 25
        backgroundsQ.push(statusEstoque === "Crítico" ? [CONFIG.cores.ALERTA_CRITICO] : statusEstoque === "Atenção" ? [CONFIG.cores.ALERTA_ATENCAO] : statusEstoque === "Suprimento Ok" ? [CONFIG.cores.ALERTA_OK] : [null]);
        backgroundsW.push(isOuvidoria ? ["#ea9999"] : [null]);
    });
    aba.getRange(2, 18, numLinhas, 1).setBackgrounds(backgroundsQ); 
    aba.getRange(2, 26, numLinhas, 1).setBackgrounds(backgroundsW); // Coluna Z (Ouvidoria)

    if (!isCoage && metadados.length > 0) {
        const rangeTotal = aba.getRange(2, 1, numLinhas, headers.length);
        const backgroundsTotal = rangeTotal.getBackgrounds();
        const notes = aba.getRange(2, 2, numLinhas, 1).getNotes(); 

        for (let i = 0; i < numLinhas; i++) {
            if (metadados[i] && metadados[i].conflito) {
                for (let j = 0; j < headers.length; j++) {
                    if (backgroundsTotal[i][j] === '#ffffff' || backgroundsTotal[i][j] === '#FFFFFF') { 
                        backgroundsTotal[i][j] = '#ea9999'; 
                    }
                }
                notes[i][0] = metadados[i].nota;
            }
        }
        rangeTotal.setBackgrounds(backgroundsTotal);
        aba.getRange(2, 2, numLinhas, 1).setNotes(notes);
    }

    for (let c = 0; c < headers.length; c++) {
       let largura = (headers[c].length * 10) + 35;
       if (largura < 70) largura = 70;
       aba.setColumnWidth(c + 1, largura);
    }
    
    if (aba.getFilter()) aba.getFilter().remove();
    aba.getDataRange().createFilter();
    aba.hideColumns(1); 
}

function _obterAnoSeguro(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor.getFullYear();
  const str = String(valor).trim();
  const partes = str.split('/');
  if (partes.length === 3) {
    const ano = parseInt(partes[2]);
    if (!isNaN(ano) && ano > 2000 && ano < 2100) return ano;
  }
  return null;
}

function carregarRegrasDistribuicao(ss) {
  let abaConfig = ss.getSheetByName("Config_Equipe");
  if (!abaConfig) {
    abaConfig = ss.insertSheet("Config_Equipe");
    abaConfig.getRange("A1:B1").setValues([["Família", "Responsável"]]).setFontWeight("bold").setBackground("#d9d2e9");
    return {};
  }
  const lastRow = abaConfig.getLastRow();
  if (lastRow < 2) return {};
  const dados = abaConfig.getRange(2, 1, lastRow - 1, 2).getValues();
  const regrasDinamicas = {};
  dados.forEach(linha => {
    const chave = _norm(linha[0]); 
    const responsavel = String(linha[1]).trim();
    if (chave && responsavel) regrasDinamicas[chave] = responsavel;
  });
  return regrasDinamicas;
}

function aplicarTravaSegurancaConfigCMM() {
  const ui = SpreadsheetApp.getUi();
  const idPainelDestino = "1Rc5fUr-nP3g8SU9083Y-N83QZtg2wMCMlmlwcXJYl9k"; 
  try {
    const ss = SpreadsheetApp.openById(idPainelDestino);
    let aba = ss.getSheetByName("ConfigCMM");
    if (!aba) { aba = ss.insertSheet("ConfigCMM"); aba.getRange("A1:B1").setValues([["Código do Item", "CMM Real (Manual)"]]); }
    const cell = aba.getRange("A2:A");
    const regra = SpreadsheetApp.newDataValidation().requireFormulaSatisfied('=COUNTIF($A$2:$A; A2)<=1').setAllowInvalid(false).build();
    cell.setDataValidation(regra);
    ui.alert("Sucesso!", "Trava aplicada.", ui.ButtonSet.OK);
  } catch (e) { ui.alert("Erro: " + e.message); }
}

// --- FUNÇÕES DE ATALHO PARA O MENU ---
function atualizarTodos() { distribuirDadosPorEquipe(null); }
function atualizarBianca() { distribuirDadosPorEquipe("Bianca"); }
function atualizarKatia() { distribuirDadosPorEquipe("Katia"); }
function atualizarLeonardo() { distribuirDadosPorEquipe("Leonardo"); }
function atualizarMoises() { distribuirDadosPorEquipe("Moises"); }
function atualizarRafaelle() { distribuirDadosPorEquipe("Rafaelle"); }
function atualizarLuciana() { distribuirDadosPorEquipe("Luciana"); }
