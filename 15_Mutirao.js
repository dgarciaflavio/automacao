// =================================================================
// --- BLOCO 15: STATUS REPORT (ATUALIZADO: FINANCEIRO + ÚLTIMO PREÇO) ---
// =================================================================

// ⚙️ CONFIGURAÇÃO GERAL
const NOME_ABA_TRABALHO = "Status Report"; 

// 📧 CONFIGURAÇÃO DE E-MAILS (SEGURA)
const CONFIG_EMAILS_MUTIRAO = {
  MODO_TESTE: true, 

  EMAIL_TESTE: PropertiesService.getScriptProperties().getProperty('EMAIL_MUTIRAO_TESTE'), 

  get LISTA_GERAL() {
    const listaTexto = PropertiesService.getScriptProperties().getProperty('EMAIL_MUTIRAO_LISTA');
    return listaTexto ? listaTexto.split(',').map(e => e.trim()) : [];
  }
};

/**
 * FUNÇÃO 1: Janela de Contexto (Menu)
 */
function processarMutirao() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Segoe UI', sans-serif; padding: 20px; background-color: #f3f3f3; }
      h3 { margin-top: 0; color: #333; }
      .container { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      label { display: block; margin-bottom: 12px; cursor: pointer; font-size: 15px; color: #444; }
      input[type="radio"] { transform: scale(1.3); margin-right: 10px; cursor: pointer; accent-color: #1c4587; }
      .btn { 
        background-color: #1c4587; color: white; border: none; 
        padding: 12px 0; width: 100%; border-radius: 5px; 
        font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 15px;
        transition: background 0.3s;
      }
      .btn:hover { background-color: #0f2e5e; }
      .loading { display: none; color: #666; font-size: 13px; text-align: center; margin-top: 10px; }
    </style>
    <div class="container">
      <h3>Definir Origem</h3>
      <p>Estes itens pertencem a qual grupo?</p>
      <form id="formOrigem">
        <label><input type="radio" name="opcao" value="MUTIRAO" checked> 🏃‍♂️ <b>Mutirão</b></label>
        <label><input type="radio" name="opcao" value="ACAO"> 🎯 <b>Grupo Ação</b></label>
        <button type="button" class="btn" onclick="enviar()">Confirmar</button>
        <div id="msg" class="loading">🔄 Processando Dados Financeiros e Estoque...</div>
      </form>
    </div>
    <script>
      function enviar() {
        var radios = document.getElementsByName('opcao');
        var selecionado = 'MUTIRAO';
        for (var i = 0; i < radios.length; i++) {
          if (radios[i].checked) { selecionado = radios[i].value; break; }
        }
        document.getElementById('msg').style.display = 'block';
        google.script.run
          .withSuccessHandler(function() { google.script.host.close(); })
          .withFailureHandler(function(e) { alert("Erro: " + e); google.script.host.close(); })
          .executarMutiraoComContexto(selecionado);
      }
    </script>
  `).setWidth(320).setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'Orquestrador de Estoque');
}

/**
 * FUNÇÃO 2: Motor Lógico Principal
 */
function executarMutiraoComContexto(tipoOrigem) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const abaMutirao = ss.getSheetByName(NOME_ABA_TRABALHO); 
    const abaDados = ss.getSheetByName("dados");
    const abaCompilados = ss.getSheetByName("Compilados");

    if (!abaMutirao) throw new Error(`A guia '${NOME_ABA_TRABALHO}' não foi encontrada.`);
    if (!abaDados || !abaCompilados) throw new Error("Abas de dados ('dados' ou 'Compilados') não encontradas.");

    const NOME_CONTEXTO = (tipoOrigem === "MUTIRAO") ? "MUTIRÃO" : "GRUPO AÇÃO";
    const COR_TITULO = (tipoOrigem === "MUTIRAO") ? "#1c4587" : "#cc0000"; 

    // --- 1. PRESERVAR DADOS MANUAIS (Obs e Qtd Solicitada) ---
    const lastRowMult = abaMutirao.getLastRow();
    if (lastRowMult < 2) { ui.alert(`A guia '${NOME_ABA_TRABALHO}' parece vazia.`); return; }

    const mapObservacoesSalvas = new Map();
    const mapQtdSolicitada = new Map();

    // Lê até a Coluna K (Obs) para garantir que pegamos tudo
    const dadosAtuais = abaMutirao.getRange(2, 1, lastRowMult - 1, 11).getValues();
    
    dadosAtuais.forEach(linha => {
      const cod = _norm(linha[0]);
      // Coluna D (Indice 3) é a Qtd Solicitada
      const qtd = linha[3]; 
      // Coluna K (Indice 10) é a Obs
      const obs = String(linha[10]).trim(); 
      
      if (cod) {
        if (obs) mapObservacoesSalvas.set(cod, obs);
        if (qtd !== "" && qtd != null) mapQtdSolicitada.set(cod, qtd);
      }
    });

    // --- 2. LEITURA DOS CÓDIGOS DE ENTRADA (Coluna A) ---
    const rangeEntrada = abaMutirao.getRange(2, 1, lastRowMult - 1, 1).getValues();
    const setCodigos = new Set();
    
    rangeEntrada.forEach(linha => {
      const cod = _norm(linha[0]);
      if (cod) setCodigos.add(cod);
    });

    if (setCodigos.size === 0) { ui.alert("Nenhum código válido na Coluna A."); return; }

    // --- 3. BUSCA PREÇO UNITÁRIO (ÚLTIMA ENTRADA) ---
    const mapaPrecos = _buscarUltimosPrecos(setCodigos);

    // --- 4. DADOS DE ESTOQUE (Aba 'dados') ---
    const lastRowDados = abaDados.getLastRow();
    const mapaDados = new Map();
    if (lastRowDados >= 5) {
      const vDados = abaDados.getRange(5, 1, lastRowDados - 4, 39).getValues();
      for (let i = 0; i < vDados.length; i++) {
        const linha = vDados[i];
        const cod = _norm(linha[1]); 
        if (cod && setCodigos.has(cod)) {
           if (!mapaDados.has(cod)) {
             mapaDados.set(cod, { 
               estoqueTotal: 0, aeSet: new Set(), notesSet: new Set(),
               processosSet: new Set(), descricaoSet: new Set(),
               saldoAta: 0, vencAtaSet: new Set()
             });
           }
           const d = mapaDados.get(cod);
           d.estoqueTotal += (parseFloat(linha[6]) || 0); 
           if (linha[2]) d.descricaoSet.add(String(linha[2]).trim());
           
           if (linha[15]) {
             const valP = String(linha[15]).trim();
             if (valP.startsWith("1")) d.aeSet.add(valP); else d.notesSet.add(valP);
           }
           if (linha[20]) {
             const valU = String(linha[20]).trim();
             if (valU.startsWith("6")) d.notesSet.add(valU); else d.aeSet.add(valU);
           }

           if (linha[38]) d.processosSet.add(String(linha[38]).trim());
           const saldoAtual = parseFloat(linha[33]) || 0;
           if (saldoAtual > d.saldoAta) d.saldoAta = saldoAtual;
           if (linha[19]) d.vencAtaSet.add(linha[19]);
        }
      }
    }

    // --- 5. COMPILADOS (Empenhos) ---
    const lastRowComp = abaCompilados.getLastRow();
    const mapaEmpenhos = new Map();
    if (lastRowComp >= 2) {
      const vComp = abaCompilados.getRange(2, 1, lastRowComp - 1, 19).getValues();
      for (let i = 0; i < vComp.length; i++) {
        const cod = _norm(vComp[i][5]);
        const status = _norm(vComp[i][18]);
        if (cod && setCodigos.has(cod) && (status.includes("PENDENTE") || status.includes("RESÍDUO"))) {
           if (!mapaEmpenhos.has(cod)) mapaEmpenhos.set(cod, new Set());
           mapaEmpenhos.get(cod).add(`${vComp[i][0]} (${status})`);
        }
      }
    }

    // --- 6. EXTERNA (Planejadores) ---
    const mapaPlanejadores = new Map();
    try {
      const ssEquipe = SpreadsheetApp.openById(CONFIG.ids.painelEquipe);
      const abasEquipe = ssEquipe.getSheets();
      const ignorarAbas = ["COAGE", "ESPELHOBASE", "CONFIGCMM", "CONFIG_EQUIPE", "DASHBOARD", "RESUMO", "PÁGINA1"];
      abasEquipe.forEach(aba => {
        const nomeAba = aba.getName();
        if (ignorarAbas.includes(nomeAba.toUpperCase())) return;
        if (aba.getLastRow() < 2) return;
        const dadosCodigos = aba.getRange(2, 2, aba.getLastRow() - 1, 1).getValues();
        for (let r = 0; r < dadosCodigos.length; r++) {
          const codItem = _norm(dadosCodigos[r][0]);
          if (codItem && setCodigos.has(codItem)) {
            if (!mapaPlanejadores.has(codItem)) mapaPlanejadores.set(codItem, new Set());
            mapaPlanejadores.get(codItem).add(nomeAba);
          }
        }
      });
    } catch (e) { console.warn("Erro busca externa: " + e.message); }

    // --- 7. MONTAGEM DOS DADOS ---
    const outputDados = [];      
    const outputDescricoes = []; 
    const outputPrecos = [];
    const outputQtdSolicitada = [];
    const outputFormulasTotal = [];
    const outputObservacoes = []; 
    const listaCompletaEmail = []; 

    rangeEntrada.forEach((linhaInput, index) => {
      const codAtual = _norm(linhaInput[0]);
      if (!codAtual) { 
        // Empurra linhas vazias para manter alinhamento
        outputDescricoes.push([""]);
        outputPrecos.push([""]);
        outputQtdSolicitada.push([""]);
        outputFormulasTotal.push([""]);
        outputDados.push(["", "", "", "", ""]);
        outputObservacoes.push([""]);
        return; 
      }

      const info = mapaDados.get(codAtual) || { 
        estoqueTotal: 0, aeSet: new Set(), notesSet: new Set(),
        processosSet: new Set(), descricaoSet: new Set(), saldoAta: 0, vencAtaSet: new Set()
      };
      
      let descricaoFinal = Array.from(info.descricaoSet)[0] || "Descrição não encontrada";
      
      // Valores Financeiros
      const precoUnit = mapaPrecos.get(codAtual) || 0;
      const qtdSol = mapQtdSolicitada.has(codAtual) ? mapQtdSolicitada.get(codAtual) : "";
      
      // Fórmula OnEdit (Coluna E = C * D)
      const linhaPlanilha = index + 2; // +2 porque começa na linha 2
      const formulaTotal = `=IF(ISNUMBER(D${linhaPlanilha}); C${linhaPlanilha}*D${linhaPlanilha}; 0)`;

      const txtAE = Array.from(info.aeSet).join("\n");
      const txtNotes = Array.from(info.notesSet).join("\n");
      const txtProc = Array.from(info.processosSet).join("\n");
      const setEmp = mapaEmpenhos.get(codAtual) || new Set();
      const txtEmp = Array.from(setEmp).join("\n");
      const setPlan = mapaPlanejadores.get(codAtual) || new Set();
      const txtPlan = setPlan.size > 0 ? Array.from(setPlan).join(" / ") : "Não Encontrado";

      // Formatação para Relatórios
      let txtAENotes = "";
      if (txtAE) txtAENotes += `📑 <b>AE:</b> ${txtAE}`;
      if (txtNotes) txtAENotes += (txtAENotes ? "<br>" : "") + `📝 <b>Notes:</b> ${txtNotes}`;
      if (!txtAENotes) txtAENotes = "-";

      let txtEmpenhoFinal = txtEmp ? `🚚 ${txtEmp}` : "-";
      let txtProcFinal = txtProc ? `📂 ${txtProc}` : "-";

      let txtVencimento = "";
      if (info.vencAtaSet.size > 0) {
        const datasFormatadas = Array.from(info.vencAtaSet).map(d => {
           return d instanceof Date ? d.toLocaleDateString() : String(d).trim();
        }).filter(d => d !== "");
        if (datasFormatadas.length > 0) txtVencimento = datasFormatadas[0];
      }
      
      let statusAta = "";
      if (info.saldoAta > 0 || (txtVencimento && txtVencimento.length > 5)) {
        statusAta = `✅ <b>Ata:</b> ${info.saldoAta.toLocaleString()} un<br>📅 <b>Vence:</b> ${txtVencimento}`;
      } else {
        statusAta = "🚫 Sem Ata";
      }

      let xlAta = (info.saldoAta > 0) ? `Saldo: ${info.saldoAta}\nVence: ${txtVencimento}` : "Sem Ata";
      let xlAENotes = (txtAE ? `[AE] ${txtAE}\n` : "") + (txtNotes ? `[NOTES] ${txtNotes}` : "");

      const obsFinal = mapObservacoesSalvas.get(codAtual) || "";
      
      outputDescricoes.push([descricaoFinal]); 
      outputPrecos.push([precoUnit]);
      outputQtdSolicitada.push([qtdSol]);
      outputFormulasTotal.push([formulaTotal]);
      
      // Colunas F a J
      outputDados.push([txtAE, txtEmp, info.estoqueTotal, txtProc, txtPlan]); 
      outputObservacoes.push([obsFinal]); 

      listaCompletaEmail.push({
        codigo: codAtual,
        descricao: descricaoFinal, 
        preco: precoUnit,
        qtdSol: qtdSol,
        estoque: info.estoqueTotal,
        ata: statusAta,            
        empenho: txtEmpenhoFinal,
        aeNotes: txtAENotes,
        processos: txtProcFinal,
        xlAta: xlAta,
        xlEmpenho: txtEmp,
        xlAENotes: xlAENotes,
        xlProcessos: txtProc
      });
    });

    // --- 8. ESCRITA NA PLANILHA ---
    // Limpa a área de dados (mantendo coluna A intacta)
    // Colunas B a K (Indices 2 a 11)
    abaMutirao.getRange(2, 2, abaMutirao.getMaxRows(), 10).clearContent();

    if (outputDados.length > 0) {
      const numLinhas = outputDados.length;

      // Coluna B: Descrição
      abaMutirao.getRange(2, 2, numLinhas, 1).setValues(outputDescricoes);
      
      // Coluna C: Valor Unitário
      abaMutirao.getRange(2, 3, numLinhas, 1).setValues(outputPrecos).setNumberFormat("R$ #,##0.00");
      
      // Coluna D: Qtd Solicitada
      abaMutirao.getRange(2, 4, numLinhas, 1).setValues(outputQtdSolicitada).setNumberFormat("#,##0");
      
      // Coluna E: Valor Total (Fórmula)
      abaMutirao.getRange(2, 5, numLinhas, 1).setFormulas(outputFormulasTotal).setNumberFormat("R$ #,##0.00").setFontWeight("bold");

      // Colunas F a J: Dados Gerais
      const rangeGeral = abaMutirao.getRange(2, 6, numLinhas, 5);
      rangeGeral.setValues(outputDados);
      rangeGeral.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      rangeGeral.setVerticalAlignment("middle");
      
      // Estoque (Coluna H -> Indice 8)
      abaMutirao.getRange(2, 8, numLinhas, 1).setNumberFormat("#,##0");

      // Coluna K: Observações
      abaMutirao.getRange(2, 11, numLinhas, 1).setValues(outputObservacoes);

      // --- CABEÇALHOS (GARANTIA) ---
      const headers = ["Código", "Descrição", "Valor Unitário", "Qtd. Total Solicitada", "Valor Total Solicitado", "AE / Notes", "Empenho", "Estoque", "Processos", "Planejador", "Observações"];
      abaMutirao.getRange(1, 1, 1, 11).setValues([headers])
        .setFontWeight("bold")
        .setBackground(COR_TITULO)
        .setFontColor("white")
        .setHorizontalAlignment("center");
        
      abaMutirao.setColumnWidth(2, 300); // Desc
      abaMutirao.setColumnWidth(3, 100); // Preço
      abaMutirao.setColumnWidth(4, 100); // Qtd
      abaMutirao.setColumnWidth(5, 120); // Total

      if (listaCompletaEmail.length > 0) {
        const resp = ui.alert(
          "Concluído", 
          `Status Report Atualizado!\nNovas Colunas Financeiras Adicionadas.\n\nDeseja enviar o E-MAIL (PDF + Excel)?`,
          ui.ButtonSet.YES_NO
        );

        if (resp === ui.Button.YES) {
          enviarEmailComAnexos(listaCompletaEmail, NOME_CONTEXTO, COR_TITULO);
        } else {
          ui.alert("Ok, planilha atualizada. E-mail não enviado.");
        }
      }
    }
  } catch (e) { ui.alert("Erro ao processar: " + e.message); }
}

/**
 * FUNÇÃO AUXILIAR: Busca Preços na Fonte Global
 */
function _buscarUltimosPrecos(setCodigos) {
  const mapa = new Map();
  try {
    const dados = obterDadosEntradasGlobal(); // Já vem do 03_Helpers.js
    
    // Varre todos os dados globais
    dados.forEach(linha => {
      // Coluna C (Index 2) = Código
      // Coluna I (Index 8) = Valor
      // Coluna O (Index 14) = Data Recebimento
      
      const cod = _norm(linha[2]);
      if (setCodigos.has(cod)) {
        const valor = parseFloat(linha[8]) || 0;
        const dataRaw = linha[14];
        let dataRec = null;

        if (dataRaw instanceof Date) {
          dataRec = dataRaw;
        } else if (typeof dataRaw === 'string') {
          dataRec = _parseDataSegura(dataRaw);
        }

        if (dataRec && valor > 0) {
          // Se já tem preço salvo, verifica se a data atual é mais recente
          if (mapa.has(cod)) {
            const anterior = mapa.get(cod);
            if (dataRec > anterior.data) {
              mapa.set(cod, { valor: valor, data: dataRec });
            }
          } else {
            mapa.set(cod, { valor: valor, data: dataRec });
          }
        }
      }
    });

  } catch (e) {
    console.error("Erro ao buscar preços: " + e.message);
  }

  // Retorna apenas Mapa simplificado: Cod -> Valor
  const mapaFinal = new Map();
  for (let [k, v] of mapa) {
    mapaFinal.set(k, v.valor);
  }
  return mapaFinal;
}

/**
 * FUNÇÃO 3: Gera PDF + Excel e envia E-mail (ATUALIZADA)
 */
function enviarEmailComAnexos(listaItens, nomeContexto, corTitulo) {
  const isTeste = CONFIG_EMAILS_MUTIRAO.MODO_TESTE;
  const destinatarios = isTeste ? CONFIG_EMAILS_MUTIRAO.EMAIL_TESTE : CONFIG_EMAILS_MUTIRAO.LISTA_GERAL.join(",");
  const assunto = `ℹ️ INFORMATIVO: Itens do ${nomeContexto} (${new Date().toLocaleDateString()})`;

  // --- 1. GERAÇÃO DO PDF (HTML) ---
  const tableHeader = `
    <thead>
      <tr>
        <th width="8%">Cód.</th>
        <th width="20%">Descrição</th>
        <th width="8%" style="text-align: right;">Unitário</th>
        <th width="5%" style="text-align: center;">Qtd Sol.</th>
        <th width="5%" style="text-align: center;">Estoque</th>
        <th width="12%">Cobertura</th>
        <th width="12%">Empenho</th>
        <th width="15%">AE / Notes</th>
        <th width="15%">Processos</th>
      </tr>
    </thead>
  `;

  let tableRows = "<tbody>";
  listaItens.forEach(item => {
    // Formatação de Moeda para o PDF
    const precoFmt = item.preco ? item.preco.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'}) : '-';
    
    tableRows += `
      <tr>
        <td><strong>${item.codigo}</strong></td>
        <td>${item.descricao.substring(0, 90)}</td>
        <td style="text-align: right;">${precoFmt}</td>
        <td style="text-align: center;"><b>${item.qtdSol || '-'}</b></td>
        <td style="text-align: center;">${item.estoque}</td>
        <td>${item.ata}</td>
        <td>${item.empenho.replace(/\n/g, "<br>")}</td>
        <td>${item.aeNotes}</td>
        <td>${item.processos.replace(/\n/g, "<br>")}</td>
      </tr>`;
  });
  tableRows += "</tbody>";

  let htmlContent = `
    <html>
    <head>
      <style>
        @page { size: landscape; margin: 10mm; }
        body { font-family: Arial, sans-serif; font-size: 10px; color: #333; }
        h2 { background-color: ${corTitulo}; color: white; padding: 8px; text-align: center; margin-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ccc; padding: 4px; text-align: left; vertical-align: top; }
        th { background-color: #f2f2f2; font-weight: bold; }
        thead { display: table-header-group; }
        tr { page-break-inside: avoid; }
      </style>
    </head>
    <body>
      <h2>Informativo: ${nomeContexto} (${new Date().toLocaleDateString()})</h2>
      <table>${tableHeader}${tableRows}</table>
      <p style="margin-top: 10px; font-size: 10px; color: #666; text-align: center;">Gerado automaticamente.</p>
    </body></html>`;

  const pdfBlob = HtmlService.createHtmlOutput(htmlContent).getAs(MimeType.PDF);
  const nomeArquivoBase = `Relatorio_${nomeContexto.replace(" ", "_")}_${new Date().toLocaleDateString()}`;
  pdfBlob.setName(`${nomeArquivoBase}.pdf`);

  // --- 2. GERAÇÃO DO EXCEL (.xlsx) ---
  let excelBlob = null;
  try {
    const tempSS = SpreadsheetApp.create("Temp_Export");
    const sheet = tempSS.getSheets()[0];
    
    const headers = ["Código", "Descrição", "Valor Unit.", "Qtd Solicitada", "Valor Total", "Estoque", "Cobertura/Ata", "Empenho", "AE / Notes", "Processos"];
    const dadosExcel = [headers];
    
    listaItens.forEach(item => {
      const total = (item.preco && item.qtdSol) ? (item.preco * item.qtdSol) : 0;
      dadosExcel.push([
        item.codigo,
        item.descricao,
        item.preco,
        item.qtdSol,
        total,
        item.estoque,
        item.xlAta,       
        item.xlEmpenho,   
        item.xlAENotes,   
        item.xlProcessos  
      ]);
    });

    sheet.getRange(1, 1, dadosExcel.length, headers.length).setValues(dadosExcel);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9d9d9");
    sheet.getRange(2, 3, dadosExcel.length, 1).setNumberFormat("R$ #,##0.00");
    sheet.getRange(2, 5, dadosExcel.length, 1).setNumberFormat("R$ #,##0.00");
    
    SpreadsheetApp.flush();

    const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    
    excelBlob = response.getBlob();
    excelBlob.setName(`${nomeArquivoBase}.xlsx`);

    DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  } catch (ex) {
    console.error("Erro ao gerar Excel: " + ex.message);
  }

  // --- 3. ENVIO DO EMAIL ---
  let htmlEmail = `
    <div style="font-family: Arial, sans-serif; color: #333;">
      <p>Prezados,</p>
      <p>Segue abaixo a relação atualizada do <strong>${nomeContexto}</strong> com valores atualizados.</p>
      <p><em>(Em anexo: Versão PDF para impressão e Planilha Excel para edição).</em></p>
      <hr>
      ${htmlContent} 
    </div>
  `;

  const anexos = [pdfBlob];
  if (excelBlob) anexos.push(excelBlob);

  try {
    MailApp.sendEmail({
      to: destinatarios,
      subject: assunto,
      htmlBody: htmlEmail, 
      attachments: anexos
    });
    SpreadsheetApp.getUi().alert(`E-mail enviado com PDF e Excel!`);
  } catch (e) {
    console.error(e);
    SpreadsheetApp.getUi().alert(`Erro no Email: ${e.message}`);
  }
}
