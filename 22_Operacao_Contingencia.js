// =================================================================
// --- BLOCO 22: OPERAÇÃO CONTINGÊNCIA (ATUALIZADO: CORTE 91 DIAS) ---
// =================================================================

function executarOperacaoContingencia() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const abaDestino = ss.getSheetByName("OperacaoContingencia");
    const abaDados = ss.getSheetByName("dados");
    const abaCompilados = ss.getSheetByName("Compilados");
    const abaUserSEI = ss.getSheetByName("User_SEI");
    const abaProcSEI = ss.getSheetByName("Proc_SEI");

    if (!abaDestino) throw new Error("Aba 'OperacaoContingencia' não encontrada. Crie a aba antes de executar.");
    if (!abaDados || !abaCompilados) throw new Error("Abas de dados base não encontradas.");

    const toast = (msg) => ss.toast(msg, "Contingência", 5);
    toast("Identificando itens críticos...");

    // 1. MAPEAMENTO DE USUÁRIOS SEI (Login -> Nome)
    const mapaUsuarios = new Map();
    if (abaUserSEI) {
      const dadosUser = abaUserSEI.getDataRange().getValues();
      dadosUser.forEach(r => {
        const nome = String(r[0]).trim();
        const login = String(r[1]).trim();
        if (login) mapaUsuarios.set(login, nome);
      });
    }

    // 2. MAPEAMENTO DE PROCESSOS -> USUÁRIO (Processo -> Nome do Usuário)
    const mapaProcUser = new Map();
    if (abaProcSEI) {
      const dadosProc = abaProcSEI.getDataRange().getValues();
      dadosProc.forEach(r => {
        const numProc = String(r[0]).trim();
        const login = String(r[1]).trim();
        if (numProc) {
          const nomeCompleto = mapaUsuarios.get(login) || login || "Não atribuído";
          mapaProcUser.set(numProc, nomeCompleto);
        }
      });
    }

    // 3. IDENTIFICAR ITENS CRÍTICOS (Estoque <= 91 dias)
    const lastRowDados = abaDados.getLastRow();
    const processosCriticos = new Set();
    const dadosItensPorProcesso = new Map();

    // DEFINIÇÃO DO CORTE (Margem de segurança para decimais como 90,53)
    const DIAS_CORTE = 91; 

    if (lastRowDados >= 5) {
      const vDados = abaDados.getRange(5, 1, lastRowDados - 4, 39).getValues();
      vDados.forEach(linha => {
        const estoqueDias = parseFloat(linha[8]) || 0; // Coluna I (Saldo em dias)
        let processo = String(linha[38]).trim();       // Coluna AM (Processo SEI)
        const codItem = _norm(linha[1]);

        // TRATAMENTO PARA ITENS SEM PROCESSO
        if (!processo || processo === "-" || processo === "0") {
            processo = "Item sem processo";
        }

        // AGORA ACEITA <= 91 DIAS
        if (codItem && estoqueDias <= DIAS_CORTE) {
          processosCriticos.add(processo);
          
          if (!dadosItensPorProcesso.has(processo)) {
            dadosItensPorProcesso.set(processo, []);
          }
          dadosItensPorProcesso.get(processo).push({
            codigo: codItem,
            descricao: String(linha[2]).trim(),
            estoqueDias: estoqueDias,
            estoqueQtd: parseFloat(linha[6]) || 0,
            ae: String(linha[20]).trim().startsWith("1") ? String(linha[20]).trim() : "",
            notes: String(linha[15]).trim().startsWith("6") ? String(linha[15]).trim() : ""
          });
        }
      });
    }

    if (processosCriticos.size === 0) {
      ui.alert(`Nenhum item com estoque crítico (<= ${DIAS_CORTE} dias) foi encontrado.`);
      return;
    }

    // 4. BUSCAR EMPENHOS (Compilados)
    const lastRowComp = abaCompilados.getLastRow();
    const mapaEmpenhos = new Map();
    if (lastRowComp >= 2) {
      const vComp = abaCompilados.getRange(2, 1, lastRowComp - 1, 19).getValues();
      vComp.forEach(r => {
        const cod = _norm(r[5]);
        const status = _norm(r[18]);
        if (status.includes("PENDENTE") || status.includes("RESÍDUO")) {
          if (!mapaEmpenhos.has(cod)) mapaEmpenhos.set(cod, []);
          mapaEmpenhos.get(cod).push(`${r[0]} (${status})`);
        }
      });
    }

    // 5. MONTAGEM DO RELATÓRIO
    const output = [];
    const cabecalho = [
      "Processo SEI", "Responsável (Usuário)", "Código Item", "Descrição", 
      "Estoque (Dias)", "Estoque (Qtd)", "AE / Notes", "Empenhos Pendentes", "Status"
    ];

    processosCriticos.forEach(proc => {
      const itens = dadosItensPorProcesso.get(proc);
      
      // FORÇA "Não mapeado" SE FOR O CASO DE ITEM SEM PROCESSO
      let responsavel = "";
      if (proc === "Item sem processo") {
          responsavel = "Não mapeado";
      } else {
          responsavel = mapaProcUser.get(proc) || "Não mapeado";
      }

      itens.forEach((it, index) => {
        const empenhos = (mapaEmpenhos.get(it.codigo) || []).join("\n");
        let infoAENotes = "";
        if (it.ae) infoAENotes += `AE: ${it.ae}`;
        if (it.notes) infoAENotes += (infoAENotes ? "\n" : "") + `Note: ${it.notes}`;

        output.push([
          index === 0 ? proc : "", 
          index === 0 ? responsavel : "", 
          it.codigo,
          it.descricao,
          it.estoqueDias,
          it.estoqueQtd,
          infoAENotes,
          empenhos,
          it.estoqueDias <= 30 ? "CRÍTICO" : "ALERTA"
        ]);
      });
    });

    // 6. ESCRITA NA ABA
    abaDestino.clear();
    abaDestino.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
      .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");

    if (output.length > 0) {
      const range = abaDestino.getRange(2, 1, output.length, cabecalho.length);
      range.setValues(output);
      range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      range.setVerticalAlignment("middle");
      
      // Formatação Condicional de Status
      const cores = output.map(r => {
        const cor = r[8] === "CRÍTICO" ? "#ea9999" : "#ffe599";
        return new Array(cabecalho.length).fill(cor);
      });
      range.setBackgrounds(cores);
      
      abaDestino.autoResizeColumns(1, cabecalho.length);
      abaDestino.setColumnWidth(4, 300); // Descrição
      abaDestino.setColumnWidth(7, 150); // AE/Notes
      abaDestino.setColumnWidth(8, 200); // Empenhos
    }

    ui.alert(`Operação Contingência concluída!\nItens analisados com corte de ${DIAS_CORTE} dias.`);

  } catch (e) {
    ui.alert("Erro na Operação Contingência: " + e.message);
  }
}
