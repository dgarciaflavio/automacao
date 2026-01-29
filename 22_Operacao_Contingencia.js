// =================================================================
// --- BLOCO 22: OPERAÇÃO CONTINGÊNCIA (ATUALIZADO: COM PREÇO UNITÁRIO) ---
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
    const setCodigosParaPreco = new Set(); // Conjunto para buscar preços

    const DIAS_CORTE = 91; 

    if (lastRowDados >= 5) {
      const vDados = abaDados.getRange(5, 1, lastRowDados - 4, 39).getValues();
      vDados.forEach(linha => {
        const estoqueDias = parseFloat(linha[8]) || 0; // Coluna I (Saldo em dias)
        
        let processoRaw = linha[38] ? String(linha[38]) : "";
        let processo = processoRaw.trim().replace(/\s+/g, " "); // Limpeza de espaços
        
        const codItem = _norm(linha[1]);

        // TRATATIVA: Se for vazio, traço ou zero -> "Item sem processo"
        if (!processo || processo === "-" || processo === "0" || processo === "") {
            processo = "Item sem processo";
        }

        if (codItem && estoqueDias <= DIAS_CORTE) {
          processosCriticos.add(processo);
          setCodigosParaPreco.add(codItem); // Adiciona para busca de preço
          
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

    // 4. BUSCAR ÚLTIMOS PREÇOS (Nova Lógica)
    toast("Buscando últimos preços praticados...");
    const mapaPrecos = _buscarUltimosPrecosContingencia(setCodigosParaPreco);

    // 5. BUSCAR EMPENHOS
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

    // 6. MONTAGEM DO RELATÓRIO
    const output = [];
    // Novo Cabeçalho com Valor Unitário
    const cabecalho = [
      "Processo SEI", 
      "Responsável (Usuário)", 
      "Código Item", 
      "Descrição", 
      "Valor Unit. (Último)", // <--- Nova Coluna E
      "Estoque (Dias)", 
      "Estoque (Qtd)", 
      "AE / Notes", 
      "Empenhos Pendentes", 
      "Status"
    ];

    processosCriticos.forEach(proc => {
      const itens = dadosItensPorProcesso.get(proc);
      
      let responsavel = "";
      if (proc === "Item sem processo") {
          responsavel = "Não mapeado";
      } else {
          responsavel = mapaProcUser.get(proc) || "Não mapeado";
      }

      itens.forEach((it) => {
        const empenhos = (mapaEmpenhos.get(it.codigo) || []).join("\n");
        const preco = mapaPrecos.get(it.codigo) || 0; // Pega o preço
        
        let infoAENotes = "";
        if (it.ae) infoAENotes += `AE: ${it.ae}`;
        if (it.notes) infoAENotes += (infoAENotes ? "\n" : "") + `Note: ${it.notes}`;

        output.push([
          proc,        
          responsavel, 
          it.codigo,
          it.descricao,
          preco, // <--- Valor Unitário
          it.estoqueDias,
          it.estoqueQtd,
          infoAENotes,
          empenhos,
          it.estoqueDias <= 30 ? "CRÍTICO" : "ALERTA"
        ]);
      });
    });

    // 7. ESCRITA NA ABA
    abaDestino.clear();
    // Cabeçalho roxo, fonte branca
    abaDestino.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
      .setFontWeight("bold").setBackground("#4c1130").setFontColor("white");

    if (output.length > 0) {
      const range = abaDestino.getRange(2, 1, output.length, cabecalho.length);
      range.setValues(output);
      range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      range.setVerticalAlignment("middle");
      
      // Formatação de Moeda na Coluna E (Índice 5)
      abaDestino.getRange(2, 5, output.length, 1).setNumberFormat("R$ #,##0.00");
      
      // Ajuste de Larguras
      abaDestino.autoResizeColumns(1, cabecalho.length);
      abaDestino.setColumnWidth(4, 300); // Descrição
      abaDestino.setColumnWidth(8, 150); // AE/Notes
      abaDestino.setColumnWidth(9, 200); // Empenhos
    }

    ui.alert(`Operação Contingência concluída!\nItens analisados com corte de ${DIAS_CORTE} dias.`);

  } catch (e) {
    ui.alert("Erro na Operação Contingência: " + e.message);
  }
}

/**
 * FUNÇÃO AUXILIAR: Busca Preços na Fonte Global (Específica para Contingência)
 */
function _buscarUltimosPrecosContingencia(setCodigos) {
  const mapa = new Map();
  try {
    const dados = obterDadosEntradasGlobal(); // Usa o helper global do arquivo 03_Helpers
    
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
          // Se já tem preço salvo, verifica se a data atual é mais recente (maior)
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
    console.error("Erro ao buscar preços (Contingência): " + e.message);
  }

  // Retorna apenas Mapa simplificado: Cod -> Valor
  const mapaFinal = new Map();
  for (let [k, v] of mapa) {
    mapaFinal.set(k, v.valor);
  }
  return mapaFinal;
}

function _norm(v) { 
  return v ? String(v).trim().toUpperCase() : ""; 
}

function _parseDataSegura(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;
  if (typeof valor === 'string') {
    const partes = valor.trim().split('/');
    if (partes.length === 3) {
      return new Date(partes[2], partes[1] - 1, partes[0]);
    }
  }
  return null; 
}
