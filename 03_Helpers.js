// =================================================================
// --- BLOCO 3: FUNÇÕES AUXILIARES (HELPERS) - VERSÃO ROBUSTA ---
// =================================================================

function obterDadosEntradasGlobal() {
  try {
    // TENTATIVA DE RECUPERAÇÃO DE ID (BLINDAGEM CONTRA ERRO 'CONFIG NOT DEFINED')
    let idFonte;
    try {
      if (typeof CONFIG !== 'undefined' && CONFIG.ids) {
        idFonte = CONFIG.ids.fonteDadosGeral;
      }
    } catch (ignore) {}

    // Se o CONFIG falhou, busca direto nas Propriedades do Script (Fallback)
    if (!idFonte) {
      idFonte = PropertiesService.getScriptProperties().getProperty('ID_FONTE_GERAL');
    }

    if (!idFonte) {
      throw new Error("ID da Fonte de Dados Geral não encontrado (CONFIG ou ScriptProperties).");
    }

    const ssOrigem = SpreadsheetApp.openById(idFonte);
    // Tenta pelo nome no CONFIG ou usa o padrão "dados"
    const nomeAba = (typeof CONFIG !== 'undefined' && CONFIG.abas) ? CONFIG.abas.fonteDadosNome : "dados";
    const abaOrigem = ssOrigem.getSheetByName(nomeAba);
    
    if (!abaOrigem) {
      throw new Error(`Aba '${nomeAba}' não encontrada na planilha fonte.`);
    }

    const dados = abaOrigem.getDataRange().getValues();
    if (dados.length > 0) {
      dados.shift(); // Remove cabeçalho
    }
    
    // --- OTIMIZAÇÃO DE MEMÓRIA (FILTRO >= 2023) ---
    const anoCorte = 2023;
    const dadosFiltrados = dados.filter(linha => {
      const data = linha[0]; // Coluna A (Data)
      if (!data) return false; 
      
      if (data instanceof Date) {
        return data.getFullYear() >= anoCorte;
      }
      return true; 
    });
    
    console.log(`Dados carregados: ${dados.length} linhas totais. Mantidas: ${dadosFiltrados.length} (>= ${anoCorte}).`);
    return dadosFiltrados;

  } catch (e) {
    // O erro que você viu vinha daqui. Agora deve estar resolvido.
    throw new Error("Erro Crítico ao ler Fonte de Dados Global: " + e.message);
  }
}

function _norm(v) { 
  return v ? String(v).trim().toUpperCase() : ""; 
}

function _addDays(date, days) {
  let result = new Date(date);
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    added++;
  }
  return result;
}

function _diasParaTexto(total) {
  if (total < 0) total = 0;
  const m = Math.floor(total / 30);
  const d = total % 30;
  let txt = '';
  if (m > 0) txt += `${m} ${m > 1 ? 'meses' : 'mês'}`;
  if (d > 1) txt += (txt ? ' e ' : '') + `${d} dias`;
  else if (d === 1) txt += (txt ? ' e ' : '') + `${d} dia`;
  return txt || '0 dias';
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

function _calcularStatusUnificado(qEmpenhada, qSaidaOficial, saldoFisico, isRecProvisorio, statusAtual) {
  if (qEmpenhada > 0 && qSaidaOficial > qEmpenhada) return 'Recebido a Maior';
  if (qEmpenhada > 0 && qSaidaOficial === qEmpenhada) return 'Concluído'; 
  if (qEmpenhada === 0 && qSaidaOficial > 0) return 'Recebido. Falta associar EMS';
  if (qEmpenhada === 0) return 'Solicitar Associação no EMS';

  if (isRecProvisorio && qSaidaOficial === 0) {
      if (saldoFisico > 0 && saldoFisico <= (qEmpenhada * 0.10)) return 'Resíduo 10%';
      if (saldoFisico > 0) return 'Rec. Prov. / Com Residuo';
      return 'Recebimento Provisório';
  }

  if (statusAtual === 'Empenho não está na guia "Entradas"') return statusAtual;
  
  if (qSaidaOficial === 0 && qEmpenhada > 0 && saldoFisico === qEmpenhada) return 'Pendente';
  if (saldoFisico > (qEmpenhada * 0.10)) return 'Pendente com Resíduo';
  if (saldoFisico > 0 && saldoFisico <= (qEmpenhada * 0.10)) return 'Resíduo 10%';
  if (saldoFisico <= 0) return 'Resíduo 10%'; 

  return statusAtual || 'Pendente'; 
}
