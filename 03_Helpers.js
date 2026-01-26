// =================================================================
// --- BLOCO 3: FUNÇÕES AUXILIARES (HELPERS) ---
// =================================================================

function obterDadosEntradasGlobal() {
  try {
    const ssOrigem = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
    const abaOrigem = ssOrigem.getSheetByName(CONFIG.abas.fonteDadosNome);
    
    if (!abaOrigem) {
      throw new Error(`Aba '${CONFIG.abas.fonteDadosNome}' não encontrada na planilha fonte.`);
    }

    const dados = abaOrigem.getDataRange().getValues();
    if (dados.length > 0) {
      dados.shift(); // Remove cabeçalho
    }
    
    // --- OTIMIZAÇÃO DE MEMÓRIA (FILTRO >= 2023) ---
    const anoCorte = 2023;
    const dadosFiltrados = dados.filter(linha => {
      const data = linha[0]; // Coluna A (Data)
      if (!data) return false; // Ignora linhas sem data
      
      // Se for objeto Date
      if (data instanceof Date) {
        return data.getFullYear() >= anoCorte;
      }
      return true; // Se não for data (texto), mantém por segurança para não perder cabeçalhos extras ou erros
    });
    
    console.log(`Dados carregados: ${dados.length} linhas totais. Mantidas: ${dadosFiltrados.length} (>= ${anoCorte}).`);
    return dadosFiltrados;

  } catch (e) {
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

// =================================================================
// --- FUNÇÃO RECUPERADA: CÁLCULO DE STATUS (REGRA ESTRITA L=0) ---
// =================================================================
function _calcularStatusUnificado(qEmpenhada, qSaidaOficial, saldoFisico, isRecProvisorio, statusAtual) {
  
  // 1. NOVO ALERTA (PRIORIDADE 1): Se tem empenho E saiu mais do que o empenhado -> PROBLEMA
  if (qEmpenhada > 0 && qSaidaOficial > qEmpenhada) return 'Recebido a Maior';

  // 2. Se tem empenho E a saída OFICIAL bateu EXATAMENTE o empenhado -> SUCESSO
  if (qEmpenhada > 0 && qSaidaOficial === qEmpenhada) return 'Concluído'; 

  // 3. Se NÃO tem empenho (qE=0) mas tem saída (qS>0) -> ERRO DE CADASTRO
  if (qEmpenhada === 0 && qSaidaOficial > 0) return 'Recebido. Falta associar EMS';

  // 4. Se não tem empenho nem saída -> ITEM FANTASMA
  if (qEmpenhada === 0) return 'Solicitar Associação no EMS';

  // 5. LÓGICA DE RECEBIMENTO PROVISÓRIO (REGRA DE OURO: Só se Oficial for ZERO)
  // Se entrou qualquer coisa no oficial (qSaidaOficial > 0), NÃO É MAIS PROVISÓRIO.
  if (isRecProvisorio && qSaidaOficial === 0) {
      // Se houver saldo físico irrelevante (<= 10%), marca como Resíduo 10% (evita cobrança)
      if (saldoFisico > 0 && saldoFisico <= (qEmpenhada * 0.10)) {
          return 'Resíduo 10%';
      }
      // Se tiver saldo físico relevante, marca como Provisório Parcial
      if (saldoFisico > 0) {
          return 'Rec. Prov. / Com Residuo';
      }
      // Se está 100% entregue no físico (mas oficial ainda é 0)
      return 'Recebimento Provisório';
  }

  // Mantém status antigo se não achar na fonte
  if (statusAtual === 'Empenho não está na guia "Entradas"') return statusAtual;
  
  // 6. Regras de Pendência Padrão
  // Aqui usamos o Saldo Físico Real (considerando o que já chegou no provisório)
  if (qSaidaOficial === 0 && qEmpenhada > 0 && saldoFisico === qEmpenhada) return 'Pendente';
  
  // Se sobrou saldo, verifica se é resíduo técnico (<= 10%) ou pendência real
  if (saldoFisico > (qEmpenhada * 0.10)) return 'Pendente com Resíduo';
  if (saldoFisico > 0 && saldoFisico <= (qEmpenhada * 0.10)) return 'Resíduo 10%';
  
  // Se saldo físico é 0 ou negativo, mas não caiu no Concluído acima (ex: oficial < empenho, mas fisico >= empenho)
  // Significa que fisicamente está ok, mas falta nota oficial.
  if (saldoFisico <= 0) return 'Resíduo 10%'; // Ou Concluído Físico

  // Fallback
  return statusAtual || 'Pendente'; 
}