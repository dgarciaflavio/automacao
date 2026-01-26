// =================================================================
// --- BLOCO 1: CONFIGURAÇÃO GLOBAL (SEGURA) ---
// =================================================================

var CONFIG = {
  ids: {
    materiais: PropertiesService.getScriptProperties().getProperty('ID_MATERIAIS'),
    medicamentos: PropertiesService.getScriptProperties().getProperty('ID_MEDICAMENTOS'),
    fonteDadosGeral: PropertiesService.getScriptProperties().getProperty('ID_FONTE_GERAL'),
    correcaoExterna: PropertiesService.getScriptProperties().getProperty('ID_CORRECAO'),
    compiladosLocal: PropertiesService.getScriptProperties().getProperty('ID_COMPILADOS'),
    painelEquipe: PropertiesService.getScriptProperties().getProperty('ID_PAINEL_EQUIPE')
  },
  abas: {
    materiais: "Material",
    medicamentos: "Medicamentos",
    entradas: "Entradas",
    empenhos: "Empenhos Enviados",
    recProvisorio: "Rec.Provisorio",
    fonteDadosNome: "dados",
    estoqueRemoto: "Cont.Estoque",
    historico: "Historico_BI"
  },
  destino: {
    nomeAba: "Compilados",
    colunaControleCobranca: 22
  },
  emails: {
    // Busca os e-mails seguros nas Propriedades do Script
    para: PropertiesService.getScriptProperties().getProperty('EMAIL_GESTAO_PARA'),
    copia: PropertiesService.getScriptProperties().getProperty('EMAIL_GESTAO_COPIA')
  },
  cores: {
    'CONCLUÍDO': '#d9ead3',                   
    'PENDENTE': '#f4cccc',                    
    'PENDENTE COM RESÍDUO': '#fff2cc',        
    'RESÍDUO 10%': '#cfe2f3',                 
    'SOLICITAR ASSOCIAÇÃO NO EMS': '#ffe599', 
    'RECEBIDO. FALTA ASSOCIAR EMS': '#b4a7d6',
    'ELIMINADA': '#999999',                   
    'RECEBIMENTO PROVISÓRIO': '#fce5cd',      // Laranja Claro (Padrão)
    'REC. PROV. / COM RESIDUO': '#f9cb9c',    // <--- NOVO: Laranja mais escuro/destaque
    'RECEBIDO A MAIOR': '#f6b26b', 
    'ALERTA_CRITICO': '#ea9999',
    'ALERTA_ATENCAO': '#ffe599',
    'ALERTA_OK': '#b6d7a8'
  }
};

// =================================================================
// --- BLOCO 1.5: VALIDAÇÃO DE SISTEMA ---
// =================================================================

function validarConexoes() {
  const ids = CONFIG.ids;
  const erros = [];
  const nomesAmigaveis = {
    materiais: "Planilha de Materiais",
    medicamentos: "Planilha de Medicamentos",
    fonteDadosGeral: "Fonte de Dados Geral (Importação)",
    correcaoExterna: "Planilha de Correção Externa",
    compiladosLocal: "Compilados (Local)"
  };

  for (let chave in ids) {
    const idAtual = ids[chave];
    if (!idAtual) continue;

    try {
      SpreadsheetApp.openById(idAtual);
    } catch (e) {
      const nomePlanilha = nomesAmigaveis[chave] || chave;
      erros.push(`❌ ${nomePlanilha}\n   (ID: ${idAtual})`);
    }
  }

  if (erros.length > 0) {
    throw new Error(
      "BLOQUEIO DE SEGURANÇA: Não foi possível acessar as seguintes planilhas:\n\n" + 
      erros.join("\n\n") + 
      "\n\nVerifique os IDs no Bloco 1 ou suas permissões de acesso antes de continuar."
    );
  }
}
