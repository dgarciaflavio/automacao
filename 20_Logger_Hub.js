// =================================================================
// --- BLOCO 20: LOGGER DE ACESSO AO HUB (ATIVIDADE) ---
// =================================================================

/**
 * ⚠️ INSTRUÇÃO DE INSTALAÇÃO:
 * Para que este log funcione, você deve configurar um ACIONADOR (TRIGGER) INSTALÁVEL.
 * * 1. No menu esquerdo do Apps Script, clique no ícone de relógio (Acionadores).
 * 2. Clique em "+ Adicionar Acionador".
 * 3. Escolha a função: "registrarAcessoHub".
 * 4. Selecione a origem do evento: "Da planilha".
 * 5. Selecione o tipo de evento: "Ao abrir".
 * 6. Salve e autorize.
 */

function registrarAcessoHub() {
  try {
    // IDs (Baseados no Config.js, mas repetidos aqui para segurança caso rode isolado)
    const ID_COMPILADOS = CONFIG.ids.compiladosLocal;
    const NOME_ABA_LOG = "Log_Acesso_Hub";
    
    // Identifica Usuário
    const emailUsuario = Session.getActiveUser().getEmail() || "Usuário Desconhecido/Externo";
    const dataAcesso = new Date();
    
    // Abre Planilha de Destino (Compilados)
    const ssComp = SpreadsheetApp.openById(ID_COMPILADOS);
    let abaLog = ssComp.getSheetByName(NOME_ABA_LOG);
    
    // Cria a aba se não existir
    if (!abaLog) {
      abaLog = ssComp.insertSheet(NOME_ABA_LOG);
      abaLog.appendRow(["Data/Hora Acesso", "E-mail Usuário", "Tipo Acesso"]);
      abaLog.getRange("A1:C1").setFontWeight("bold").setBackground("#d9d2e9");
      abaLog.setColumnWidth(1, 160);
      abaLog.setColumnWidth(2, 250);
    }
    
    // Registra o Acesso
    abaLog.appendRow([dataAcesso, emailUsuario, "Abertura do Hub"]);
    
    // Opcional: Limpeza de Logs antigos (mantém últimos 2000)
    const lastRow = abaLog.getLastRow();
    if (lastRow > 2000) {
       abaLog.deleteRows(2, lastRow - 2000);
    }
    
  } catch (e) {
    console.error("Erro ao registrar log de acesso: " + e.message);
    // Não usamos ui.alert aqui para não interromper o usuário se o log falhar
  }
}
