// =================================================================
// --- BLOCO 18: CONTROLE MOBILE (COM BARRA DE PROGRESSO) ---
// =================================================================

function monitorarPainelMobile(e) {
  // Configurações
  const NOME_ABA_MOBILE = "Painel_Mobile";
  const COLUNA_CHECKBOX = 2; // Coluna B
  const COLUNA_STATUS = 3;   // Coluna C
  
  // Validações
  const range = e.range;
  const sheet = range.getSheet();
  
  if (sheet.getName() !== NOME_ABA_MOBILE || range.getColumn() !== COLUNA_CHECKBOX) {
    return;
  }
  
  // Só executa se TRUE
  if (e.value !== "TRUE") return;

  const linha = range.getRow();
  const celulaStatus = sheet.getRange(linha, COLUNA_STATUS);
  
  // Limpa check e avisa início
  range.setValue(false);
  celulaStatus.setValue("🚀 Iniciando motor...");
  SpreadsheetApp.flush(); 

  try {
    switch (linha) {
      case 2: // CICLO COMPLETO (Fracionado para mostrar progresso)
        // Passo 1: Leitura
        celulaStatus.setValue("📥 1/5 Lendo Dados...");
        SpreadsheetApp.flush();
        const dados = obterDadosEntradasGlobal();
        
        // Passo 2: Materiais
        celulaStatus.setValue("📦 2/5 Proc. Materiais...");
        SpreadsheetApp.flush();
        processarMateriaisRemoto(dados);
        
        // Passo 3: Medicamentos
        celulaStatus.setValue("💊 3/5 Proc. Meds...");
        SpreadsheetApp.flush();
        processarMedicamentosRemoto(dados);

        // Passo 4: Compilação Local
        celulaStatus.setValue("📊 4/5 Compilando...");
        SpreadsheetApp.flush();
        compilarDados(dados);

        // Passo 5: Estoque
        celulaStatus.setValue("📈 5/5 Sinc. Estoque...");
        SpreadsheetApp.flush();
        sincronizarControleEstoque();

        celulaStatus.setValue("✅ TUDO PRONTO: " + new Date().toLocaleTimeString().slice(0,5));
        break;

      case 3: // Materiais
        celulaStatus.setValue("📦 Processando...");
        SpreadsheetApp.flush();
        processarMateriaisRemoto(); 
        celulaStatus.setValue("✅ Mat. OK: " + new Date().toLocaleTimeString().slice(0,5));
        break;

      case 4: // Medicamentos
        celulaStatus.setValue("💊 Processando...");
        SpreadsheetApp.flush();
        processarMedicamentosRemoto(); 
        celulaStatus.setValue("✅ Meds OK: " + new Date().toLocaleTimeString().slice(0,5));
        break;

      case 5: // Distribuir Equipe
        celulaStatus.setValue("👥 Distribuindo...");
        SpreadsheetApp.flush();
        distribuirDadosPorEquipe(); 
        celulaStatus.setValue("✅ Equipe OK: " + new Date().toLocaleTimeString().slice(0,5));
        break;

      case 6: // Sincronizar Estoque
        celulaStatus.setValue("📈 Atualizando...");
        SpreadsheetApp.flush();
        sincronizarControleEstoque(); 
        celulaStatus.setValue("✅ Estoque OK: " + new Date().toLocaleTimeString().slice(0,5));
        break;

      default:
        celulaStatus.setValue("⚠️ Botão sem função (Linha " + linha + ")");
    }
  } catch (erro) {
    celulaStatus.setValue("❌ Erro: " + erro.message);
    console.error(erro);
  }
}