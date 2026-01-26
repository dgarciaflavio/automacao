// =================================================================
// --- BLOCO 21: BUSCA REMOTA AVANÇADA (COM LINK DIRETO) ---
// =================================================================

function localizarItemNoPainelEquipe() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Solicita o Código
  const prompt = ui.prompt(
    "Localizar Planejador (Remoto)", 
    "Digite o Código do Item (ou parte dele):", 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  
  const termoBusca = String(prompt.getResponseText()).trim().toUpperCase();
  if (termoBusca === "") {
    ui.alert("Por favor, digite um código válido.");
    return;
  }

  const toast = (msg) => SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Buscando...", 5);
  toast("Varrendo planilha da equipe...");

  try {
    // 2. Conexão Remota
    const idPainel = CONFIG.ids.painelEquipe; // Pega do Config.js
    const ssRemota = SpreadsheetApp.openById(idPainel);
    const abas = ssRemota.getSheets();
    
    // 3. Filtros
    const ignorar = [
      "COAGE", "ESPELHOBASE", "CONFIGCMM", "CONFIG_EQUIPE", 
      "DASHBOARD", "RESUMO", "PÁGINA1", "PAGINA1"
    ];
    
    let encontrados = [];

    // 4. Varredura
    for (let i = 0; i < abas.length; i++) {
      const aba = abas[i];
      const nomeAba = aba.getName();
      
      if (ignorar.includes(nomeAba.toUpperCase())) continue;
      
      const lastRow = aba.getLastRow();
      if (lastRow < 2) continue;

      // Lê Código (B) e Descrição (C)
      const dados = aba.getRange(2, 2, lastRow - 1, 2).getValues(); 
      
      for (let j = 0; j < dados.length; j++) {
        const codItem = String(dados[j][0]).toUpperCase().trim();
        const descItem = dados[j][1];
        
        if (codItem.includes(termoBusca)) {
          // Captura o GID (ID da aba) para criar o link direto
          const sheetId = aba.getSheetId(); 
          const linhaReal = j + 2;
          
          encontrados.push({
            planejador: nomeAba,
            codigo: codItem,
            descricao: descItem,
            linha: linhaReal,
            link: `https://docs.google.com/spreadsheets/d/${idPainel}/edit#gid=${sheetId}&range=B${linhaReal}`
          });
        }
      }
    }

    // 5. Exibição (HTML Rico com Link)
    if (encontrados.length === 0) {
      ui.alert("Não Encontrado", `O item "${termoBusca}" não consta em nenhuma aba de planejador.`, ui.ButtonSet.OK);
    } else {
      
      // Monta o HTML do Pop-up
      let htmlContent = `
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 10px; }
          .card { background: #f9f9f9; border: 1px solid #ddd; padding: 15px; border-radius: 8px; margin-bottom: 10px; }
          .header { font-weight: bold; color: #1c4587; font-size: 16px; margin-bottom: 5px; }
          .info { color: #555; font-size: 14px; margin-bottom: 10px; }
          .btn { 
            display: inline-block; text-decoration: none; 
            background-color: #1155cc; color: white; 
            padding: 8px 15px; border-radius: 4px; font-size: 13px; font-weight: bold;
          }
          .btn:hover { background-color: #0c43a3; }
        </style>
        <h3>🔍 Resultado da Busca</h3>
        <p>Termo: "<strong>${termoBusca}</strong>"</p>
      `;

      encontrados.forEach(item => {
        htmlContent += `
          <div class="card">
            <div class="header">👤 ${item.planejador}</div>
            <div class="info">
              📦 <b>${item.codigo}</b><br>
              📝 ${String(item.descricao).substring(0, 40)}...<br>
              📍 Linha ${item.linha}
            </div>
            <a href="${item.link}" target="_blank" class="btn">
              🚀 Abrir Linha ${item.linha} no Painel
            </a>
          </div>
        `;
      });

      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(400)
        .setHeight(500);
      
      ui.showModalDialog(htmlOutput, 'Localizador de Itens');
    }

  } catch (e) {
    console.error(e);
    ui.alert("Erro na busca: " + e.message);
  }
}