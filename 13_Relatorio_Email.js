// =================================================================
// --- BLOCO 13: RELATÓRIO GERENCIAL POR E-MAIL ---
// =================================================================

function enviarRelatorioGerencial() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const resp = ui.alert(
      "Enviar Relatório",
      `Deseja enviar o resumo atual do status para a lista de e-mails configurada?\n\nPara: ${CONFIG.emails.para}\nCópia: ${CONFIG.emails.copia}`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;

    const abaComp = ss.getSheetByName(CONFIG.destino.nomeAba); 
    if (!abaComp) throw new Error("Aba Compilados não encontrada.");

    const lastRow = abaComp.getLastRow();
    if (lastRow < 2) throw new Error("Sem dados para reportar.");

    const dados = abaComp.getRange(2, 1, lastRow - 1, 19).getValues();
    
    const stats = {};
    let total = 0;
    
    dados.forEach(r => {
      const st = r[18] ? String(r[18]).trim().toUpperCase() : "SEM STATUS";
      stats[st] = (stats[st] || 0) + 1;
      total++;
    });

    const listaStats = Object.keys(stats)
      .map(k => ({ status: k, qtd: stats[k], pct: stats[k]/total }))
      .sort((a, b) => b.qtd - a.qtd);

    let html = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px;">
      <h2 style="background-color: #4285f4; color: white; padding: 10px; border-radius: 5px;">
        🚀 Status Geral de Empenhos
      </h2>
      <p>Olá equipe,</p>
      <p>Segue o panorama atualizado do processamento de <strong>Materiais e Medicamentos</strong>.</p>
      
      <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
        <tr style="background-color: #f2f2f2; text-align: left;">
          <th style="padding: 8px; border-bottom: 2px solid #ddd;">Status</th>
          <th style="padding: 8px; border-bottom: 2px solid #ddd;">Qtd</th>
          <th style="padding: 8px; border-bottom: 2px solid #ddd;">%</th>
        </tr>
    `;

    listaStats.forEach(item => {
      let corTexto = "#333";
      let bg = "#fff";
      if(item.status === 'CONCLUÍDO') { bg = '#d9ead3'; corTexto = '#274e13'; }
      if(item.status.includes('PENDENTE')) { bg = '#f4cccc'; corTexto = '#990000'; }
      if(item.status.includes('RESÍDUO')) { bg = '#cfe2f3'; }

      html += `
        <tr style="background-color: ${bg}; color: ${corTexto};">
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${item.status}</td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;"><strong>${item.qtd}</strong></td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">${(item.pct * 100).toFixed(1)}%</td>
        </tr>
      `;
    });

    html += `
      </table>
      
      <div style="margin-top: 20px; padding: 10px; background-color: #eee; border-radius: 5px;">
        <strong>Total Processado:</strong> ${total} itens<br>
        <strong>Atualização:</strong> ${new Date().toLocaleString('pt-BR')}
      </div>
      
      <p style="font-size: 12px; color: #888; margin-top: 30px;">
        Enviado automaticamente pelo Orquestrador (Operador: ${Session.getActiveUser().getEmail()})
      </p>
    </div>
    `;

    MailApp.sendEmail({
      to: CONFIG.emails.para,
      cc: CONFIG.emails.copia,
      subject: `📊 Relatório de Status: ${new Date().toLocaleDateString('pt-BR')}`,
      htmlBody: html
    });

    ss.toast("E-mail enviado com sucesso!", "Relatório Gerencial");

  } catch (e) {
    ui.alert("Erro ao enviar relatório: " + e.message);
  }
}