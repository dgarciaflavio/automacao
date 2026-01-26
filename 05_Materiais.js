// =================================================================
// --- BLOCO 5: PROCESSAMENTO REMOTO - MATERIAIS (LÓGICA HÍBRIDA) ---
// =================================================================
function processarMateriaisRemoto(dadosGlobais) {
  try {
    // 1. Fallback: Se a função for chamada isoladamente
    if (!dadosGlobais) {
      dadosGlobais = obterDadosEntradasGlobal();
    }

    const ssMat = SpreadsheetApp.openById(CONFIG.ids.materiais);
    const wsEmpenhos = ssMat.getSheetByName(CONFIG.abas.empenhos);
    const wsRecProv = ssMat.getSheetByName(CONFIG.abas.recProvisorio);
    const wsDestino = ssMat.getSheetByName(CONFIG.abas.materiais);

    if (!wsEmpenhos || !wsDestino || !wsRecProv) {
      throw new Error("Abas 'Empenhos Enviados', 'Rec.Provisorio' ou 'Material' não encontradas em Materiais.");
    }

    // 2. Mapeamento de Recebimento Provisório
    const recProvisorioMap = new Map();
    const lastRowRec = wsRecProv.getLastRow();
    if (lastRowRec >= 2) {
      wsRecProv.getRange(2, 1, lastRowRec - 1, 6).getValues().forEach(row => {
        const itemCode = _norm(row[0]);
        const qtd = Math.round(parseFloat(row[2]) || 0); 
        const empStr = String(row[5]).trim();
        if (itemCode && empStr.includes('/')) {
           const parts = empStr.split('/');
           if (parts.length === 2) {
             const numero = parts[0].padStart(4, '0');
             const ano = (parts[1].length === 2) ? '20' + parts[1] : parts[1];
             recProvisorioMap.set(`${ano}${numero}-${itemCode}`, qtd);
           }
        }
      });
    }

    // 3. Mapeamento de Anotações Existentes
    const anotacoesMap = new Map();
    const lastRowDest = wsDestino.getLastRow();
    if (lastRowDest >= 2) {
       wsDestino.getRange(2, 1, lastRowDest - 1, 14).getValues().forEach(r => {
         const k = `${r[0]}-${_norm(r[5])}`;
         if (r[0]) anotacoesMap.set(k, r.slice(9, 14));
       });
    }

    // 4. Mapeamento dos Dados Globais (EM MEMÓRIA)
    const entradasMap = new Map();
    const modalidadeMap = new Map();

    dadosGlobais.forEach(r => {
       const emp = String(r[0]).trim();
       const item = _norm(r[2]);
       const key = `${emp}-${item}`;
       
       if(emp && item) {
         const qE = Math.round(parseFloat(r[9]) || 0); 
         const qS = Math.round(parseFloat(r[11]) || 0);
         
         if(entradasMap.has(key)) {
           let e = entradasMap.get(key);
           e.qE += qE; 
           e.qS += qS;
         } else {
           entradasMap.set(key, { 
             proc: r[19], 
             forn: r[6], 
             desc: r[24], 
             unit: r[17], 
             qE: qE, 
             val: r[14], 
             qS: qS, 
             local: r[25] 
           });
         }
       }
       if(r[19]) modalidadeMap.set(String(r[19]).trim(), r[21]);
    });

    // 5. Cruzamento de Dados
    const output = [];
    wsEmpenhos.getRange(2, 1, wsEmpenhos.getLastRow()-1, 6).getValues().forEach(r => {
      const emp = String(r[3]).trim();
      const cod = _norm(r[4]);
      
      if(!emp || !cod) return;
      
      const key = `${emp}-${cod}`;
      let status = 'Empenho não está na guia "Entradas"';
      
      let ent = entradasMap.get(key) || { qE: 0, qS: 0, local: '' };
      
      const localLimpo = String(ent.local).replace(/\s+/g, '').toUpperCase();
      if(!['ALM', 'MAI', '5Х5', ''].includes(localLimpo)) return;
      
      if(entradasMap.has(key)) status = ''; 

      const localFinal = (localLimpo === '5X5') ? 'ALM' : (ent.local || '');
      const forn = ent.forn || r[1] || "Não informado";
      const desc = ent.desc || r[5] || "";
      const unit = ent.unit || "";
      const val = ent.val || "";
      const proc = ent.proc || "";
      
      const qE = ent.qE; 
      let qS_Oficial = ent.qS; // Valor da Coluna L (Dados Globais)
      
      // --- LÓGICA HÍBRIDA DE SALDO ---
      // Usamos o maior valor entre Oficial e Provisório para calcular o Saldo Físico Real.
      // Isso garante que se tem 960 no físico (prov) e 14 no oficial, o saldo considere os 960.
      let qS_Fisico = qS_Oficial;
      const isProvisorio = recProvisorioMap.has(key);

      if (isProvisorio) {
          const qtdProv = recProvisorioMap.get(key);
          if (qtdProv > qS_Fisico) {
              qS_Fisico = qtdProv;
          }
      }
      
      const saldoFisico = qE - qS_Fisico;
      
      // Passamos:
      // 1. qS_Oficial -> Para o Helper decidir se sai do modo Provisório (se > 0)
      // 2. saldoFisico -> Para o Helper decidir se é Pendente, Resíduo 10% ou Concluído
      status = _calcularStatusUnificado(qE, qS_Oficial, saldoFisico, isProvisorio, status);

      let obsAtraso = "";
      const dVenc = (r[0] instanceof Date ? _addDays(r[0], 10) : null);
      if(status === 'Concluído') {
        obsAtraso = 'Entregue';
      } else if (status.includes('Pendente') && dVenc) { 
         const diff = new Date().setHours(0,0,0,0) - dVenc.getTime();
         obsAtraso = diff > 0 ? _diasParaTexto(Math.floor(diff/86400000)) : "No prazo";
      }

      const linha = [
        emp, 
        localFinal, 
        (r[0] instanceof Date ? r[0] : null),
        dVenc,
        forn, 
        cod, 
        desc, 
        unit, 
        qE,
        ...(anotacoesMap.get(key) || Array(5).fill('')),
        val, 
        qS_Fisico, // Visualmente mostramos o que está lá fisicamente (960)
        saldoFisico, 
        obsAtraso, 
        status, 
        proc, 
        modalidadeMap.get(proc) || null
      ];
      output.push(linha);
    });

    // 6. Escrita na Aba de Destino
    wsDestino.getRange(2, 1, wsDestino.getMaxRows()-1, wsDestino.getMaxColumns()).clearContent().clearFormat();
    
    if(output.length > 0) {
       wsDestino.getRange(2, 1, output.length, output[0].length).setValues(output);
       wsDestino.getRange(2, 3, output.length, 2).setNumberFormat("dd/mm/yyyy");
       wsDestino.getRange(2, 16, output.length, 1).setNumberFormat('0.00'); 
       
       const coresGrid = output.map(r => {
          const statusCell = String(r[18]).trim().toUpperCase();
          const dataVenc = r[3];
          let cor = CONFIG.cores[statusCell] || null;
          
          if (!cor && statusCell === 'PENDENTE' && dataVenc instanceof Date) {
             const diff = new Date().setHours(0,0,0,0) - dataVenc.getTime();
             if (diff > 0) cor = '#FFCDD2';
             else cor = '#FFFFE0';
          } else if (!cor && statusCell === 'PENDENTE') {
             cor = CONFIG.cores['PENDENTE'];
          }
          return new Array(21).fill(cor);
       });
       
       wsDestino.getRange(2, 1, output.length, 21).setBackgrounds(coresGrid);
       wsDestino.hideColumns(10); 
       
       const range = wsDestino.getDataRange();
       if(range.getFilter()) range.getFilter().remove();
       const filtro = range.createFilter();
       
       const locaisVisiveis = new Set(['ALM']);
       output.forEach(r => { if(r[18] === 'Recebimento Provisório') locaisVisiveis.add(r[1]); });
       
       const ocultar = [...new Set(output.map(r => r[1]))].filter(x => !locaisVisiveis.has(x));
       if(ocultar.length) {
         filtro.setColumnFilterCriteria(2, SpreadsheetApp.newFilterCriteria().setHiddenValues(ocultar).build());
       }
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro Materiais: " + e.message);
  }
}