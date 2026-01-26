// =================================================================
// --- BLOCO 6: PROCESSAMENTO REMOTO - MEDICAMENTOS (LÓGICA HÍBRIDA) ---
// =================================================================
function processarMedicamentosRemoto(dadosGlobais) {
  try {
    if (!dadosGlobais) {
      dadosGlobais = obterDadosEntradasGlobal();
    }

    const ssMed = SpreadsheetApp.openById(CONFIG.ids.medicamentos);
    const wsEmp = ssMed.getSheetByName(CONFIG.abas.empenhos);
    const wsRec = ssMed.getSheetByName(CONFIG.abas.recProvisorio);
    const wsDest = ssMed.getSheetByName(CONFIG.abas.medicamentos);

    if (!wsEmp || !wsDest || !wsRec) {
      throw new Error("Abas obrigatórias não encontradas na planilha de Medicamentos.");
    }

    const recMap = new Map();
    const lrRec = wsRec.getLastRow();
    if(lrRec >= 2) {
       wsRec.getRange(2, 1, lrRec-1, 6).getValues().forEach(r => {
          const i = _norm(r[0]); 
          const e = String(r[5]).trim();
          if(i && e.includes('/')) {
             const p = e.split('/');
             const num = p[0].padStart(4,'0');
             const ano = (p[1].length===2) ? '20'+p[1] : p[1];
             recMap.set(`${ano}${num}-${i}`, Math.round(parseFloat(r[2])||0));
          }
       });
    }

    const notMap = new Map();
    const lrDest = wsDest.getLastRow();
    if(lrDest >= 2) {
       wsDest.getRange(2, 1, lrDest-1, 14).getValues().forEach(r => {
          if(r[0] && r[5]) notMap.set(`${r[0]}-${_norm(r[5])}`, r.slice(9,14));
       });
    }

    const entMap = new Map();
    const procModMap = new Map();

    dadosGlobais.forEach(r => {
       const k = `${String(r[0]).trim()}-${_norm(r[2])}`;
       if(r[0] && r[2]) {
          const qE = Math.round(parseFloat(r[9]) || 0);  
          const qS = Math.round(parseFloat(r[11]) || 0); 
          
          if(entMap.has(k)) { 
             let x = entMap.get(k); 
             x.qE += qE; 
             x.qS += qS; 
          } else {
             entMap.set(k, { 
                proc: r[25],
                forn: r[6], 
                desc: r[24], 
                unit: r[17], 
                qE: qE, 
                val: r[14], 
                qS: qS, 
                procT: r[19]
             });
          }
       }
       if(r[19]) procModMap.set(String(r[19]).trim(), r[21]);
    });

    const output = [];
    wsEmp.getRange(2, 1, wsEmp.getLastRow()-1, 6).getValues().forEach(r => {
       const emp = String(r[3]).trim(); 
       const cod = _norm(r[4]);
       
       if(!emp || !cod || !/^\d/.test(cod)) return;
       
       const k = `${emp}-${cod}`;
       const ent = entMap.get(k) || { qE: 0, qS: 0 };
       const anot = notMap.get(k) || Array(5).fill('');
       
       const dEnv = (r[0] instanceof Date) ? r[0] : null;
       const dVen = dEnv ? _addDays(dEnv, 10) : null;
       
       const qE = ent.qE;
       let qS_Oficial = ent.qS;
       
       let status = (qE === 0) ? "Solicitar Associação no EMS" : "";
       
       // --- LÓGICA HÍBRIDA ---
       let qS_Fisico = qS_Oficial;
       const isProvisorio = recMap.has(k);

       if (isProvisorio) {
          const qtdProv = recMap.get(k);
          if (qtdProv > qS_Fisico) qS_Fisico = qtdProv;
       }

       const saldoFisico = qE - qS_Fisico;
       
       status = _calcularStatusUnificado(qE, qS_Oficial, saldoFisico, isProvisorio, status);

       let atraso = "";
       if(status === 'Concluído') {
          atraso = 'Entregue';
       } else if (status.includes('Pendente') && dVen) {
          const diff = new Date().setHours(0,0,0,0) - dVen.getTime();
          atraso = diff > 0 ? _diasParaTexto(Math.floor(diff/86400000)) : "No prazo";
       }

       output.push([
          emp, 
          ent.proc || '', 
          dEnv, 
          dVen, 
          ent.forn || r[1] || "Não informado", 
          cod, 
          ent.desc || r[5] || "", 
          ent.unit || '', 
          qE,
          ...anot, 
          ent.val || '', 
          qS_Fisico, // Exibe o Físico
          saldoFisico, 
          atraso, 
          status, 
          ent.procT || '', 
          procModMap.get(ent.procT) || '' 
       ]);
    });

    wsDest.getRange(2, 1, wsDest.getMaxRows()-1, wsDest.getMaxColumns()).clearContent().clearFormat();
    
    if(output.length > 0) {
       wsDest.getRange(2, 1, output.length, output[0].length).setValues(output);
       wsDest.getRange(2, 3, output.length, 2).setNumberFormat("dd/mm/yyyy");
       wsDest.getRange(2, 21, output.length, 1).setNumberFormat("@");
       
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
       
       wsDest.getRange(2, 1, output.length, 21).setBackgrounds(coresGrid);
       wsDest.hideColumns(10); 
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro Medicamentos: " + e.message);
  }
}