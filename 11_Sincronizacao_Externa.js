// =================================================================
// --- BLOCO 11: SINCRONIZAÇÃO EXTERNA ---
// =================================================================

function buscarEmpenhosCodigosErrados() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const dadosGlobais = obterDadosEntradasGlobal();

    let abaCorrigir = ss.getSheetByName("Empenhos a Corrigir");
    
    if (!abaCorrigir) {
      abaCorrigir = ss.insertSheet("Empenhos a Corrigir");
    } else {
      abaCorrigir.clear();
    }

    const ssExterna = SpreadsheetApp.openById(CONFIG.ids.correcaoExterna);
    const abaExterna = ssExterna.getSheetByName("Empenhos");
    
    const lastRowExt = abaExterna.getLastRow();
    const mapExt = new Set(); 
    const listExt = [];
    
    if(lastRowExt > 1) {
       abaExterna.getRange(2,1,lastRowExt-1,6).getValues().forEach(r => {
          const e = _norm(r[3]); 
          const c = _norm(r[4]);
          if(e) { 
            mapExt.add(e); 
            if(c) listExt.push({e,c}); 
          }
       });
    }
    
    const mapEnt = new Map(); 
    const mapCodEmp = new Map();
    
    dadosGlobais.forEach(r => {
       const e = _norm(r[0]); 
       const c = _norm(r[2]);
       
       if(e && c && r[1] && _norm(r[13]) !== 'ELIMINADA') {
          if(!mapEnt.has(e)) mapEnt.set(e, new Set()); 
          mapEnt.get(e).add(c);
          if(!mapCodEmp.has(c)) mapCodEmp.set(c, new Set()); 
          mapCodEmp.get(c).add(e);
       }
    });
    
    const rel = []; 
    const chk = new Set();
    
    mapEnt.forEach((v,k) => {
       if(mapExt.has(k)) {
          v.forEach(c => {
             if(!listExt.some(x => x.e === k && x.c === c) && !chk.has(`${k}|${c}`)) {
                rel.push([k, c, "-", "FALTANTE", "Item no EMS mas falta na Planilha de Empenhos Enviados"]);
                chk.add(`${k}|${c}`);
             }
          });
       }
    });
    
    listExt.forEach(x => {
       const v = mapEnt.get(x.e);
       if(v && !v.has(x.c)) {
          const assoc = mapCodEmp.has(x.c) ? Array.from(mapCodEmp.get(x.c)).join(", ") : "Cod Inexistente";
          rel.push([x.e, x.c, assoc, "CÓDIGO ERRADO", "Cod dos enviados não bate com EMS"]);
       }
    });

    if(rel.length > 0) {
       abaCorrigir.getRange(1,1,1,5).setValues([["Empenho", "Código", "Empenho", "Erro", "Descrição"]]);
       abaCorrigir.getRange(2,1,rel.length,5).setValues(rel);
       abaCorrigir.getRange(2,1,rel.length,1).setNumberFormat("@");
       abaCorrigir.getRange(1,1,1,5).setFontWeight("bold").setBackground("#fce5cd");
       
       const resposta = ui.alert(
         "Divergências Encontradas", 
         `Foram encontrados ${rel.length} erros.\nDeseja enviar este relatório por e-mail para a equipe agora?`, 
         ui.ButtonSet.YES_NO
       );
       
       if (resposta == ui.Button.YES) {
         enviarEmailDivergencias(rel);
         ui.alert("E-mail enviado com sucesso!");
       }

    } else {
       ui.alert("Análise concluída: Tudo OK! Nenhuma divergência encontrada.");
    }

  } catch(e) { 
    ui.alert("Erro na análise: " + e.message); 
  }
}

function sincronizarEmpenhosNaExterna() {
   const ui = SpreadsheetApp.getUi();
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   try {
     const abaCorr = ss.getSheetByName("Empenhos a Corrigir");
     
     const dadosEnt = obterDadosEntradasGlobal();
     
     if(!abaCorr || abaCorr.getLastRow()<2) return ui.alert("Nada a corrigir");
     
     const list = abaCorr.getRange(2,1,abaCorr.getLastRow()-1,4).getValues();
     const mapEnt = new Map();
     
     dadosEnt.forEach(r => {
        if(r[0]&&r[2]) mapEnt.set(`${_norm(r[0])}|${_norm(r[2])}`, {d:r[1], ig:r[6], it:r[19]});
     });

     const ssExt = SpreadsheetApp.openById(CONFIG.ids.correcaoExterna);
     const abaExt = ssExt.getSheetByName("Empenhos");
     const novos = [];
     
     list.forEach(r => {
        if(r[3]==='FALTANTE') {
           const k = `${_norm(r[0])}|${_norm(r[1])}`;
           if(mapEnt.has(k)) {
              const d = mapEnt.get(k);
              const form = `=IF(A${abaExt.getLastRow()+novos.length+2}="";"";IFERROR(VLOOKUP(E${abaExt.getLastRow()+novos.length+2};ITENS!A:B;2;0);"Cadastrar"))`;
              novos.push([d.d, d.ig, d.it, r[0], r[1], form]);
           }
        }
     });

     if(novos.length) {
        abaExt.getRange(abaExt.getLastRow()+1, 1, novos.length, 6).setValues(novos);
        abaCorr.clearContents();
        ui.alert(`${novos.length} itens enviados.`);
     } else ui.alert("Nenhum item faltante válido.");
   } catch(e) { ui.alert(e.message); }
}

function repararDadosFaltantesNaExterna() {
   const ui = SpreadsheetApp.getUi();
   try {
      const dadosEnt = obterDadosEntradasGlobal();
      const map = new Map();
      
      dadosEnt.forEach(r => { if(r[0]) map.set(_norm(r[0]), {d:r[1], g:r[6], t:r[19]}); });

      const ssExt = SpreadsheetApp.openById(CONFIG.ids.correcaoExterna);
      const abaExt = ssExt.getSheetByName("Empenhos");
      const data = abaExt.getRange("A2:D"+abaExt.getLastRow()).getValues();
      
      let cnt = 0;
      data.forEach((r, i) => {
         const e = _norm(r[3]);
         if(e && map.has(e)) {
            const inf = map.get(e);
            let m=false;
            if(!r[0]){ r[0]=inf.d; m=true; }
            if(!r[1]){ r[1]=inf.g; m=true; }
            if(!r[2]){ r[2]=inf.t; m=true; }
            if(m) {
               abaExt.getRange(i+2, 1, 1, 3).setValues([[r[0], r[1], r[2]]]);
               cnt++;
            }
         }
      });
      ui.alert(`${cnt} linhas reparadas.`);
   } catch(e) { ui.alert(e.message); }
}