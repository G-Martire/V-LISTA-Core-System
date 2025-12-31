function mainManutenzione() {
  Logger.log("--- INIZIO STEP 1: LETTURA E ASSOCIAZIONE ---");
  
  const datiInput = ssCF.getRangeByName("Check_V_lista").getValues(); 
  const datiDB = ssCF.getRangeByName("Tb_lista_in_D").getValues();
  const rangeCheck = ssCF.getRangeByName("Check_V_lista");
  
  let lista_per_LISTA_IN = [];
  let lista_per_ESTERNI = [];

  for (let i = 1; i < datiInput.length; i++) {
    let check = datiInput[i][0];    
    let dataRim = datiInput[i][1]; 
    let sku = datiInput[i][2];      
    if (!sku) continue; 

    let rigaCompleta = datiDB.find(r => r[0] === sku);
    if (!rigaCompleta) continue;

    let pacchettoCompleto = {
      azione: check,
      nuovaData: dataRim,
      datiRiga: rigaCompleta 
    };

    if (check === "da master" || check === "da listone") {
      lista_per_ESTERNI.push(pacchettoCompleto);
    } 
    else if (check !== "" || dataRim !== "") {
      lista_per_LISTA_IN.push(pacchettoCompleto);
    }
  }

  Logger.log("Associazione completata. Gruppo 1: " + lista_per_LISTA_IN.length + " | Gruppo 2: " + lista_per_ESTERNI.length);

  // ESECUZIONE
  let skuDaCancellareG1 = scriviDatiInterni(lista_per_LISTA_IN);
  let skuDaCancellareG2 = scriviDatiEsterni(lista_per_ESTERNI);
  
  // UNIONE LISTE CANCELLAZIONE
  let skuTotaliDaRimuovere = [...skuDaCancellareG1, ...skuDaCancellareG2];
  
  // --- NUOVA LOGICA: AGGIORNAMENTO CONTATORE FATTE ---
  if (skuTotaliDaRimuovere.length > 0) {
    const rangeImpost = ssCF.getRangeByName("Impost_V_lista");
    if (rangeImpost) {
      // La cella FATTE è la quarta colonna dell'intervallo (G1)
      let cellaFatte = rangeImpost.getCell(1, 4); 
      let valoreAttuale = cellaFatte.getValue() || 0;
      let nuovoValore = valoreAttuale + skuTotaliDaRimuovere.length;
      
      cellaFatte.setValue(nuovoValore);
      Logger.log("Contatore FATTE aggiornato: " + valoreAttuale + " -> " + nuovoValore);
    }

    // PULIZIA FINALE RIGHE
    eseguiPuliziaRighe(skuTotaliDaRimuovere);
  }

  // reset check 
  if (rangeCheck) {
    rangeCheck.offset(1, 0, rangeCheck.getNumRows() - 1, 2).clearContent();
    Logger.log("Reset dei check (colonne A e B) completato.");
  }
}



/** FUNZIONE 2 AGGIORNATA **/
function scriviDatiInterni(listaPacchetti) {
  Logger.log("--- INIZIO SCRITTURA GRUPPO 1 ---");
  const foglioIn = ssCF.getSheetByName("LISTA_IN");
  const rangeDB = ssCF.getRangeByName("Tb_lista_in_D");
  const rigaInizio = rangeDB.getRow();
  const datiDB = rangeDB.getValues();
  let skuDaCancellare = [];

  listaPacchetti.forEach(p => {
    let skuCorrente = p.datiRiga[0]; // FIX: Prende lo SKU dal pacchetto
    let indiceRiga = datiDB.findIndex(r => r[0] === skuCorrente);
    let rigaFoglio = indiceRiga + rigaInizio;

    if (p.azione === "fatto") {
      skuDaCancellare.push(skuCorrente);
      Logger.log("SKU " + skuCorrente + " segnato come FATTO");
    } else {
      if (p.azione !== "") foglioIn.getRange(rigaFoglio, 10).setValue(p.azione);
      if (p.nuovaData !== "") {
        foglioIn.getRange(rigaFoglio, 8).setValue(p.nuovaData);
        foglioIn.getRange(rigaFoglio, 10).setValue("rimandato");
      }
      Logger.log("SKU " + skuCorrente + " aggiornato correttamente.");
    }
  });
  return skuDaCancellare;
}

/** FUNZIONE 3 AGGIORNATA **/
function scriviDatiEsterni(listaPacchetti) {
  Logger.log("--- INIZIO SCRITTURA GRUPPO 2 ---");
  let skuDaCancellare = [];

  listaPacchetti.forEach(p => {
    let skuCorrente = p.datiRiga[0]; // FIX: Prende lo SKU dal pacchetto
    let nomeFoglio = (p.azione === "da master") ? "MasterPlan_DB" : "LISTONE";
    let foglioDest = ssCF.getSheetByName(nomeFoglio);

    let datiDaIncollare = ['="SKU-" & ROW()-2', ...p.datiRiga.slice(1)];
    foglioDest.insertRowBefore(3);
    foglioDest.getRange(3, 1, 1, datiDaIncollare.length).setValues([datiDaIncollare]);

    Logger.log("SKU " + skuCorrente + " spostato in " + nomeFoglio);
    skuDaCancellare.push(skuCorrente);
  });
  return skuDaCancellare;
}




function eseguiPuliziaRighe(skuTotaliDaRimuovere) {
  const foglioIn = ssCF.getSheetByName("LISTA_IN");
  const rangeCompleto = ssCF.getRangeByName("Tb_lista_in_D");
  const valoriAttuali = rangeCompleto.getValues();
  const rigaInizioRange = rangeCompleto.getRow(); // Riga 3

  Logger.log("Inizio pulizia fisica. SKU da rimuovere: " + skuTotaliDaRimuovere.length);

  // Ciclo inverso per non sballare gli indici
  for (let i = valoriAttuali.length - 1; i >= 0; i--) {
    let skuInRiga = valoriAttuali[i][0];
    if (skuTotaliDaRimuovere.includes(skuInRiga)) {
      let rigaDaEliminare = i + rigaInizioRange;
      foglioIn.deleteRow(rigaDaEliminare);
      Logger.log("Eliminata riga " + rigaDaEliminare + " (" + skuInRiga + ")");
    }
  }
}