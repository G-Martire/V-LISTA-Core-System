function importaNoteMultipleModulare() {
  // prima di importare nuove note, eseguiamo la pulizia. questo impedisce errori in "V_lista"
  mainManutenzione();



  const dataIn = shCF_Lista_in.getRange("C1").getValue();
  const rangeTabella = ssCF.getRangeByName("Tb_lista_in_D");
  const intestazioni = rangeTabella.getValues()[0];
  
  // --- NUOVA LOGICA: PESCHIAMO DA GOOGLE DOCS ---
  
  const corpoDoc = docword.getBody();
  const testoNotaGrezzo = corpoDoc.getText();

  // Se il Doc è vuoto, fermiamo tutto per non inserire righe bianche
  if (testoNotaGrezzo.trim() === "") {
    Logger.log("Documento vuoto. Nessuna operazione eseguita.");
    return; 
  }

  // --- CHIAMATA AL TUO MODULO DI ANALISI ---
  let listaTask = analizzatoreTesto(testoNotaGrezzo, dataIn);

  // LOGICA LIFO: Invertiamo l'ordine così l'ultima nota inserita finisce in alto
  listaTask.reverse();

  if (listaTask.length > 0) {
    const rigaIntestazione = rangeTabella.getRow();
    const colonnaInizio = rangeTabella.getColumn();
    
    // --- MODIFICA SKU DINAMICO ---
    // Prepariamo la matrice inserendo la formula testuale ="SKU-" & RIF.RIGA()
    const matriceDati = listaTask.map(task => 
      intestazioni.map(titolo => {
        let t = titolo.toUpperCase().trim();
        if (t === "SKU") {
          return '="SKU-" &ROW()-2'; // Inserisce la funzione invece del valore fisso
        }
        return task[t] || "";
      })
    );

    // Inserimento righe e scrittura dati (Sheets attiverà la formula automaticamente)
    shCF_Lista_in.insertRowsAfter(rigaIntestazione, matriceDati.length);
    shCF_Lista_in.getRange(rigaIntestazione + 1, colonnaInizio, matriceDati.length, intestazioni.length)
          .setValues(matriceDati);

    Logger.log("Importazione completata: " + listaTask.length + " note inserite con SKU dinamici.");

    // --- PULIZIA DEL DOC ---
    corpoDoc.clear();
    
    // formattazione data
    const dataFormattata = Utilities.formatDate(new Date(dataIn), "GMT+1", "dd/MM/yy").toLowerCase();

    // Template per la prossima nota
    corpoDoc.appendParagraph("@ PC LA EA WC  " + dataFormattata + " --   //");
    corpoDoc.appendParagraph("\n" + "@  --   //");
    corpoDoc.appendParagraph("\n" + "@  --   //");
  }


  // --- LOGICA DI AZZERAMENTO CONTATORE "FATTE" ---
    const oraAttuale = new Date().getHours();
    
    // Se l'esecuzione avviene tra le 00:00 e le 04:00 (o l'orario che preferisci)
    // Usiamo >= 0 e < 5 per coprire la fascia notturna di reset giornaliero
    if (oraAttuale >= 0 && oraAttuale < 5) {
      
      // Accediamo al foglio V_LISTA tramite lo Spreadsheet principale CF
      const rangeImpost = ssCF.getRangeByName("Impost_V_lista");
      
      if (rangeImpost) {
        // La cella FATTE è la 4a colonna dell'intervallo D1:G1 (cella G1)
        let cellaFatte = rangeImpost.getCell(1, 4); 
        cellaFatte.setValue(0);
        Logger.log("Reset notturno eseguito: Contatore FATTE impostato a 0.");
      }
    } else {
      Logger.log("Esecuzione diurna: il contatore FATTE non è stato azzerato (Ora attuale: " + oraAttuale + ").");
    }
   // Chiusura dell'if (listaTask.length > 0)
}


function analizzatoreTesto(testoGrezzo, dataIn) {
  const cassaforte = [];
  const LIMITE_CARATTERI = 450; // Soglia per lo spostamento su Doc esterno
  
  let testoProtetto = testoGrezzo.replace(/§§([\s\S]*?)§§/g, (match, p1) => {
    const placeholder = `___ID${cassaforte.length}___`;
    cassaforte.push(p1);
    return placeholder;
  });

  // Estraiamo i blocchi grezzi
  const blocchiGrezzi = testoProtetto.match(/@[\s\S]*?\/\/+/g) || [];
  
  // --- FILTRO RIGHE VUOTE ---
  const blocchiValidi = blocchiGrezzi.filter(blocco => {
    let pulito = blocco
      .replace(/@| \/\/+/g, "") 
      .replace(/[PLEW][ABC][\+\-]*/gi, "") 
      .replace(/--/g, "") 
      .trim();
    return pulito !== ""; 
  });

  return blocchiValidi.map((blocco, indice) => {
    let contenuto = blocco.replace(/^@/, "").replace(/\/\/+$/, "").trim();

    let indiceSplit = contenuto.indexOf("--");
    let metaString = (indiceSplit !== -1) ? contenuto.substring(0, indiceSplit).trim() : "";
    let corpoDesc = (indiceSplit !== -1) ? contenuto.substring(indiceSplit + 2).trim() : contenuto;

    // --- 1. ESTRAZIONE DATA ---
    let regexData = /(\d{1,2}\/\d{1,2}(?:\/\d{2,4})?)/;
    let matchData = metaString.match(regexData);
    let dataScadenza = "";
    
    if (matchData) {
      dataScadenza = matchData[1];
      metaString = metaString.replace(dataScadenza, "").trim();
    }

    // --- 2. VALIDAZIONE OCEAN ---
    let validazione = verificaClassiEvoluta(metaString.toUpperCase().replace(/\s+/g, ''));

    // --- 3. RIPRISTINO LINK ---
    cassaforte.forEach((originale, i) => {
      const cerca = `___ID${i}___`;
      corpoDesc = corpoDesc.split(cerca).join(originale);
    });

    // --- 4. LOGICA ARCHIVIAZIONE NOTE LUNGHE ---
    if (corpoDesc.length > LIMITE_CARATTERI) {
      try {
        const docArchivio = DocumentApp.openById(ID_DOC_ARCHIVIO);
        const bodyArchivio = docArchivio.getBody();
        const dataLog = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy HH:mm");
        
        // Appendiamo al file esterno
        bodyArchivio.appendParagraph("--- NOTA ARCHIVIATA IL " + dataLog + " ---");
        bodyArchivio.appendParagraph("DATA SCAD: " + dataScadenza);
        bodyArchivio.appendParagraph(corpoDesc);
        bodyArchivio.appendHorizontalRule();
        
        // Tagliamo la descrizione per lo Sheet
        corpoDesc = corpoDesc.substring(0, LIMITE_CARATTERI) + "... [TESTO LUNGO: VEDI DOC ALTRE ATTIVITÀ]";
      } catch (e) {
        Logger.log("Errore archiviazione Doc: " + e.message);
      }
    }

    return {
      "DESCRIZIONE": corpoDesc,
      "PRIORITA'": validazione.p,
      "LUNG. T": validazione.l,
      "EXECUTION": validazione.e,
      "WATT": validazione.w,
      "DATA_IN": dataIn,
      "DATA_SCAD": dataScadenza, 
      "STATO": validazione.residuo !== "" ? "da verificare" : "da fare",
      "ELEMENTO. VER": validazione.residuo
    };
  });
}



function verificaClassiEvoluta(stringaMeta) {
  let erroreFormato = false;
  
  const estraiParametro = (lettera, dflt) => {
    // Regex: cerca Lettera + (A o B o C) + eventuali segni + o -
    let regex = new RegExp(lettera + "([ABC][\\+\\-]*|[ABC])", "i");
    let match = stringaMeta.match(regex);
    
    if (match) return match[1].toUpperCase();
    
    // Se trovi la lettera ma non il formato ABC (es. PX o BM), segnali errore
    if (new RegExp(lettera, "i").test(stringaMeta)) erroreFormato = true;
    return dflt;
  };

  let parametri = {
    p: estraiParametro("P", "C"),
    l: estraiParametro("L", "A"),
    e: estraiParametro("E", "A"),
    w: estraiParametro("W", "C")
  };

  // Pulizia residuo per identificare "BM" o errori (Colonna J)
  let residuo = stringaMeta;
  Object.values(parametri).forEach(val => {
    // Protezione per i caratteri speciali della regex (+ e -)
    let escapedVal = val.replace(/[\+\-]/g, "\\$&");
    residuo = residuo.replace(new RegExp("[PLEW]?" + escapedVal, "i"), "");
  });
  
  residuo = residuo.replace(/-/g, "").trim();

  return {
    ...parametri,
    residuo: (residuo !== "" || erroreFormato) ? stringaMeta : ""
  };
}