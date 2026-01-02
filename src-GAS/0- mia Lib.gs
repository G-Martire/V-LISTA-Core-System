/**
 * Copia i formati tra Range con controllo di compatibilità dimensionale.
 * * @param {GoogleAppsScript.Spreadsheet.Range} rangeSorgente L'intervallo di origine.
 * @param {GoogleAppsScript.Spreadsheet.Range} rangeDestinazione L'intervallo di destinazione.
 * @param {('SFONDO'|'FORMATO_TESTO'|'TUTTO')} opzioneFormato Opzioni: SFONDO, FORMATO_TESTO, TUTTO.
 */
function copiaFormatoAdattato(rangeSorgente, rangeDestinazione, opzioneFormato) {
  const rigaS = rangeSorgente.getNumRows();
  const colS = rangeSorgente.getNumColumns();
  const rigaD = rangeDestinazione.getNumRows();
  const colD = rangeDestinazione.getNumColumns();

  // 1. VERIFICA COMPATIBILITÀ
  const isCellaSingola = (rigaS === 1 && colS === 1);
  const stessaDimensione = (rigaS === rigaD && colS === colD);
  const compatibilePerColonne = (colS === colD); // per casi come A1:C1 -> A2:C5
  const compatibilePerRighe = (rigaS === rigaD);   // per casi come A1:A3 -> C3:D3

  if (!isCellaSingola && !stessaDimensione && !compatibilePerColonne && !compatibilePerRighe) {
    const msgErrore = `ERRORE COMPATIBILITÀ: Sorgente [${rigaS}x${colS}] non può essere copiata in Destinazione [${rigaD}x${colD}].`;
    Logger.log(msgErrore);
    throw new Error(msgErrore);
  }

  // 2. ESECUZIONE COPIA
  if (opzioneFormato === "SFONDO") {
    const backgrounds = rangeSorgente.getBackgrounds();
    let matriceColori;

    if (isCellaSingola) {
      // Caso 1: 1 cella -> Espansione totale
      const colore = backgrounds[0][0];
      matriceColori = Array.from({ length: rigaD }, () => Array(colD).fill(colore));
    } 
    else if (rigaS === 1 && colS === colD) {
      // Caso 2: Una riga sorgente -> ripetuta su più righe destinazione (es: A1:C1 -> A2:C5)
      matriceColori = Array.from({ length: rigaD }, () => backgrounds[0]);
    }
    else if (colS === 1 && rigaS === rigaD) {
      // Caso 3: Una colonna sorgente -> ripetuta su più colonne destinazione (es: A1:A3 -> C3:E3)
      matriceColori = backgrounds.map(riga => Array(colD).fill(riga[0]));
    }
    else {
      // Caso 4: Dimensioni identiche
      matriceColori = backgrounds;
    }

    rangeDestinazione.setBackgrounds(matriceColori);
    Logger.log("Sfondo applicato con successo.");

  } else {
    // Per FORMATO_TESTO o TUTTO usiamo copyTo
    // copyTo gestisce nativamente la ripetizione se una dimensione è uguale (es: 1 riga su 10 righe)
    let copyOption = (opzioneFormato === "FORMATO_TESTO") ? 
        SpreadsheetApp.CopyPasteType.PASTE_FORMATTING : 
        SpreadsheetApp.CopyPasteType.PASTE_FORMATTING; 

    rangeSorgente.copyTo(rangeDestinazione, copyOption, false);
    Logger.log(`Formato (${opzioneFormato}) copiato con copyTo.`);
  }
}



/**
 * Per prestazione, trasforma l'array di lookup in una Map. Simula la funzione CERCA.VERT 
 * (VLOOKUP) su una matrice e supporta l'elaborazione Batch.
 * Esegue un controllo rigoroso sui dati sorgente e lancia un errore se la colonna chiave contiene duplicati.
 * Cerca uno o più valori in una colonna e restituisce il valore/array di valori corrispondente/i
 * dalla colonna di ritorno.
 *
 * @param {boolean} [Use_Map=true] Se vuoto (true) o impostato su vero, usa la Map (O(1)). Se falso, usa la ricerca lineare (O(N)).
 * @param {Array<Array<any>>} array_bidimensionale La matrice di dati in cui effettuare la ricerca (la tabella di lookup).
 * @param {any|Array<any>} valore_da_cercare Il valore specifico (o array di valori anche non omogenei) da trovare nella colonna cerca_IN.
 * @param {number} cerca_IN L'indice della colonna (base 0) in cui cercare valore_da_cercare (la chiave).
 * @param {number} cerca_OUT L'indice della colonna (base 0) da cui restituire il valore.
 * @returns {any|null|Array<any|null>} Il valore/i trovato/i, un array di risultati se l'input era array, o null se la ricerca fallisce.
*/
function Map_CercaVert(Use_Map = true, array_bidimensionale, valore_da_cercare, cerca_IN, cerca_OUT) {
  
  // eliminazione righe vuote in base alla chiave (cerca_IN)
  array_bidimensionale = array_bidimensionale.filter(row => {
    // 1. Assicurati che la riga abbia abbastanza colonne
    // 2. Assicurati che il valore nella colonna chiave (cerca_IN) NON sia vuoto o nullo.
    return (row.length > cerca_IN && String(row[cerca_IN]).trim() !== "");
  });
  
  // 1- Cicli di verifica
  // 1.1- Assicurati che gli indici siano validi e non negativi
  if (!Array.isArray(array_bidimensionale) || cerca_IN < 0 || cerca_OUT < 0) {
    Logger.log("Errore: Parametri di array o indici non validi.");
    return null;
  }

  if (Use_Map === true || Use_Map === "") {
    // 1.2- Ciclo verifica duplicati
    const keyCounts = {}; // Traccia il conteggio di ogni chiave
    const duplicatiTrovati = []; // Lista per i valori duplicati

    // verifica se la lista dei duplicati è vuota
    for (let i = 1; i < array_bidimensionale.length; i++) {
      const key = String(array_bidimensionale[i][cerca_IN]).trim().toUpperCase().replace(',', '');

      if (key === "") continue; // Salta celle vuote

      // 1. Verificare: Incrementa il contatore per la chiave
      keyCounts[key] = (keyCounts[key] || 0) + 1;

      // 2. Registrare i duplicati: Se il conteggio è 2, è la prima volta che lo identifichiamo come duplicato.
      if (keyCounts[key] === 2) {
        duplicatiTrovati.push(key);
      }
    }

    // Feedback di errore duplicati
    if (duplicatiTrovati.length > 0) {
      let messaggioErrore = `ERRORE CRITICO: Impossibile creare Map. Trovati ${duplicatiTrovati.length} valori chiave duplicati. \n\n`;

      // Aggiungi dettagli sui duplicati (quantità e nome)
      duplicatiTrovati.forEach(key => {
        messaggioErrore += ` - Chiave: "${key}" | Trovata ${keyCounts[key]} volte.\n`;
      });
      
      // Blocca l'esecuzione
      throw new Error(messaggioErrore);
    }

    // _____________________________________________________________________________________________________________________
    // 2- Ciclo per creare la mappa
    const lookupMap = {};

    // Ciclo per popolare la Map (inizia da 1 se la prima riga è l'intestazione)
    for (let i = 1; i < array_bidimensionale.length; i++) {
      const row = array_bidimensionale[i];
      
      // Preparazione della Chiave (come nel ciclo di verifica)
      const key = String(row[cerca_IN]).trim().toUpperCase().replace(',', '').replace(' ', '');
      const value = row[cerca_OUT]; // Il valore da restituire

      if (key === "") continue; // Salta celle vuote

      // La chiave è univoca (grazie al controllo precedente), quindi l'assegnazione è sicura.
      lookupMap[key] = value;
    }

    // _____________________________________________________________________________________________________________________
    // 3. ESECUZIONE DELLA RICERCA E RITORNO

    // 1. Determina se l'input è un array (ricerca batch)
    const isBatchSearch = Array.isArray(valore_da_cercare); 

    // 2. Normalizza l'input: trasforma il singolo valore in un array con un solo elemento
    const lookupArray = isBatchSearch ? valore_da_cercare : [valore_da_cercare];

    const resultsBatch = [];

    // 3. Cicla l'array di ricerca
    for (const singleValue of lookupArray) {
      // Prepara la chiave di ricerca (deve avere lo stesso formato di pulizia della chiave Map)
      const searchValue = String(singleValue).trim().toUpperCase().replace(',', '').replace(' ', '');
      
      // Ricerca istantanea nella Map (O(1))
      const result = lookupMap[searchValue];
      
      // Aggiungi il risultato (valore o null, se non trovato) all'array di output
      resultsBatch.push((result !== undefined) ? result : null);
    }

    // _____________________________________________________________________________________________________________________
    // 4. Restituisce il risultato nel formato appropriato
    if (isBatchSearch) {
      // Se l'input era un array, restituisce l'array completo dei risultati
      return resultsBatch; 
    } else {
      // Se l'input era un singolo valore, restituisce il singolo elemento trovato
      return resultsBatch[0]; 
    }

    // _____________________________________________________________________________________________________________________
    // Ricerca senza mappa
  } else if (Use_Map === false) {

    const isBatchSearch = Array.isArray(valore_da_cercare); 
    const lookupArray = isBatchSearch ? valore_da_cercare : [valore_da_cercare];
    const resultsBatch = [];

    // 1. Cicla l'array di ricerca
    for (const singleValue of lookupArray) {
        
      let foundResult = null;
      
      // 2. Esegui la ricerca lineare per il singolo valore
      for (let i = 0; i < array_bidimensionale.length; i++) {
        const row = array_bidimensionale[i];
        
        // 💡 NOTA IMPORTANTE: Qui non facciamo la pulizia della stringa (trim, upperCase, replace)
        // L'approccio lineare (Use_Map = false) solitamente cerca una corrispondenza esatta
        // Se volessi la ricerca "pulita", dovresti pulire sia row[cerca_IN] che singleValue qui.
        
        // Assicurati che la riga abbia abbastanza colonne per l'indice di ricerca
        if (row.length > cerca_IN && row[cerca_IN] === singleValue) {
          // Assicurati che la riga abbia abbastanza colonne per l'indice di ritorno
          if (row.length > cerca_OUT) {
            foundResult = row[cerca_OUT];
            break; // Trovato il primo risultato, esci dal ciclo interno
          }
        }
      }
      
      resultsBatch.push(foundResult);
    }
    
    // 3. Restituisce il risultato nel formato appropriato
    if (isBatchSearch) {
      return resultsBatch; 
    } else {
      return resultsBatch[0]; 
    }
  }
}




//nuova funzione: ripostare nella libreria!
/**
 * Prende una stringa o un valore, lo pulisce e lo spezza in un array di stringhe.
 * Gestisce virgole e spazi multipli. da come risultato un array.
 * carattere di base: ","
 * * @param {any} input Il valore da processare (es: "lun 29/12, ven 02/01").
 * @param {string} separatore Il carattere su cui dividere (default ",").
 * @return {string[]} Array di stringhe pulite.
 */
function NormalizzaDateInput(input, separatore = ",") {
  if (!input) return [];
  let testo = input.toString().toLowerCase().trim();
  
  // Se contiene il separatore, splitta, altrimenti crea array singolo
  let parti = testo.includes(separatore) ? testo.split(separatore) : [testo];
  
  // Pulisce ogni elemento da spazi extra interni e restituisce l'array
  return parti.map(p => p.replace(/\s+/g, ' ').trim()).filter(p => p !== "");
}
