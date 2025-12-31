// =====================================================================================================================================
// =====================================================================================================================================
// SCRIPT - questo script serve ad azzerare il valore di controllo (nella colonna K)
function reset() {
  // COORDINATE CELLE FISSE
  let coordinataGe = shCF_Gestione.getRange("J13");
  let coordinataAgX = "K";

  // COORDINATE CELLE letIABILI
  let n = coordinataGe.getValue() - 1
  let coordinataAgY = 5 + n * 12;
  let coordinataAg = shCF_Agenda.getRange(coordinataAgX + coordinataAgY)
  
  // settare volore
  coordinataAg.setValue(0)
  Logger.log("coordinataAgY: " + coordinataAgY)



  // chiamiamo le funzioni per resettare l'unione e il bordo
  reset_colore_1(coordinataAgY);
  reset_unione_2(coordinataAgY);   // toglie l'unione se NON c'è la "*"
  reset_bordo_3(coordinataAgY);
  
  // funzione "ITS" per i lavori di gruppo

  // Restituiamo il valore
  return (coordinataAgY);
}

// questo script, automatizza lo script precedente. viene eseguit ogni domenica tra le 22.00 e le 23.00
function trigger_reset() {
  ScriptApp.newTrigger('reset')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(22)
    .nearMinute(30)
    .create();
}


// reset_colore_1,  reset_unione_2,  reset_bordo_3, SI ATTIVANO insime alla funzione "reset"
// funzione reset base
function reset_colore_1 (rigaBase) {

  const x = rigaBase - 2
  const celle = "A" + (x) + ":I" + (x + 11);
  const range = shCF_Agenda.getRange(celle)

  range.setBackground(null);
  range.setFontColor(null);
}

// funzione reset (unione)
function reset_unione_2(rigaBase) {
   
  const x = rigaBase;    // Variabile base che definisce l'inizio del blocco (riga 71)
  const rangeDaVerificare = "C" + (x + 1) + ":I" + (x + 8); // A72:I79

  const range = shCF_Agenda.getRange(rangeDaVerificare);
  const arrayUnioni = range.getMergedRanges();

  // 2. Cicla e applica la logica di pulizia
  arrayUnioni.forEach(unione => {
    
    // Controllo che l'elemento non sia nullo/indefinito e abbia il metodo getValue
    if (unione && typeof unione.getValue === 'function') { 

      // Ottiene il valore e lo converte in stringa
      const valoreCella = String(unione.getValue());
      
      // Logica: Sciogli l'unione SE il valore NON contiene il simbolo "*"
      if (!valoreCella.includes("*")) {
        unione.breakApart();  // metodo per togliere le unioni
      }
    }
  });
}

// resetta il bordo
function reset_bordo_3(rigaBase) {
  const x = rigaBase ;  // Variabile base che definisce l'inizio del blocco (riga 71)
  
  
  const range_0 = "A" + x + ":I" + (x + 8);   // A71:I79
  const range_blu = "I" + x + ":I" + (x + 8);  // I71:I79

  const range_orizz = [   // [ "A71:I71", "A73:I73", "A75:I75", "A75:I75", "A79:I79" ]
    "A" + x + ":I" + x,          // A71:I71 (x+0)
    "A" + (x + 2) + ":I" + (x + 2),  // A73:I73 (x+2)
    "A" + (x + 6) + ":I" + (x + 6),  // A75:I75 (x+4)
    "A" + (x + 8) + ":I" + (x + 8)   // A79:I79 (x+8)
  ];
  
  const range_vert = [   // [ "B72:B79", "G72:G79"]
    "B" + (x) + ":B" + (x + 8),  // B72:B79 (x+1 a x+8)
    "G" + (x) + ":G" + (x + 8)   // G72:G79 (x+1 a x+8)
  ];



  // eseguiamo in sequenza i comandi: pulizia, formatt. orizz., formatt. vert, formatt. blu vert
  shCF_Agenda.getRange(range_0).setBorder(true, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  shCF_Agenda.getRangeList(range_orizz).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  shCF_Agenda.getRangeList(range_vert).setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  shCF_Agenda.getRange(range_blu).setBorder(null, null, null, true, null, null, 'blue', SpreadsheetApp.BorderStyle.SOLID_THICK);
}
// =====================================================================================================================================
// =====================================================================================================================================




function email_agenda() {  // trigger giornaliero
  let n_Agenda = ssCF.getRangeByName("dato_riga").getValues()[0][0];
  let n_Colonna = ssCF.getRangeByName("dato_colonna").getValues()[0][0];

  let colonna = 3 + n_Colonna;
  let riga = 5 + (n_Agenda-1)*12;
  
  let contenuto_legenda = legenda();
  let controllo = contenuto_legenda[0][0]; // #

  Logger.log("valore colonna: " + colonna);
  Logger.log("valore riga: " + riga);

  // let controllo = valore_controllo.charAt(0);
  let valore_cella = shCF_Agenda.getRange(riga, colonna).getValue();
  let valorePulito = String(valore_cella).substring(2);
  let nota_Corpo = shCF_Agenda.getRange(riga, colonna).getNote() || 'Nota assente';

  Logger.log("valore_cella= "+ valore_cella);

  try {
    //   Controlla se la cella contiene il carattere per inviare l'email, altrimenti non ha senso fare altro!
    if (String(valore_cella).includes(controllo)) { 

      for (let i = 1; i < contenuto_legenda.length; i++){  //   Controlla se la cella contiene la stessa tiologia
        let tipologia = contenuto_legenda[i][0];
        
        if(String(valore_cella).includes(tipologia)){

          const prima_parte_email = contenuto_legenda[i][1] ;

          const soggettoEmail = prima_parte_email + valorePulito;
          const corpoEmail = "Descrizione: " + nota_Corpo;
      
          // Invia la notifica via email
          MailApp.sendEmail(mia_Email, soggettoEmail, corpoEmail);
          Logger.log("Notifica email inviata con successo.");
        }
      }

    } else {
      Logger.log("Nessun simbolo " + controllo + " trovato per oggi.");
    }
    
  } catch (e) {
    Logger.log("Si è verificato un errore: " + e.toString());
  }

}







/**
 * Funzione principale per la formattazione automatica dell'agenda.
 * Gestisce input multipli in riga 1 e adatta il formato della cella sorgente.
 */
function formato_agenda() {
  const tabellaConversione = IntDen_convrs_data_coord;

  // 1) ESTRAZIONE DATI INPUT (E1)
  const valE1 = shCF_Agenda.getRange("E1").getValue().toString();
  // Estrae il riferimento cella (es. "C1") e il metodo ("da a" o "multi")
  const cellaBersaglioRef = valE1.match(/[A-Z]+[0-9]+/i) ? valE1.match(/[A-Z]+[0-9]+/i)[0] : "C1";
  const metodo = valE1.toLowerCase().includes("da a") ? "da a" : "multi";
  
  // Range sorgente (la cella con lo sfondo desiderato)
  const rangeSorgente = shCF_Agenda.getRange(cellaBersaglioRef);

  // 1.2) Estrazione date riga 1 (Piano B: Normalizzazione tramite funzione dedicata)
  let dateDaCercare = [];
  let colCounter = 6; // Parte da colonna F
  let valoreCella;
  
  do {
    valoreCella = shCF_Agenda.getRange(1, colCounter).getValue();
    if (valoreCella !== "" && valoreCella !== null) {
      // Usiamo la funzione di libreria per spezzare eventuali virgole e pulire spazi
      let datePulite = NormalizzaDateInput(valoreCella);
      dateDaCercare.push(...datePulite);
    }
    colCounter++;
  } while (valoreCella !== "" && colCounter < 15);

  if (dateDaCercare.length === 0) return;

  // 1.3) Preparazione Tabella per Map_CercaVert (Normalizzazione chiavi)
  const tabellaNormalizzata = tabellaConversione.map(riga => [
    riga[0].toString().toLowerCase().trim().replace(/\s+/g, ' '), 
    riga[1]
  ]);

  // 2) RICERCA COORDINATE
  const coordinateTrovate = Map_CercaVert(true, tabellaNormalizzata, dateDaCercare, 0, 1);

  if (!coordinateTrovate || coordinateTrovate.includes(null)) {
    SpreadsheetApp.getUi().alert("ERRORE: Una o più date non corrispondono alla tabella GESTIONE.");
    return;
  }

  // 3) APPLICAZIONE FORMATO
  const costFissa = 11; // Numero di righe extra da colorare (totale 12)

  if (metodo === "da a" && coordinateTrovate.length >= 2) {
    // Range unico: coord inizio -> coord fine
    const rInizio = shCF_Agenda.getRange(coordinateTrovate[0]);
    const rFine = shCF_Agenda.getRange(coordinateTrovate[coordinateTrovate.length - 1]);
    
    // Verifica se sono sulla stessa riga (settimana)
    if (rInizio.getRow() !== rFine.getRow()) {
      SpreadsheetApp.getUi().alert("ERRORE: La selezione 'da a' non può superare una singola settimana.");
      return;
    }

    const rangeDestinazione = shCF_Agenda.getRange(
      rInizio.getRow(), 
      rInizio.getColumn(), 
      costFissa + 1, 
      (rFine.getColumn() - rInizio.getColumn()) + 1
    );
    
    // Applichiamo la copia intelligente (adatta 1x1 a n*m)
    copiaFormatoAdattato(rangeSorgente, rangeDestinazione, "SFONDO");

  } else {
    // Metodo MULTI o data singola
    coordinateTrovate.forEach(coord => {
      if (coord) {
        const rangeDestinazione = shCF_Agenda.getRange(coord).offset(0, 0, costFissa + 1, 1);
        copiaFormatoAdattato(rangeSorgente, rangeDestinazione, "SFONDO");
      }
    });
  }

  // 4) RESET INPUT
  shCF_Agenda.getRange("E1").clearContent();
  shCF_Agenda.getRange(1, 6, 1, colCounter - 6).clearContent();
}