// SCRIPT ATTUALI: 2 (1 da sistemare) + 1 (da farsi)
// coordinate dati comuni FILE "cose da fare" (=CF)
const ssCF = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID_HERE");
const shCF_Listone = ssCF.getSheetByName("LISTONE");
const shCF_Agenda = ssCF.getSheetByName("AGENDA");
const shCF_Gestione = ssCF.getSheetByName("GESTIONE");
const shCF_Calendario = ssCF.getSheetByName("Calendario");
const shCF_Lista_in = ssCF.getSheetByName("LISTA_IN");
// Intervalli denominati del FILE "cose da fare" (=CF)
const IntDen_Legenda = ssCF.getRangeByName("LEGENDA").getValues(); // Intervallo Denominato Legenda
const IntDen_convrs_data_coord = ssCF.getRangeByName("conversione_data_coordinata").getValues(); // Intervallo Denominato sul foglio "gestione"



// coordinate dati comuni FILE "gestione generale" (=GG)
const ssGG = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID_HERE");
const shGG_Gestione = ssGG.getSheetByName("GESTIONE");
const shGG_Generale = ssGG.getSheetByName("Generale");



// coordinate dati comuni FILE - import dati (note) file Google Docs
const docword = DocumentApp.openById('YOUR_DOCUMENT_ID_HERE');
const ID_DOC_ARCHIVIO = 'YOUR_ARCHIVE_DOC_ID_HERE';



// altri dati
const mia_Email = "giumartex@gmail.com";







/* LEGGENDA
- (da mettere in graffa {})
- lettura in = lettura in entrata (testa)
- lettura during = lettura durante l'esecuzione del codice (corpo - in)
- esec during = esecuzione di parti di codice durante l'esecuzione di altre parti del codice (corpo - out)
- esec out = esecuzione in uscita (coda)
- calcolo
*/



// =====================================================================================================================================
// =====================================================================================================================================
// la funzione esporta l'Intervallo in un file JSON
function esportaAgendaInJSON() {
  let ss = SpreadsheetApp.getActiveSpreadsheet(); // Ottieni il riferimento al foglio di lavoro attivo
  let sheet = ss.getSheetByName("AGENDA"); // Ottieni il foglio specifico per nome
  if (!sheet) {
    Logger.log("Errore: Foglio con nome 'shCF_Agenda' non trovato.");
    return; // Esci dalla funzione se il foglio non esiste
  }
  let dataRange = sheet.getRange("A1:I48");
  let values = dataRange.getValues();

  let header = values[0];
  let jsonData = [];
  for (let i = 1; i < values.length; i++) {
    let rowData = {};
    for (let j = 0; j < header.length; j++) {
      rowData[header[j]] = values[i][j];
    }
    jsonData.push(rowData);
  }

  let jsonString = JSON.stringify(jsonData, null, 2); // Il '2' aggiunge un'indentazione per la leggibilità

  let fileName = "agenda_" + Utilities.formatDate(new Date(), Session.getTimeZone(), "yyyyMMdd_HHmmss") + ".json";
  let file = DriveApp.createFile(fileName, jsonString, "application/json");

  Logger.log("File JSON creato con successo su Google Drive: " + file.getUrl());
}





/**
 * Copia i formati tra due Range. Adatta lo sfondo se la sorgente è mono-colonna.
 *
 * @return Array_Agenda[simbolo][testo email]
 * Indice di Mappatura (Proprietà):
 * 0: Logica di invio email (simbolo unico)
 * 1: Impegno standard ITS
 * 2: Impegno eccezzionale ITS
 * 3: Campo vuoto / Default
 * 4: Novità hobby
 * 5: Altre attività
 * 6: Compleanno
 * 7: Svago/Riposo
 */
function legenda (){
  const Array_Agenda =  [  // la colonna 1 = simbolo; colonna 2= testo EMAIL (fa eccezione ovviamente la proprietà 1)
    [IntDen_Legenda[0][0].charAt(0),], // proprietà 1: invio email: serve alla logica per capire se inviare email o no
    [IntDen_Legenda[0][1].charAt(0),"Impegno standard ITS: "], // proprietà 2: campo
    [IntDen_Legenda[0][2].charAt(0),"Impegno eccezionale ITS: "], // proprietà 2: campo
    [IntDen_Legenda[0][3].charAt(0),"Campo vuoto "], // proprietà 2: campo
    [IntDen_Legenda[0][4].charAt(0),"Novità hobby: "], // proprietà 2: campo
    [IntDen_Legenda[0][5].charAt(0),"Altre mie attività: "], // proprietà 2: campo
    [IntDen_Legenda[0][6].charAt(0),"Compleanno di: "], // proprietà 2: campo
    [IntDen_Legenda[0][7].charAt(0),"Attività svago/riposo di: "]  // proprietà 2: campo
  ];

  return Array_Agenda
}