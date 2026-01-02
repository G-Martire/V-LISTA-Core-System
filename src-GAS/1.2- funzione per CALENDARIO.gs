// f1: r6; f2: r58; f3: 103; f4: 213; righe max: 387

// =====================================================================================================================================
// =====================================================================================================================================
// SCRIPT: import dati
function f1_dowload_celendario() {
  // elementi comuni
  let rigaInizio = 3;
  
  // file "gestione generale" (=GG)
  let elemento_GG = "AO";
  let date_GG = "AP";
  let verifica = "AS3";
  let elemVerifica = "non trascritto";
  let trascrizione = "trascritto";
  let ultimaRiga = "AS4";
  
  // file "cose da fare" (=CF)
  let elemento_CF = "AJ3";

  // calcoli
  let colonnaDestinazione = shCF_Gestione.getRange(elemento_CF).getValue();

  let colonnaVerifica = shGG_Gestione.getRange(verifica).getValue();
  let valoreUltimaRiga = shGG_Gestione.getRange(ultimaRiga).getValue();
  let numElementi = valoreUltimaRiga - rigaInizio + 1; // Calcola il numero di elementi da leggere

  // Itera sugli elementi
  for (let i = 0; i <numElementi; i++) {
    // ottenere il valore da verificare
    let verificare = shGG_Gestione.getRange((rigaInizio + i), colonnaVerifica, 1, 1).getValue();

    //step: 1(verificare); 2(leggere IN), 3 (leggere OUT), 4(scrivere in CF), 5(scivere "trascritto")
    if (verificare.toLowerCase() === elemVerifica.toLowerCase()) {
      // copia elemento nella colonna "AO"
      let valoreElemento = shGG_Gestione.getRange(elemento_GG + (rigaInizio + i)).getValue();
      // copia date nella colonna "AP"
      let valoreDate = shGG_Gestione.getRange(date_GG + (rigaInizio + i)).getValue();

      // scrivi i valori
      shCF_Gestione.getRange((rigaInizio + i), colonnaDestinazione, 1, 2).setValues([[valoreElemento, valoreDate]]);
      // impostare trascritto
      shGG_Gestione.getRange((rigaInizio + i), colonnaVerifica, 1, 1).setValue(trascrizione);
    }
  }

  Logger.log(" - ultriga: " + valoreUltimaRiga)
  Logger.log(" - num elementi: " + numElementi)
  Logger.log(" - val. verifica: " + verificare)
  Logger.log(" - val. elemento: " + valoreElemento + " -" + "val. date " + valoreDate)
}



// =====================================================================================================================================
// =====================================================================================================================================
// SCRIPT: METTERE I DATA A CALENDARIO (DA FARSI)
function f2_aggiornamento_calendario(){
  // posizione dei dati
  let colonnaElemento = "AB";
  let rigaInizio = 3;
  let colonnaRiga = "AD";
  let colonnnaFin = "AH";
  let colonnaVer = "AG";
  let verifica = "da cercare"

  // calcoli
  let ultimaRiga = shCF_Gestione.getLastRow(); // Ottieni l'ultimo numero di riga con dati nella colonna "AB"
  let numElementi = ultimaRiga - rigaInizio + 1; // Calcola il numero di elementi da leggere

  // Itera sugli elementi
  for (let i = 0; i < numElementi; i++) {
    // Ottieni il valore da verificare
    let colonnaVerValue = shCF_Gestione.getRange(colonnaVer + (rigaInizio + i)).getValue();

    if (colonnaVerValue.toLowerCase() === verifica.toLowerCase()) {  

      // Leggi l'elemento dalla colonna "AB"
      let elemento = shCF_Gestione.getRange(colonnaElemento + (rigaInizio + i)).getValue();

      // Leggi la riga di destinazione dalla colonna "AD"
      let rigaDestinazione = shCF_Gestione.getRange(colonnaRiga + (rigaInizio + i)).getValue();

      // Leggi la colonna di destinazione dalla colonna "AI"
      let colonnaDestinazione = shCF_Gestione.getRange(colonnnaFin + (rigaInizio + i)).getValue();

      // Scrivi l'elemento nella posizione specificata
      if (elemento !== "") { // Verifica se l'elemento non è vuoto
        // Converti il numero di colonna in lettera (es: 1 -> A, 2 -> B, ...)
        let letteraColonna = String.fromCharCode(64 + colonnaDestinazione);

        // Scrivi il valore nella cella di destinazione
        shCF_Calendario.getRange(letteraColonna + rigaDestinazione).setValue(elemento);
      }
    }   
  }
}



// =====================================================================================================================================
// =====================================================================================================================================
function f3_tabella_calendario() {
  let controllo = shCF_Agenda.getRange("O2").getValue();
  Logger.log(controllo);

  if (controllo !== "1- PC") {
    Logger.log("Il valore in X1 non è '1- PC'. Interruzione.");
    return;
  }

  // prendere i valori della tabella
  let ultimaRigaTC = shCF_Gestione.getRange("AT1").getValue();// ultima riga tabella calendario
  Logger.log("ultimaRigaTC: " + ultimaRigaTC);
  let dimensioni_Tabella = shCF_Gestione.getRange(3, 45, ultimaRigaTC, 3); // Colonna AS è la 45esima; il 4° valore indica il numero di tabelle della sidebord
  let valori = dimensioni_Tabella.getValues();
  let formats = dimensioni_Tabella.getNumberFormats(); //serve per dopo

  //____________________________________________________________________________________________________________________________________
  // [2] Modificare i valori della tabella
  let day1 = valori[1][1].getDate(); //modifica formato (solo numero)
  let day2 = valori[2][1].getDate(); //modifica formato (solo numero)

  // Formatta le date nella colonna 1 (indice 1), inizianzo dalla riga 4
  const optionsGiorno = { weekday: 'long' };
  for (let i = 4; i < valori.length; i++) {
    if (valori[i][1] instanceof Date) {
      valori[i][1] = valori[i][1].toLocaleDateString('it-IT', optionsGiorno);
    }
  }

  // reimpostare alcuni campi della tabella
  valori[0][0] = "Sett. Vis.";
  valori[0][1] = Math.floor(valori[0][1] / 12)+1;

  Logger.log(valori[0][1]);
  valori.splice(1, 2) // eliminare le righe 2 e 3

  // Aggiungi una colonna vuota per le caselle di controllo
  valori[0].unshift(""); // Aggiunge un'intestazione vuota
  // Non aggiungere la checkbox alla prima riga (intestazione)
  for (let i = 1; i < valori.length; i++) {
    valori[i].unshift('<input type="checkbox">'); // Aggiunge il tag HTML per la checkbox
  }

  //____________________________________________________________________________________________________________________________________
  // [3] Creazione pagina HTML //modficato
  let htmlTable = "<!DOCTYPE html>";
  htmlTable += "<html>";
  htmlTable += "<head>";
  htmlTable += "<title>Tabella</title>";
  htmlTable += "<style>";
  htmlTable += "body { font-family: sans-serif; margin: 0; }";
  htmlTable += ".sidebar { position: fixed; top: 0; right: 0; bottom: 0; width: 350px; background-color: #f4f4f4; border-left: 1px solid #ccc; padding: 20px; box-sizing: border-box; overflow-y: auto; padding-left: 55px; padding-right: 5px; }"; // Aggiunto padding-right //modficato // Larghezza del sidebar riportata a un valore più contenuto
  htmlTable += "table { width: auto; border-collapse: collapse; font-size: 10pt; }"; // width: auto per permettere alla tabella di espandersi
  htmlTable += "th, td { border: 1px solid #ddd; padding: 6px; text-align: left; white-space: normal; word-break: break-word; }";
  htmlTable += "th { background-color: #f0f0f0; }";
  htmlTable += "/* Stili per la larghezza delle colonne */"; //modficato
  htmlTable += "th:nth-child(1), td:nth-child(1) { width: 30px; } /* Colonna checkbox */";
  htmlTable += "th:nth-child(2), td:nth-child(2) { width: 250px; } /* Colonna 'Dettaglio' (allargata) */";
  htmlTable += "th:nth-child(3), td:nth-child(3) { width: 70px; } /* Colonna 'Giorno' (ristretta) */";
  htmlTable += "</style>";
  htmlTable += "</head>";
  htmlTable += "<body>";
  htmlTable += "<div class='sidebar'>";
  htmlTable += "<h2>Settimana " + day1 + "-" + day2 + "</h2>";
  htmlTable += "<div style='overflow-x: auto;'>"; // Assicurati che questo div sia presente
  htmlTable += "<table>";

  //____________________________________________________________________________________________________________________________________
  // [4] Popolare la pagina con la struttura della tabella //modficato

  if (valori.length > 0) {
    // Aggiungi l'intestazione (opzionale)
    htmlTable += "<tr><th>" + valori[0][0] + "</th><th>" + valori[0][1] + "</th><th></th></tr>"; // Aggiunta un'intestazione vuota per la checkbox

    for (let i = 1; i < valori.length; i++) {
      htmlTable += "<tr>";
      for (let j = 0; j < valori[i].length; j++) {
        let cellValue = valori[i][j];
        let cellFormat = formats[i] ? formats[i][j] : null; // Verifica che formats[i] esista

        // Formatta le date in modo più leggibile se il formato lo suggerisce
        if (cellValue instanceof Date) {
          const options = { year: 'numeric', month: 'short', day: 'numeric' };
          cellValue = cellValue.toLocaleDateString('it-IT', options);
        }

        htmlTable += "<td>" + cellValue + "</td>";
      }
      htmlTable += "</tr>";
    }
  } else {
    htmlTable += "<tr><td>Nessun dato trovato</td></tr>";
  }

  htmlTable += "</table>";
  htmlTable += "</div>";
  htmlTable += "</body>";
  htmlTable += "</html>";

  let htmlOutput = HtmlService.createHtmlOutput(htmlTable) //valori di riferimento
      .setWidth(350)
      .setHeight(600);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}



// =====================================================================================================================================
// =====================================================================================================================================
// VERSIONE: 17/5/25 - 16.41
function f4_1_mostraFinestraDisponibilita() {
  // dati in ingresso
  const intestazione = shCF_Calendario.getRange("C2:N2").getValues()[0];
  const data_Rif = +shCF_Gestione.getRange("AX4").getValue();
  const xy_data_Rif = +shCF_Gestione.getRange("AX5").getValue();
  const data_IN = +shCF_Gestione.getRange("AX2").getValue();
  const data_OUT = +shCF_Gestione.getRange("AX3").getValue();

  // calcoli range tabella
  const tb_riga_Iniz = data_IN - data_Rif + xy_data_Rif -1;
  const tb_riga_Fin = data_OUT - data_IN + 2;
  const tabella_IN = shCF_Calendario.getRange(tb_riga_Iniz, 3, tb_riga_Fin, 12).getValues();

  Logger.log("riga iniziale: " + tb_riga_Iniz);
  Logger.log("riga finale: " + tb_riga_Fin);

  const htmlContent = f4_2_creaTabellaHTML (intestazione, tabella_IN, tb_riga_Iniz);
  
  const htmlOutput = HtmlService.createHtmlOutput (htmlContent)
    .setWidth(1200)
    .setHeight (900);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Tabella Calendario');

  Logger.log(tabella_IN);

  const array_OUT = tb_riga_Iniz
  return array_OUT;
}

function f4_2_creaTabellaHTML(headerData, tabella_IN, startRow) {
  const giorniSettimana = ['Dom', 'Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab'];

  const datiTabella = tabella_IN.slice(1).map(riga => {
    // Formatta la prima cella (data)
    const dataOriginale = riga[0];
    let dataFormattata = '';
    if (dataOriginale instanceof Date) {
      const giornoSettimana = giorniSettimana[dataOriginale.getDay()];
      const giorno = dataOriginale.getDate();
      const mese = dataOriginale.getMonth() + 1;
      dataFormattata = `${giornoSettimana} ${giorno}/${mese.toString().padStart(2, "0")}`;
    } else {
      dataFormattata = dataOriginale;
    }
    const rigaModificata = [dataFormattata, ...riga.slice(1)];
    return rigaModificata;
  });

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <title>Tabella Calendario</title>
      <style>
        table {
          border-collapse: collapse;
          width: 100%;
        }
        th, td {
          border: 1px solid black;
          padding: 8px;
          text-align: center;
        }
        th {
          background-color: #f2f2f2;
          position: sticky; /* Mantiene l'intestazione fissa */
          top: 0;
          z-index: 1;
          background-color: #f2f2f2;
        }
        td input {
          width: 100%;
          height: 100%;
          border: none;
          text-align: center;
          padding: 0;
          font-size: 1em;
        }
        td:first-child input { /* Stile per la prima colonna (data) */
          width: auto; /* Adatta la larghezza al contenuto */
          background-color: #f0f0f0; /* Sfondo grigio per indicare che non è modificabile */
          color: #808080;
        }
        td:first-child {
          width: 100px;
        }
        th:nth-child(1),
        td:nth-child(1),
        th:nth-child(3),
        td:nth-child(3),
        th:nth-child(5),
        td:nth-child(5),
        th:nth-child(6),
        td:nth-child(6),
        th:nth-child(9),
        td:nth-child(9),
        th:nth-child(10),
        td:nth-child(10) {
          border-right: 3px solid black;
        }
        #salvaModifiche {
          margin-top: 20px;
          padding: 10px 20px;
          background-color: #4CAF50;
          color: white;
          border: none;
          cursor: pointer;
          border-radius: 5px;
          font-size: 1em;
          display: block;
          margin-left: auto;
          margin-right: auto;
        }
        #salvaModifiche:hover {
          background-color: #367c39;
        }
      </style>
      <script>
        let startRow = ${startRow}; // Usa una letiabile globale per passare startRow

        function salvaModifiche() {
          const table = document.getElementById("myTable");
          const rowCount = table.rows.length;
          const colCount = table.rows[0].cells.length;
          const data = [];

          // Ottieni i dati dalla tabella HTML
          for (let i = 1; i < rowCount; i++) {
            const rowData = [];
            for (let j = 1; j < colCount; j++) {
              const inputElement = table.rows[i].cells[j].querySelector('input');
              const cellValue = inputElement ? inputElement.value : table.rows[i].cells[j].textContent;
              rowData.push(cellValue);
            }
            data.push(rowData);
          }
          // Chiama la funzione di Apps Script per sallete i dati e poi chiudi la finestra
          google.script.run.f4_3_salvaDatiModificati(data);
        }
      </script>
    </head>
    <body>
      <table id="myTable">
        <thead>
          <tr>${headerData.map(header => `<th>${header}</th>`).join("")}</tr>
        </thead>
        <tbody>
          ${datiTabella.map(riga => `
            <tr>
              ${riga.map((cella, cellIndex) => {
                if (cellIndex === 0) {
                  return `<td>${cella}</td>`; // La prima colonna (data) è solo testo
                } else {
                  return `<td><input type="text" value="${cella}"></td>`; // Le altre colonne sono modificabili
                }
              }).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>
      <button id="salvaModifiche" onclick="salvaModifiche()">Salva Modifiche</button>
    </body>
    </html>
  `;
}

function f4_3_salvaDatiModificati(data) {
  const array_IN = f4_1_mostraFinestraDisponibilita(); // ATTENZIONE: Questa funzione mostra una finestra!
  const startRow= array_IN+1;
  const startColumn = 4;
  shCF_Calendario.getRange (startRow, startColumn, data.length, data[0].length).setValues(data);
}

// =====================================================================================================================================

// =====================================================================================================================================
