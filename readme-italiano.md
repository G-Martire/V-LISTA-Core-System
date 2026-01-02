# V-LISTA: Sistema di Gestione Cognitiva e Operativa

## I. Introduzione Filosofica: Il Perché del Sistema

Hai mai la sensazione di aver fatto poco, troppo o di non aver fatto le cose più importanti?
Il sistema **V-LISTA** nasce dall'esigenza di proteggere l'integrità logica individuale in un ambiente ad alto rumore (**High-Noise**). In un mondo che tenta di imporre una creatività forzata o una ripetitività alienante, V-LISTA funge da scudo operativo.
L'obiettivo non è "fare di più", ma **fare con consapevolezza**, minimizzando lo spreco di energia mentale (**Mental Watts**), tempo e denaro — risorse estremamente limitate.
Il fine ultimo è garantire che ogni risorsa spesa sia finalizzata esclusivamente a obiettivi di **classe A+++**.

### La Divisione dei Compiti (Uomo vs Macchina)
Il successo del metodo risiede in una rigorosa separazione delle responsabilità:

- **L'Operatore (Io):** Si focalizza sulla visione strategica, sulla creatività e sull'esecuzione dei task ad alto valore (Classe A+++).
- **Il Software (Google Sheets, GAS, Google Docs):** Si concentra nelle attività ripetitive dove l'essere umano è più inefficiente.

Il software si occupa di:
- **Sincronizzazione dati:** Riporto automatico delle note dai buffer di input.
- **Visualizzazione Dinamica:** Generazione automatica del calendario e delle priorità (V_LISTA).
- **Filtraggio:** Riduzione del rumore di fondo per permettere all'operatore di vedere solo ciò che conta.

------------------------------------------------------------------------------------------------------------

## II. Architettura del Sistema (V-LISTA-Core-System.xlsx)

Il sistema è progettato su tre livelli logici distinti:
1- **Agenda e Calendario:** Gestione delle attività vincolate temporalmente (orari fissi, eventi straordinari) e delle attività di "ricerca obbligatorie" (sport, paddle, palestra).
2- **V_LISTA e LISTA_IN:** Gestione delle attività **non** vincolate nel tempo. **LISTA_IN** funge da ingresso, mentre **V_LISTA** è l'agenda esecutiva.
3- **Masterplan (In fase di sviluppo):** Modulo per la progettazione a lungo termine (es. studio della lingua inglese per trasferimento all'estero).

### II.2 Criteri di Divisione e Parametri Operativi
La classificazione dei task avviene tramite una matrice di fattori:
- **Priorità (PR):** Indice di importanza strategica.
- **Lunghezza/Complessità (L.T.):** Supporto per note "radiali" con multi-collegamenti, superando la rigidità della scrittura lineare.
- **Tempo di Esecuzione (EXEC):** Stima temporale (30 min, 1 ora, più giorni).
- **Energia Mentale (WATT):** Carico cognitivo richiesto. Il sistema impedisce l'allocazione di task ad alto Watt (es. inglese) in fasce orarie a bassa energia.

------------------------------------------------------------------------------------------------------------

## III. Logica del Flusso Dati

Per garantire flessibilità e creatività, il sistema segue un flusso basato sull'**Analisi di Pareto**:
1. **INPUT_BUFFER (Google Docs):** Acquisizione radiale immediata (zero latenza) per note brevi e pensieri improvvisi.
2. **LISTA_IN (Google Sheets):** Importazione automatica. Se una nota supera i 300 caratteri, viene trasferita in **ELAB_EXPANSION_LOG** (Google Docs) per un'analisi approfondita.
3. **V_LISTA:** Fase finale di programmazione ed esecuzione.

> **Nota Tecnica:** L'intera infrastruttura è coordinata da script personalizzati in **Google Apps Script (.gs)** che permettono l'interconnessione tra i file e l'automazione dei calcoli di priorità.
