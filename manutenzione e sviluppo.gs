/* 

ANALISI

// legge la cartella Files Affidi da Importare
// per ogni file presente procede con l'importazione su foglio 'Diffide da Inviare' (con indicazione di data di importazione)
// stampa diffide
// stampa etichette
// sposta Diffide da Inviare su foglio Diffide inviate (con indicazione di data di importazione lotto e data di invio diffide)
// cancella foglio Diffide da inviare
// produce tabelle per gestionale Fabrizio


Correzione BUGS 

09/01/2017 
 - a fine procedura non viene aggiornato lo stato del file importato -- risolto
 - selezionando più file da importare viene preso in considerazione solo l'ultimo - risolto
11/01/2017
 - se si stampano diffide appartenenti a file diversi, la tipologia flusso sulla stampa risulta errata - risolto
 - dopo la stampa delle diffide cancellare la selezione (altrimenti ricliccando su stampa vengono prodotti dei documenti scorretti) - risolto 
 - prima riga delle etichette troppo bassa. Provare a cancellare l'header da script 
 18/01/2017
 - sviluppare controllo sull'ultima riga vuota e in mancanza aggiungerla. Ciò perchè attenzione se gli spreadsheet non hanno righe vuote sotto i dati si può incorrere in un errore di tipo (Service Error: Spreadsheet).

SVILUPPO

11/01/2017
- inserire data importazione su foglio dettaglio fattura - fatto
- le fatture sulle diffide vanno ordinate per data di emissione - fatto
- per lo start-up impostare inizio riferimento pratica corretto (es. 6526/MP ) - fatto
- file unico con più flussi da suddividere in base a importo totale pratica (MCR 500 - 999, MP 1000 - 2999, IOL >2999) - fatto
- Stampa lista di controllo in coda alla stampa diffide. Rif Pratica, Tipologia flusso, Codice cliente, Ragione sociale, Indirizzo - fatto
- template diffide spese legali aggiugere 00 dopo la virgola - fatto
- utilizzare la colonna stato affido per determinare il flusso - fatto
- gestione inserimenti exFlusso - OPPOSIZIONI, FALLIMENTI, CONCORDATI gestione ordinaria - importati IOL, progressivati, segnalati, no stampa diffida - fatto
- escludere dalla stampa diffide le pratiche exFlusso (possibilmente in fase di selezione sul browser) - fatto
- Gestione dato fiscale - fatto 
- Liste di esportazione ( in attesa di esito da Fabrizio)
- alert in caso di difformità tra importo totle e fatture
- gestione ordine di importazione dei files in modo batch
- Formato importi con virgola e doppia cifra digitale - complicato
- DUPLICATI - importare, progressivare, segnalare
- lista di controllo in orizzontale unire prime due colonne e aggiungere colonna Descrizione

Minori
- indicazione utenti connessi
- contatore numero documenti stampati in giornata
- date di importazione e di stampa fissate al momento del lancio del comando




QUOTAS
Documents created	250 / day
Email recipients per day	100* / day
*/

