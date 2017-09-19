/* 

ANALISI
Il database è implementato su un file in formato gsheet, alcuni fogli di tale files rappresentano le tabelle (relazioni): 
  - Files affidi
  - Diffide da inviare
  - Dettaglio fatture

Entità 
1) Files Affidi
   si tratta di files in formato gsheet contenenti 3 fogli (CASINOFATTURE, CASIFATTURE, INFO CAMERALI)
   i files sono conservati in una cartella di Google Drive
   con il comando "aggiorna" del panel Files Affidi, viene letta la cartella in cerca di nuovi files da 'importare' nel foglio Files Affidi
   quando un file viene inserito in tabella assume lo stato 'Assegnato' e marcato con la data di importazione del file e il tag 'new' 
   quando le pratiche presenti nel file vengono importate nella tabella Diffide da inviare (che sarebbe più corretto chiamare Pratiche), il file assume lo stato 'Importato'
   forse sarebbe meglio chiamarlo 'Esportato' nel senso che le pratiche in esso presenti vengono esportate nella tabella Diffide da inviare

2) Diffide da inviare (Pratiche)
   si tratta della testata pratica 
   come stato iniziale ha importata
   può assumere gli stato inviata, esportata (va aggiunto uno stato 'archiviata')

3) Dettaglio fatture
   si tratta delle singole fatture presenti nella pratica
   ogni fattura contiene la chiave esterna che l'associa alla testata pratica corrispondente
   
   
   
// importazione delle pratiche presenti sui files 
    // da priorità agli indirizzi presenti su Info Camerali
    // determina il flusso di appartenenza (IOL, MP, ecc) e di conseguenza determina il numero progressivo
    // scrive le pratiche sui fogli 'Diffide da Inviare' e 'Dettaglio fatture' previo eventuale accorpamento dei ratei fatture
    
// stampa diffide 
   // permette di selezionare le pratiche desiderate per stampare le lettere di diffida
   // stampa le diffide in un unico documento in formato gdoc a partire da un template dedicato
   // stampa la lista di controllo in un altro documento in formato gdoc a partire da un template dedicato
   
// esporta su file per gestionale Fabrizio 
    // esporta tutte le pratiche che si trovano nello stato 'Inviata' ad esclusione di quelle con protocollo IOL
    // sovrascrive 4 files in formato gdoc che poi vanno convertiti manualmente in formato xls
    
// inserisce pratiche extra flusso con protocollo IOL e indicazione della tipologia (es. OPPOSIZIONE, FALLIMENTO, ecc)


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
- Liste di esportazione ( in attesa di esito da Fabrizio) - fatto
15/05/2017
- Controllo prima di esportazione (verifica che le pratiche siano già inviate) - fatto
16/05/2017
- Sostituisce Object all'Array nell'inserimento ex-flusso (in modo da non risentire della modifica del tracciato del DB) - fatto
19/09/2017
- Aggiungere controllo per le spese legali su pratiche IOL con importo tra 0  e 3000 - fatto


- funzione di eliminazione ultima importazione (cancella testata pratiche e dettagli fatture)
- funzione di eliminazione ultima esportazione (ripristina lo stato ad 'Inviata')
- funzione di esportazione autonoma (selezionare le pratiche da esportare)
- alert in caso di difformità tra importo totale e fatture
- gestione ordine di importazione dei files in modo batch
- Formato importi con virgola e doppia cifra digitale - complicato
- DUPLICATI - importare, progressivare, segnalare
- lista di controllo in orizzontale unire prime due colonne e aggiungere colonna Descrizione
- controllo correttezza importazione diffide (totale pratiche, totale fatture ?, totale importo scoperto)
- archiviazione pratiche (probabilmente necessarie per ridurre i tempi di caricamento)

Minori

- indicazione utenti connessi
- contatore numero documenti stampati in giornata
- date di importazione e di stampa fissate al momento del lancio del comando


QUOTAS
Documents created	250 / day
Email recipients per day	100* / day
*/

