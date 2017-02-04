// crea elenco Files Affidi da Importare sullo sheet
// e restituisce numero files presenti nella folder 
// e numero files nuovi scritti sullo sheet

function writeFilesToSheet() {
    Logger.log('writeFilesToSheet')
    
    var files = affidiDaImportareFolder.getFiles()
    var file, data 
    var sheet = sheetFilesAffidi;
    var alreadyOnSheet = false
    var onFolder = 0
    var lastRow = sheet.getLastRow()
    var wroteOnSheet = lastRow -1
    var row
    var numberOfFilesInFolder = countFilesInFolder(affidiDaImportareFolderId)
    var dataAssegnazione = Utilities.formatDate(new Date(), 'CET', 'dd/MM/YYYY HH.mm.ss')
    //itera lungo i file trovati sulla folder
    while (files.hasNext()) {
            file = files.next();
            onFolder++
            alreadyOnSheet = false
             // imposta il valore di una nuova riga dello sheet
            //con i dati del file sulla folder
            var name = file.getName()

            data = [ 
              name,
              file.getDateCreated(),
              file.getUrl(),
              "Assegnato", //Stato
              dataAssegnazione, 
              ""
            ];
        
            Logger.log(data)
            Logger.log('numberOfFilesInFolder ' + numberOfFilesInFolder)
            Logger.log('alreadyOnSheet ' + alreadyOnSheet)
            Logger.log('wroteOnSheet ' + wroteOnSheet)
            
             //verifica se il file presente sulla folder è già presente sullo sheet
             // se non ci sono file già scritti sullo sheet salta al prossimo file sulla folder
              for (row=1; row<=wroteOnSheet; row++){  
                Logger.log('wroteOnSheet ' + wroteOnSheet)
                Logger.log(data[2])
                Logger.log(sheet.getName())
              // se il file della folder è gia sullo sheet pulisci colonna Badge
              Logger.log('data[2] ' + data[2])
              Logger.log('sheet col 3' + sheet.getRange(row+1,3).getValue())

              // controlla che il file l'url del file letto dalla tabella è lo stesso di quello presente sullo sheet
              if (data[2] == sheet.getRange(row+1,3).getValue()){
                            // memorizza che il file era già sullo sheet
                            alreadyOnSheet = true
                            // pulisce eventuale 'new' su Badge
                            //sheet.getRange(row+1,7).setValue('')
                        }
                  }
              if(!alreadyOnSheet && (numberOfFilesInFolder > wroteOnSheet) || wroteOnSheet == 0){
                  // scrive il file sullo sheet
                  data[6] = 'new' // Badge
                  Logger.log('data ' + data)
                  Logger.log(sheet.getName())
                  sheet.appendRow(data);
                  wroteOnSheet++ 
              }
     }           
        //torna su flow (readFilesAffidiFromFolder) 
}

 
function readAffido(ssAffido){

  
Logger.log('readAffido')
Logger.log(ssAffido.getUrl())

// legge le diffide dal file di affido, le protocolla e controlla i duplicati
// Legge e mette in relazione i 3 fogli del file affido

 var newIdDiffida
 var URLFileAffido = ssAffido.getUrl()
 var nomeFileAffido = ssAffido.getName()
 var sheet = ssAffido.getSheetByName('CASI NO FATTURE')
 var objCasiNoFatture = grabObjectFromSheet(sheet)

 var sheet = ssAffido.getSheetByName('CASI FATTURE')
 var objCasiFatture = grabObjectFromSheet(sheet)

 var sheet = ssAffido.getSheetByName('VISURA CAMERALE CUSTOMER')
 var objVisureCamerali = grabObjectFromSheet(sheet)
 Logger.log(objVisureCamerali)
 
 var newRiferimentoPratica
 var offsetRiferimentoPratica

 //crea array di oggetti  objDiffideDaInviare
 
 // inizia con Casi NO Fatture
 var objDiffideDaImportare = []
 var tipoFlusso 
 var errors = 0
 var arrayObjErrors=[]
 var objErrors
 var statoAffido, tipologiaPraticaWf

 
 // sheet diffideDaInviare
 var rowDiffide = sheetDiffideDaInviare.getLastRow()
 var lastColDiffide = sheetDiffideDaInviare.getLastColumn();
 var headers = sheetDiffideDaInviare.getRange(1,1,1,lastColDiffide).getValues();
     Logger.log('headers ' + headers)
 var speseLegali
 var dataImportazione = Utilities.formatDate(new Date(), 'CET', 'dd/MM/YYYY HH.mm.ss')
 
 // sheet DettaglioFatture
 var rifPraticaFlusso, dataFattura, dateTimeFattura
 var rowFatture = sheetDettaglioFatture.getLastRow() + 1
 var lastColDettaglioFatture = sheetDettaglioFatture.getLastColumn()
 Logger.log('lastColDettaglioFatture ' + lastColDettaglioFatture)
     
 // ciclo i (objCasiNoFatture)
 for (var i=0; i<objCasiNoFatture.length; i++){
      tipoFlusso = '***'
   Logger.log('i = ' + i)
     //incrementa ID e riferimento pratica 
     newIdDiffida = Utilities.formatDate(new Date(), 'CET', 'YYYYMMddHHmmss')
     //crea oggetto relativo a pratica protocollata
     var importoScoperto = objCasiNoFatture[i].importoScoperto

  
     
    // gestisce gli le pratiche EX FLUSSO
    
    tipologiaPraticaWf = objCasiNoFatture[i].tipologiaPraticaWf
    switch (true) {
        
      case (tipologiaPraticaWf == 'OPPOSIZIONE' || tipologiaPraticaWf == 'FALLIMENTO' || tipologiaPraticaWf == 'CONCORDATO'):
            statoAffido = 'AFFIDATA_AVV_ORD'
            break;
      default:
            var statoAffido = objCasiNoFatture[i].statoAffido
            tipologiaPraticaWf = ''
            break;
     }   
     Logger.log("Stato affido " + statoAffido)
     // gestisce lo stato affido da file
     switch (true) {
       case (statoAffido == 'AFFIDATA_AVV_MCR' ):
         tipoFlusso = 'MCR'
         offsetRiferimentoPratica = 2127
         speseLegali = 100.00
         break;
       case (statoAffido == 'AFFIDATA_AVV_ST'):
         tipoFlusso = 'MP'
         offsetRiferimentoPratica = 6528
         speseLegali = 200.00
         break;
       case (statoAffido == 'AFFIDATA_AVV_ORD'):
         tipoFlusso = 'IOL'
         offsetRiferimentoPratica = 7219
         switch (true) {
           case (importoScoperto<10000):
             speseLegali = 300.00
             break;
           case (importoScoperto<20000):
             speseLegali = 400.00
             break;
           case (importoScoperto>20000):
             speseLegali = 500.00
             break;
           default:
             break;
         }
       default:
         break;
     }       
          
     Logger.log(tipoFlusso)
     // restituisce un querySheet ossia lo sheet che contiene le sole diffide relative al tipoflusso
     var querySheetDiffide = querySheet(tipoFlusso)  
     Logger.log(querySheetDiffide.getName())
     //legge l'ultimo protocollo
     var lastRowPratica = querySheetDiffide.getLastRow()
     Logger.log('lastRowPratica = ' + lastRowPratica)
     if (lastRowPratica == 1){
       // se è la prima protocollazione aggiunge l'offset 
       var newRiferimentoPratica = 1 + offsetRiferimentoPratica; 
     }
     else
     {
     Logger.log(querySheetDiffide.getRange(lastRowPratica,2).getValue())
        var newRiferimentoPratica = querySheetDiffide.getRange(lastRowPratica,2).getValue()+1 
     }
    
     var lastColQuerySheet = querySheetDiffide.getLastColumn()
     Logger.log("Colonne di Diffide da Inviare " + lastColQuerySheet)
     
     var headers = querySheetDiffide.getRange(1,1,1,lastColQuerySheet).getValues()
     Logger.log('new Riferimento Pratica  = ' + newRiferimentoPratica)
     Logger.log(tipoFlusso)

     objDiffideDaImportare[i]={
     
         'ID diffida' : newIdDiffida,
         'Riferimento pratica': newRiferimentoPratica,
         'Tipologia flusso': tipoFlusso,
         'Stato affido': statoAffido,
         'Tipologia pratica': tipologiaPraticaWf,
         'Nome file affido': nomeFileAffido,
         'URL file affido': URLFileAffido,
         'Codice cliente': objCasiNoFatture[i].codcliente,
         'Dato fiscale': objCasiNoFatture[i].datoFiscale,
         'Ragione sociale': objCasiNoFatture[i].ragioneSociale,
         'Indirizzo': objCasiNoFatture[i].indirizzoResidenza,
         'CAP': objCasiNoFatture[i].capResidenza,
         'Comune' : objCasiNoFatture[i].comuneResidenza,
         'Provincia': objCasiNoFatture[i].provinciaResidenza,
         'Telefono':objCasiNoFatture[i].telefono,
         'Provenienza indirizzo': 'CACS',
         'Importo totale': objCasiNoFatture[i].importoScoperto, // l'importo totale è quello del file CasiNoFatture !!
         'Data importazione': dataImportazione,
         'Stato': 'Importata'
       
     }
     // cerca info camerali di Casi No Fatture    
     // ATTENZIONE il match è effettuato  tra il dato fiscale di CASI NO FATTURE e partita IVA o Codice FiscaLE di INFO VISURE CAMERALI
     // in quanto il codice cliente su Visure Camerali non corrisponde
     for (var j in objVisureCamerali){
           if (objCasiNoFatture[i].datoFiscale === objVisureCamerali[j].piva || objCasiNoFatture[i].datoFiscale === objVisureCamerali[j].codiceFiscale){  
                 objDiffideDaImportare[i]['Indirizzo'] = objVisureCamerali[j].indirizzo
                 objDiffideDaImportare[i]['CAP'] = objVisureCamerali[j].cap
                 objDiffideDaImportare[i]['Comune'] = objVisureCamerali[j].comune
                 objDiffideDaImportare[i]['Provincia'] = objVisureCamerali[j].provincia
                 objDiffideDaImportare[i]['Telefono'] = objCasiNoFatture[i].telefono
                 objDiffideDaImportare[i]['Provenienza indirizzo'] = 'Info camerali'  
                 //objDiffideDaImportare[i]['Data importazione'] = dataImportazione
                 //objDiffideDaImportare[i]['Stato'] = 'Importata'  
           } 
     }
    // inizializza variabili relative a fatture
    var progressivoFattura = 0
    var importoTotale = 0
    
    //crea array interno per la proprietà fatture
    // filtra i objCasiFatture in base a codice cliente
    Logger.log('length objCasiFatture ' + objCasiFatture.length)
    var fatture = []
    for (var z=0; z<objCasiFatture.length; z++){
          rifPraticaFlusso = newRiferimentoPratica + "/" + tipoFlusso
          // inizializza il progressivo fattura
          Logger.log('codClienteFatture ' + objCasiFatture[z].codcliente + "--->" + 'codClienteDiffide ' + objCasiNoFatture[i].codcliente)
          if (objCasiFatture[z].codcliente === objCasiNoFatture[i].codcliente){
               progressivoFattura++
               var dataImportazioneFattura = dataImportazione
               if (objCasiFatture[z].dataFattura){
                 dataFattura = new Date(objCasiFatture[z].dataFattura) 
                 dateTimeFattura = dataFattura.getTime() // non attribuisce un valore cronologicamente congruente alla data !!!
               }
               else
               {
                 dataFattura = ''
                 dateTimeFattura = 1 // inserisce un datatime fittizio per ordinare le fatture (poi verrà eliminato)
               }
               // compone l'array con le fatture, inserisce anche il numero progressivo di fattura (z)
               fatture.push([dateTimeFattura, newIdDiffida, rifPraticaFlusso, objCasiFatture[z].codcliente,objCasiFatture[z].numeroFattura,objCasiFatture[z].dataFattura, objCasiFatture[z].importoScoperto, dataImportazione]) 
               //importoTotale += objCasiFatture[z].importoScoperto
          }   
     }

     Logger.log('fatture con shift e sort by date ' + JSON.stringify(fatture)) 
     // finalizza l'oggetto diffida
     objDiffideDaImportare[i]['Fatture presenti'] = progressivoFattura
     //objDiffideDaImportare[i]['Importo totale'] = importoTotale
     objDiffideDaImportare[i]['Spese legali'] = speseLegali
     
     // ordina fatture per data di emissione
      fatture.sort(function(a,b) {
        return a[0]-b[0]
      });
     Logger.log('fatture ordinate per data ' + JSON.stringify(fatture))
     // elimina il primo elemento da tutti gli array
     fatture.map(function(val){
       return val.shift(0);
     });
     
   Logger.log(fatture) // il Codice Cliente è ancora correttamente formattato come testo 
     // inserisce l'array 2D fatture nella proprietà Fatture come stringa 
     objDiffideDaImportare[i]['Fatture'] = fatture
     Logger.log(objDiffideDaImportare[i]['Fatture'])  // il Codice Cliente è ancora correttamente formattato come testo 

     // scrive gli affidi sul foglio Diffide Da Inviare
     
       rowDiffide++
       // scrive i dati sul foglio Diffide da inviare 
        for (var prop in objDiffideDaImportare[i]){ 

          Logger.log('prop ' + prop)
          // scrive le fatture in formato stringa in una colonna del foglio Diffide da Inviare
          if (prop == 'Fatture'){ 
            var fattureString = JSON.stringify(objDiffideDaImportare[i][prop])
            Logger.log(fattureString)
            Logger.log('fatture da scrivere nel foglio dettaglio fatture ')
            
            var index = headers[0].indexOf('Fatture')
            Logger.log('index Fatture ' + index)
            sheetDiffideDaInviare.getRange(rowDiffide,index+1).setValue(fattureString)
          }
          else
          {
            // scrive su tutte le colonne del foglio Diffide da inviare tranne l'ultima  
            var index = headers[0].indexOf(prop)
            if (index >=0 ){
              sheetDiffideDaInviare.getRange(rowDiffide,index+1).setValue(objDiffideDaImportare[i][prop])
              Utilities.sleep(50)
            }
          }
         }
        // per la proprietà 'Fatture' scrive i dati delle fatture sul foglio Dettaglio fatture
        var fatture = objDiffideDaImportare[i]['Fatture']
        Logger.log('fatture scritte nel foglio dettaglio fatture ' + fatture) // il Codice Cliente è ancora correttamente formattato come testo
        Logger.log(lastColDettaglioFatture)
        sheetDettaglioFatture.getRange(rowFatture,1,fatture.length,lastColDettaglioFatture).setValues(fatture)
          //sheetDettaglioFatture.appendRow(fatture[r])
        rowFatture = rowFatture + fatture.length
}
  
  Logger.log(objDiffideDaImportare)
  var results = [objDiffideDaImportare, dataImportazione,arrayObjErrors]
  return results
}


function updateFileState(url, dataImportazioneAffido){
  Logger.log('updateFileState')
  var fileState = 'ERR'
  var sheet = sheetFilesAffidi
  var lastRow = sheet.getLastRow()
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
      Logger.log(i + ' ' + data[i][4])
      if (data[i][2] == url) { 
        url = sheet.getRange(i+1, 3).getValue()
        sheet.getRange(i+1, 4).setValue('Importato')
        fileState=sheet.getRange(i+1, 4).getValue()
        //var dataImportazione = 
        sheet.getRange(i+1, 6).setValue(dataImportazioneAffido);
        sheet.getRange(i+1, 7).setValue('');
      }
  }
var fileName = SpreadsheetApp.openByUrl(url).getBlob().getName()
var fileNameUpdated = fileName + ' importato il ' + dataImportazioneAffido 
SpreadsheetApp.openByUrl(url).getBlob().setName(fileNameUpdated)
}

