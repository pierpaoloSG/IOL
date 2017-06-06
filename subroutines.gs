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
 var headersFatture = sheetDettaglioFatture.getDataRange().getValues()[0]
 
 //crea array di oggetti  objDiffideDaInviare
 
 // inizia con Casi NO Fatture
 var objDiffideDaImportare = []
 var tipoFlusso 
 var progressivoImportazioni = 0
 var numErr = 0
 var importErr, amountErr
 var objErrors=[]
 
 function importError(numErr, typeError, idDiffida, codiceCliente, details) {
    this.numErr = numErr; 
    this.typeError = typeError;
    this.idDiffida = idDiffida;
    this.codiceCliente = codiceCliente;
    this.details = details
}
 
 function amountError(typeError, description, totalAmount, totalInvoicesAmount){
           this.typeError = typeError;
           this.description = description;
           this.totalAmount = totalAmount;
           this.totalInvoicesAmount = totalInvoicesAmount;
           }
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
 var logRateiFattura = []

 //*******************************************    
 // ciclo i (objCasiNoFatture)
 for (var i=0; i<objCasiNoFatture.length; i++){
      tipoFlusso = '***'
   Logger.log('i = ' + i)
     //incrementa ID e riferimento pratica 
     newIdDiffida = Utilities.formatDate(new Date(), 'CET', 'YYYYMMddHHmmss')
     //crea oggetto relativo a pratica protocollata
     var importoScoperto = objCasiNoFatture[i].importoScoperto

  
     
    // gestisce le pratiche EX FLUSSO
    
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
         'Codice cliente': String(objCasiNoFatture[i].codcliente),
         'Dato fiscale': String(objCasiNoFatture[i].datoFiscale), // verrà controllato per individuare se CF o PIVA
         'Ragione sociale': String(objCasiNoFatture[i].ragioneSociale),
         'Indirizzo': objCasiNoFatture[i].indirizzoResidenza,
         'CAP': String(objCasiNoFatture[i].capResidenza),
         'Comune' : objCasiNoFatture[i].comuneResidenza,
         'Provincia': objCasiNoFatture[i].provinciaResidenza,
         'Telefono': String(objCasiNoFatture[i].telefono),
         'Provenienza indirizzo': 'CACS',
         'Importo totale': objCasiNoFatture[i].importoScoperto, // l'importo totale è quello del file CasiNoFatture !!
         'Data importazione': dataImportazione,
         'Stato': 'Importata'
       
     }
     
      // gestisce il DATO FISCALE
     objDiffideDaImportare[i]['Codice fiscale'] = String(ControllaCF(objDiffideDaImportare[i]['Dato fiscale']))
     objDiffideDaImportare[i]['Partita IVA'] =  String(ControllaPIVA(objDiffideDaImportare[i]['Dato fiscale']))
     
     Logger.log(objDiffideDaImportare[i]['Codice fiscale'])
     Logger.log(objDiffideDaImportare[i]['Partita IVA'])
     


     // cerca info camerali di Casi No Fatture    
     // ATTENZIONE il match è effettuato  tra il dato fiscale di CASI NO FATTURE e partita IVA o Codice FiscaLE di INFO VISURE CAMERALI
     // in quanto il codice cliente su Visure Camerali non corrisponde
     
     for (var j in objVisureCamerali){
       // se il dato fiscale ha una corrispondenza nel foglio Visure Camerali allora vengono sostituiti i dati con questi ultimi
           if (objDiffideDaImportare[i]['Dato fiscale'] === objVisureCamerali[j].piva || objDiffideDaImportare[i]['Dato fiscale']  === objVisureCamerali[j].codiceFiscale){
                 objDiffideDaImportare[i]['Partita IVA'] = String(objVisureCamerali[j].piva)
                 objDiffideDaImportare[i]['Codice fiscale'] = String(objVisureCamerali[j].codiceFiscale)
                 objDiffideDaImportare[i]['Indirizzo'] = objVisureCamerali[j].indirizzo
                 objDiffideDaImportare[i]['CAP'] = objVisureCamerali[j].cap
                 objDiffideDaImportare[i]['Comune'] = objVisureCamerali[j].comune
                 objDiffideDaImportare[i]['Provincia'] = objVisureCamerali[j].provincia
                 objDiffideDaImportare[i]['Telefono'] = String(objCasiNoFatture[i].telefono)
                 objDiffideDaImportare[i]['Provenienza indirizzo'] = 'Info camerali'  
                 //objDiffideDaImportare[i]['Data importazione'] = dataImportazione
                 //objDiffideDaImportare[i]['Stato'] = 'Importata'  
           } 
     }
     
     
     /* da testare !!
        // recupera la regione
        objDiffideDaImportare[i]['Regione']  = retrieveRegione( objDiffideDaImportare[i]['Provincia'])
     */
     
    // inizializza variabili relative a fatture
    var progressivoFattura = 0
    var fatturaInRatei = false
    var totaleScadutoFatture, 
        totaleAScadereFatture,
        totaleScopertoFatture,
        totaleImportiFatture
    
    //crea array interno per la proprietà fatture
    // filtra i objCasiFatture in base a codice cliente
    Logger.log('length objCasiFatture ' + objCasiFatture.length)
    var fatture = []
 
  
   var  fatturePerCodiceCliente = objCasiFatture.filter(function (el) {
    return el.codcliente == objCasiNoFatture[i].codcliente
   });
   
      
   if (fatturePerCodiceCliente.length > 0){
       rifPraticaFlusso = newRiferimentoPratica + "/" + tipoFlusso
       var numeriFatture = []  
       for (var f = 0; f<fatturePerCodiceCliente.length; f++){
           numeriFatture.push(fatturePerCodiceCliente[f].numeroFattura)
        }
   
        function onlyUnique(value, index, self) { 
          return self.indexOf(value) === index;
        }
 
       var numeriFattureDistinte = numeriFatture.filter(onlyUnique) 
   	
  
      for (var nf=0; nf<numeriFattureDistinte.length; nf++){
        var numeroFattura = numeriFattureDistinte[nf]
        var fatturaUnicaPerCodiceCliente = fatturePerCodiceCliente.filter(function (el) {
          return el.numeroFattura == numeroFattura
        });
        
        Logger.log(fatturaUnicaPerCodiceCliente) 
        
        totaleScadutoFatture = 0
        totaleAScadereFatture = 0
        totaleScopertoFatture = 0
        totaleImportiFatture = 0
        
        for (var rf=0; rf<fatturaUnicaPerCodiceCliente.length; rf++){
            totaleScadutoFatture += fatturaUnicaPerCodiceCliente[rf].importoScaduto
            totaleAScadereFatture += fatturaUnicaPerCodiceCliente[rf].importoAScadere
            totaleScopertoFatture += fatturaUnicaPerCodiceCliente[rf].importoScoperto
            
        var dataImportazioneFattura = dataImportazione
            if (fatturaUnicaPerCodiceCliente[rf].dataFattura){
                dataFattura = fatturaUnicaPerCodiceCliente[rf].dataFattura
                if (typeof(dataFattura) == 'string'){
                  dateTimeFattura = convertStringToDate(dataFattura).getTime() // non attribuisce un valore cronologicamente congruente alla data !!!
                }
                else
                {
                  if (isValidDate(dataFattura)){ 
                    dateTimeFattura = dataFattura.getTime() 
                  }  
                }
              }
            else
            {
              dataFattura = ''
              dateTimeFattura = progressivoFattura // inserisce un datatime fittizio per ordinare le fatture (poi verrà eliminato)
            } 
            
          }
        progressivoFattura++
        
        //var fatturaBuffer = [new Date(dataFattura).getTime(), newIdDiffida, rifPraticaFlusso, String(objCasiNoFatture[i].codcliente),numeroFattura,dataFattura,totaleImportiFatture , dataImportazione]
        fatture.push([dateTimeFattura, newIdDiffida, rifPraticaFlusso, String(objCasiNoFatture[i].codcliente),numeroFattura,dataFattura, totaleScopertoFatture , dataImportazione]) 
      }
  }  
  
  Logger.log(fatture)

 
     Logger.log('fatture con shift e sort by date ' + JSON.stringify(fatture))  
     // finalizza l'oggetto diffida
     objDiffideDaImportare[i]['Fatture presenti'] = progressivoFattura
     //objDiffideDaImportare[i]['Importo totale'] = importoTotale
     objDiffideDaImportare[i]['Spese legali'] = speseLegali
     objDiffideDaImportare[i]['Totale scoperto fatture'] = totaleScopertoFatture
     // verifica se l'importo totale di CasiNoFatture è difforme dal totale degli importi delle fatture da CasiFatture                       
          if (objDiffideDaImportare[i]['Importo totale'] != objDiffideDaImportare[i]['Importo totale fatture']){
             numErr++;
             amountErr = new amountError ('Importi','Difformità tra l\'importo totale della pratiaca e la somma degli importi delle fatture', objDiffideDaImportare[i]['Importo totale'], objDiffideDaImportare[i]['Totale scoperto fatture'])
             importErr = new importError (numErr, 'in fase di importazione',  objDiffideDaImportare[i]['ID diffida'], objDiffideDaImportare[i]['Codice cliente'], amountErr);
             
             objErrors.push(importErr)
          }   
          
        // ordina fatture per data di emissione
      fatture.sort(function(a,b) {
        return a[0]-b[0]
      });
      


     Logger.log('fatture ordinate per data ' + JSON.stringify(fatture))
     // elimina il primo elemento da tutti gli array
     fatture.map(function(val){
       return val.shift(0);
     });

      
     objDiffideDaImportare[i]['Fatture'] = fatture

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
        Logger.log('fatture scritte nel foglio dettaglio fatture ' + fatture) 
        Logger.log(lastColDettaglioFatture)
        sheetDettaglioFatture.getRange(rowFatture,1,fatture.length,lastColDettaglioFatture).setValues(fatture)
          //sheetDettaglioFatture.appendRow(fatture[r])
        rowFatture = rowFatture + fatture.length
}
  
  Logger.log(objDiffideDaImportare)
  Logger.log(objErrors)
  var results = [objDiffideDaImportare, dataImportazione,objErrors,logRateiFattura]
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

