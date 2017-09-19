
function exportDiffide(objDiffide, typePratiche){
Logger.log('******************************************objDiffide')
Logger.log(JSON.stringify(objDiffide))
Logger.log(objDiffide.length)


Logger.log(typePratiche)

switch (typePratiche){

  case 'IOL':
      // export IOL
      var IDFolderExport = '0BwA0zaKu6qcrUE1DVS1mbFVjeXc'
      var urlExportTestataPratiche = 'https://docs.google.com/spreadsheets/d/1lM22uPC41C32QdYkYZS7riAfCvPrrV2NXJj6y1iuwg4/edit'
      var ssExportTestataPratiche = SpreadsheetApp.openByUrl(urlExportTestataPratiche)
      var sheetExportTestataPratiche = ssExportTestataPratiche.getSheetByName('Export Testata Pratiche')
      
      var urlExportRecapiti = 'https://docs.google.com/spreadsheets/d/1PYMar4kIRYyl04m5AT6uk-Ucm4MWd8hlKG67VyaKno0/edit'
      var ssExportRecapiti = SpreadsheetApp.openByUrl(urlExportRecapiti)
      var sheetExportRecapiti = ssExportRecapiti.getSheetByName('Export Recapiti')
      
      var urlExportLettere = 'https://docs.google.com/spreadsheets/d/18yH2BqBjLMFjZmHsGwuGZj06hRRF16eGWTBaNyuFKEE/edit'
      var ssExportLettere = SpreadsheetApp.openByUrl(urlExportLettere)
      var sheetExportLettere = ssExportLettere.getSheetByName('Export Lettere')
      
      var urlExportDettagliFatture = 'https://docs.google.com/spreadsheets/d/1kzOjyZrnPVuGUkAHLEi7OjETVKOv9bOI3qZSPltDIGQ/edit'
      var ssExportDettagliFatture = SpreadsheetApp.openByUrl(urlExportDettagliFatture)
      var sheetExportDettagliFatture = ssExportDettagliFatture.getSheetByName('Export Dettagli Fatture')
  break
  
  case 'MCR':
  
    var IDFolderExport = '0BwA0zaKu6qcrMXJ5WUdGMWdkYzQ'
    var urlExportTestataPratiche = 'https://docs.google.com/spreadsheets/d/1Rw510g5XyTOSsOSh2Z2lp_rWiGJsKKpmpkhe9T-A4dc/edit'
    var ssExportTestataPratiche = SpreadsheetApp.openByUrl(urlExportTestataPratiche)
    var sheetExportTestataPratiche = ssExportTestataPratiche.getSheetByName('Export Testata Pratiche')
    
    var urlExportRecapiti = 'https://docs.google.com/spreadsheets/d/1mee1e2KxKNrllgmUZT63zCP7dF2uGUhZvn8_80F-y_s/edit'
    var ssExportRecapiti = SpreadsheetApp.openByUrl(urlExportRecapiti)
    var sheetExportRecapiti = ssExportRecapiti.getSheetByName('Export Recapiti')
    
    var urlExportLettere = 'https://docs.google.com/spreadsheets/d/1AlhYztDsNmmooTddIBQPGe1SUVxKOR_olWoEfhU4utM/edit'
    var ssExportLettere = SpreadsheetApp.openByUrl(urlExportLettere)
    var sheetExportLettere = ssExportLettere.getSheetByName('Export Lettere')
    
    var urlExportDettagliFatture = 'https://docs.google.com/spreadsheets/d/1g5UmmxxJFy18Y1YM50AhDbMpsJ1kBgIYgUz9weYtG7w/edit'
    var ssExportDettagliFatture = SpreadsheetApp.openByUrl(urlExportDettagliFatture)
    var sheetExportDettagliFatture = ssExportDettagliFatture.getSheetByName('Export Dettagli Fatture')
  break;
  default:
}

//Crea le tabelle da importare nel gestionale di Fabrizio

Logger.log('Export Diffide')
Logger.log(objDiffide)
Logger.log(objDiffide.length)


//var dataImportazione = Utilities.formatDate(new Date(), 'CET', 'dd/MM/YYYY HH.mm')
var dataExport = Utilities.formatDate(new Date(), 'CET', 'dd/MM/YYYY HH.mm.ss')

var arrayDiffideDaInviare = objDiffide['diffide']
//Logger.log('arrayDiffideDaInviare ' + arrayDiffideDaInviare)


var lastRow = sheetDiffideDaInviare.getLastRow()
var lastCol = sheetDiffideDaInviare.getLastColumn()
var range = sheetDiffideDaInviare.getRange(2,1,lastRow, lastCol)

var data = sheetDiffideDaInviare.getDataRange().getValues()
var objDiffideDaInviare = ObjApp.rangeToObjectsNoCamel(data)
//Logger.log(objDiffideDaInviare)
var objDataRow 
var count = 0
var objData = []

 var  objDiffideDaInviare = objDiffideDaInviare.filter(function (el) {
    return el['Tipologia flusso'] == typePratiche
  });

Logger.log(arrayDiffideDaInviare)
for (var i=0; i<arrayDiffideDaInviare.length; i++)  {
    for (var record in objDiffideDaInviare){
      if ((objDiffideDaInviare[record]['ID diffida'] == arrayDiffideDaInviare[i]) && (objDiffideDaInviare[record]['Stato'] == 'Inviata' )){
            // crea l'oggetto per la diffida trovata
            objDataRow = objDiffideDaInviare[record]
            //objDataRow['Data importazione'] = dataImportazione   
            objDataRow['Data esportazione'] = dataExport                                                  
            objDataRow['Stato'] = 'Esportata'
            count++
            //Logger.log(count)
            //Logger.log(objDataRow)
            objData.push(objDataRow)
      }
  }

}

Logger.log('numero diffide da esportare ' + objData.length)
Logger.log('diffide da esportare ')
Logger.log('objData')
Logger.log(objData)
Logger.log(objData.length)

  // esporta le diffide da inviare su testate pratiche per Fabrizio
  Logger.log('exportTestataPratiche')
 var objDiffideToExport = objData
  Logger.log(objDiffideToExport)
  var testataOuter = []
  var addressOuter = []
  var lettereOuter = []
  var lettere = []
  var testata, address, lettere
  var rows

  for (var i in objDiffideToExport)  {

    testata = []
    testata[0] = 0 // IDAccount
    testata[1] = '' // Titolo
    testata[2] = '' // Nome
    testata[3] = objDiffideToExport[i]['Ragione sociale']  // Cognome
    testata[4] = objDiffideToExport[i]['Indirizzo'] // Indirizzo (comprende il civico)
    testata[5] = ''  // manca
    testata[6] = objDiffideToExport[i]['Comune'] // Comune
    testata[7] = objDiffideToExport[i]['CAP'] // Cap
    testata[8] = objDiffideToExport[i]['Provincia'] 
    testata[9] = objDiffideToExport[i]['Nazione'] = 'Italia' // manca in tabella
    testata[10] = objDiffideToExport[i]['Riferimento pratica'] + '/' + objDiffideToExport[i]['Tipologia flusso']
    testata[11] = Utilities.formatDate(objDiffideToExport[i]['Data importazione'], 'CET', '#yyyy-MM-dd#') // affido del (viene esportato su google doc in formato data)
    testata[12] = objDiffideToExport[i]['Codice cliente'] // CodiceCliente
    testata[13] = objDiffideToExport[i]['Codice fiscale'] // CodiceFiscale
    testata[14] = objDiffideToExport[i]['Partita IVA'] // CodiceFiscale
    testata[15] = 9 // IDAccountStatus
    testata[16] = objDiffideToExport[i]['Stato camerale'] // StatoCamerale
    testata[17] = objDiffideToExport[i]['Provenienza indirizzo'] // ProvenienzaIndirizzo
    testataOuter.push(testata)    

 
    var data = sheetExportTestataPratiche.getDataRange().getValues();
    Logger.log('data ' + data.length)
    var headersTestata = data[0]
    if (data.length==1){
      rows = 1
    }
    else
    {
      rows = data.length-1
    }
    Logger.log('testata '  + i + ' ' + rows)
    
    //cancella le precedenti esportazioni
    //sheetExportTestataPratiche.getRange(2,1,rows,headers.length).clearContent()
    //sheetExportTestataPratiche.getRange(2,1,testataOuter.length,headers.length).setValues(testataOuter)
    
    // accoda alle precedenti
    //sheetExportTestataPratiche.getRange(rows+2,1,testataOuter.length,headers.length).setValues(testataOuter)
    

    
    //var urlMySQL = 'jdbc:mysql://<host-ip>:3306/<database>';
    //var urlMySQL = 'jdbc:mysql://hostingmysql330.register.it:3306/scenariopubblico_com_daniele' ; 
    //var db = ObjDB.open( urlMySQL, "SFN1_danielez","Akluhdsab6874");
    //ObjDB.insertRow( db, 'testata pratiche', testataOuter )
    // export Recapiti

    address  = [,,,,,,,]
    address[0] = objDiffideToExport[i]['Codice cliente'] // CodiceCliente
    address[1] = '' // IDRecapito
    address[2] = objDiffideToExport[i]['Telefono'] // Numero Telefonico
    address[3] = '' // Descrizione
    address[4] = 'Ufficio' // Tipologia
    address[5] = -1 // Principale
    address[6] = '' // IDRiferimento
    
    addressOuter.push(address)
     
    var data = sheetExportRecapiti.getDataRange().getValues();
    var headersRecapiti = data[0]
    if (data.length==1){
      rows = 1
    }
    else
    {
      rows= data.length-1
    }
    Logger.log('address '  + i + ' ' + rows)
    //cancella le precedenti esportazioni
    //sheetExportRecapiti.getRange(2,1,rows,headers.length).clearContent()
    //sheetExportRecapiti.getRange(2,1,testataOuter.length,headers.length).setValues(addressOuter)
    
    // accoda alle precedenti
    //sheetExportRecapiti.getRange(rows+2,1,testataOuter.length,headers.length).setValues(addressOuter)
    
    //ObjDB.insertRow( db, 'recapiti', addressOuter )
    
    // export Lettere
    
    var dataInvio = Utilities.formatDate(new Date(objDiffideToExport[i]['Data invio']), 'CET', '#yyyy-MM-dd#')
    Logger.log(dataInvio)
    
    lettere  = [,,,,]
    lettere[0] = 'in uscita' // Tipologia
    lettere[1] = dataInvio // Data //
    lettere[2] = 'Inviata' // Status
    lettere[3] = objDiffideToExport[i]['Codice cliente'] // CodiceCliente

    lettereOuter.push(lettere)
    
    var data = sheetExportLettere.getDataRange().getValues();
    var headersLettere = data[0]
     if (data.length==1){
      rows = 1
    }
    else
    {
      rows= data.length-1
    }
         Logger.log('lettere ' + i + ' ' + rows)

  //cancella le precedenti esportazioni
  


  //sheetExportLettere.getRange(2,1,rows,headers.length).clearContent()
  //sheetExportLettere.getRange(2,1,lettereOuter.length,headers.length).setValues(lettereOuter)    
  
  // accoda alle precedenti
  //sheetExportLettere.getRange(rows+2,1,lettereOuter.length,headers.length).setValues(lettereOuter)     
  }
  
  //cancella le precedenti esportazioni
  
      sheetExportTestataPratiche.getRange(2,1,rows,headersTestata.length).clearContent()
      sheetExportTestataPratiche.getRange(2,1,testataOuter.length,headersTestata.length).setValues(testataOuter)
      
      sheetExportRecapiti.getRange(2,1,rows,headersRecapiti.length).clearContent()
      sheetExportRecapiti.getRange(2,1,testataOuter.length,headersRecapiti.length).setValues(addressOuter)
        
      sheetExportLettere.getRange(2,1,rows,headersLettere.length).clearContent()
      sheetExportLettere.getRange(2,1,lettereOuter.length,headersLettere.length).setValues(lettereOuter) 
  
  // accoda alle precedenti

//    sheetExportTestataPratiche.getRange(rows+2,1,testataOuter.length,headersTestata.length).setValues(testataOuter)
//    sheetExportRecapiti.getRange(rows+2,1,testataOuter.length,headersRecapiti.length).setValues(addressOuter)
//    sheetExportLettere.getRange(rows+2,1,lettereOuter.length,headersLettere.length).setValues(lettereOuter)
    
  //ObjDB.insertRow( db, 'lettere', lettereOuter )
  removeEmptyRows(sheetExportTestataPratiche)
  removeEmptyRows(sheetExportLettere)
  removeEmptyRows(sheetExportRecapiti)
 
  // esporta i dettagli fatture per Fabrizio
  Logger.log('exportDettagliFatture')
  
  var rows
  
  var data = sheetDettaglioFatture.getDataRange().getValues()
  var fattureSheet = ObjApp.rangeToObjectsNoCamel(data)
  
  var fattureOuter, fatturaInner, prop 
  var fattureOuter=[]
  for (var i=0; i<objDiffideToExport.length; i++){
      var  fattureSheetFiltered = fattureSheet.filter(function (el) {
        return el['Data fattura'] != '' 
      });
     for (j=0; j<fattureSheetFiltered.length; j++){
         if (objDiffideToExport[i]['ID diffida'] == fattureSheetFiltered[j]['idDiffida']){
             var importo = fattureSheetFiltered[j]['Importo scoperto']
             Logger.log('importo: ' + importo)
             Logger.log(typeof(importo))
             
             if (fattureSheetFiltered[j]['Data emissione']!=''){
               var dataFattura = Utilities.formatDate(new Date(fattureSheetFiltered[j]['Data emissione']), 'CET', '#yyyy-MM-dd#') 
             }
             else
             {
                var dataFattura = ''
             }
                        
             fatturaInner = [,,,,,,,]     
             fatturaInner[0] = '' // IDFattura
             fatturaInner[1] = '' // Descrizione
             fatturaInner[2] = fattureSheetFiltered[j]['Totale importi fatture']
             fatturaInner[3] = dataFattura // dataFattura
             fatturaInner[4] = fattureSheetFiltered[j]['Numero fattura'] // NumeroFattura
             fatturaInner[5] = 0 // IDAccount
             fatturaInner[6] = fattureSheetFiltered[j]['Codice cliente'] // CodiceCliente     
             fattureOuter.push(fatturaInner)
             Logger.log('fattureOuter')
             Logger.log(fattureOuter)
             //sheetExportTestataPratiche.appendRow(testata)
           }
      }
   }

  var data = sheetExportDettagliFatture.getDataRange().getValues();
 
  var headers = data[0]
   if (data.length==1){
      rows = 1
    }
    else
    {
      rows= data.length-1
    }  
  //cancella le precedenti esportazioni
  sheetExportDettagliFatture.getRange(2,1,rows,headers.length).clearContent()
  sheetExportDettagliFatture.getRange(2,1,fattureOuter.length,headers.length).setValues(fattureOuter)
  
  // accoda alle precedenti
  //sheetExportDettagliFatture.getRange(rows+2,1,fattureOuter.length,headers.length).setValues(fattureOuter)
 
  Logger.log('fattureOuter')
  Logger.log(fattureOuter)
  
    
  removeEmptyRows(sheetExportDettagliFatture)
  updateSheets(objDiffideToExport)
   
  //ObjDB.insertRow( db, 'dettagli fatture', fattureOuter )
  var results = [objData.length, urlExportTestataPratiche,urlExportRecapiti,urlExportLettere,urlExportDettagliFatture]
  return results
}
 

