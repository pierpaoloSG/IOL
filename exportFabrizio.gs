
function exportTestataPratiche(objDiffideToExport){
  // esporta le diffide da inviare su testate pratiche per Fabrizio
  Logger.log('exportTestataPratiche')
  
  
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
    testata[11] = objDiffideToExport[i]['Data importazione'] // affido del (viene esportato su google doc in formato data)
    testata[12] = objDiffideToExport[i]['Codice cliente'] // CodiceCliente
    testata[13] = objDiffideToExport[i]['Codice fiscale'] // CodiceFiscale
    testata[14] = objDiffideToExport[i]['Partita IVA'] // CodiceFiscale
    testata[15] = 9 // IDAccountStatus
    testata[16] = '' // StatoCamerale
    testata[17] = objDiffideToExport[i]['Provenienza indirizzo'] // ProvenienzaIndirizzo
    testataOuter.push(testata)    
    
    var data = sheetExportTestataPratiche.getDataRange().getValues();
    var headers = data[0]
    if (data.length==1){
      rows = 1
    }
    else
    {
      rows = data.length-1
    }
    sheetExportTestataPratiche.getRange(2,1,rows,headers.length).clearContent()
    sheetExportTestataPratiche.getRange(2,1,testataOuter.length,headers.length).setValues(testataOuter)
    
    
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
    var headers = data[0]
     if (data.length==1){
      rows = 1
    }
    else
    {
      rows= data.length-1
    }
    sheetExportRecapiti.getRange(2,1,rows,headers.length).clearContent()

    sheetExportRecapiti.getRange(2,1,testataOuter.length,headers.length).setValues(addressOuter)
    
    
    
    // export Lettere
    
    lettere  = [,,,,]
    lettere[0] = 'in uscita' // Tipologia
    lettere[1] = '#' + objDiffideToExport[i]['Data invio'] + '#' // Data //
    lettere[2] = 'Inviata' // Status
    lettere[3] = objDiffideToExport[i]['Codice cliente'] // CodiceCliente

    lettereOuter.push(lettere)
    
    var data = sheetExportLettere.getDataRange().getValues();
    var headers = data[0]
     if (data.length==1){
      rows = 1
    }
    else
    {
      rows= data.length-1
    }
    Logger.log(lettereOuter)
    sheetExportLettere.getRange(2,1,rows,headers.length).clearContent()
    sheetExportLettere.getRange(2,1,lettereOuter.length,headers.length).setValues(lettereOuter)    
  }
  
  removeEmptyRows(sheetExportTestataPratiche)
  removeEmptyRows(sheetExportLettere)
  removeEmptyRows(sheetExportRecapiti)
  exportDettagliFatture(objDiffideToExport)
  return 
  
}

function exportDettagliFatture(objDiffideToExport){
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
             Logger.log(typeof(importo))
             
             if (fattureSheetFiltered[j]['Data emissione']!=''){
               var dateFormatted = Utilities.formatDate(new Date(fattureSheetFiltered[j]['Data emissione']), 'CET', 'dd/MM/YYYY') 
               var dataFattura = '#'+dateFormatted+'#'
             }
             else
             {
                var dataFattura = ''
             }
                        
             fatturaInner = [,,,,,,,]     
             fatturaInner[0] = '' // IDFattura
             fatturaInner[1] = '' // Descrizione
             fatturaInner[2] = fattureSheetFiltered[j]['Importo scoperto']
             fatturaInner[3] = dataFattura // dataFattura
             fatturaInner[4] = fattureSheetFiltered[j]['Numero fattura'] // NumeroFattura
             fatturaInner[5] = 0 // IDAccount
             fatturaInner[6] = fattureSheetFiltered[j]['Codice cliente'] // CodiceCliente     
             fattureOuter.push(fatturaInner)
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
  sheetExportDettagliFatture.getRange(2,1,rows,headers.length).clearContent()
  Logger.log(fattureOuter)
  sheetExportDettagliFatture.getRange(2,1,fattureOuter.length,headers.length).setValues(fattureOuter)
  removeEmptyRows(sheetExportDettagliFatture)
  return 
}
  

