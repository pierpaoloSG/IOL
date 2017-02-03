
function exportTestataPratiche(objDiffideInviate){
  // esporta le diffide da inviare su testate pratiche per Fabrizio
  Logger.log('exportTestataPratiche')
  
  
  var  objDiffideToExport = objDiffideInviate.filter(function (el) {
  return el['Tipologia flusso'] != 'IOL' 
  });
  
  Logger.log(objDiffideToExport)
  var testataOuter = []
  var addressOuter = []
  var lettereOuter = []
  var lettere = []
  var testata, address, lettere
  var rows
   
  

  for (var i in objDiffideToExport)  {
  
    testata = [,,,,,,,,,,,,,]
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
    testata[10] = objDiffideToExport[i]['Riferimento pratica'] + '/' + objDiffideInviate[i]['Tipologia flusso']
    testata[11] = objDiffideToExport[i]['Data importazione'] // affido del (viene esportato su google doc in formato data)
    testata[12] = objDiffideToExport[i]['Codice cliente'] // CodiceCliente
    testata[13] = objDiffideToExport[i]['Dato fiscale'] // CodiceFiscale
    testata[14] = '' // Partita IVA
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
    lettere[1] = '#' + objDiffideToExport[i]['Data invio'] + '#' // CodiceFiscale // Data
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
     for (j=0; j<fattureSheet.length; j++){
         if (objDiffideToExport[i]['ID diffida'] == fattureSheet[j]['idDiffida']){
           
            var importo = fattureSheet[j]['Importo scoperto']
            Logger.log(typeof(importo))

            // var importoFormatted = 
            var dataFattura = fattureSheet[j]['Data emissione']
            var dataFatturaFormatted = Utilities.formatDate(dataFattura, 'CET', 'dd-MMM-YY')

            fatturaInner = [,,,,,,,]     
            fatturaInner[0] = '' // IDFattura
            fatturaInner[1] = '' // Descrizione
            fatturaInner[2] = fattureSheet[j]['Importo scoperto']
            fatturaInner[3] = dataFatturaFormatted // dataOraFattura
            fatturaInner[4] = fattureSheet[j]['Numero fattura'] // NumeroFattura
            fatturaInner[5] = 0 // IDAccount
            fatturaInner[6] = fattureSheet[j]['Codice cliente'] // CodiceCliente     
            fattureOuter.push(fatturaInner)    
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
  return 
}
  

