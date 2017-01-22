
function exportTestataPratiche(objDiffideInviate){
  // esporta le diffide da inviare su testate pratiche per Fabrizio
  Logger.log('exportTestataPratiche')
  
  var testataOuter = []
  var addressOuter = []
  var lettereOuter = []
  var lettere = []
  var testata, address, lettere
  var rows

  for (var i in objDiffideInviate)  {
    testata = [,,,,,,,,,,,,,]
    testata[0] = 0 // IDAccount
    testata[1] = '' // Titolo
    testata[2] = '' // Nome
    testata[3] = objDiffideInviate[i]['Ragione sociale']  // Cognome
    testata[4] = objDiffideInviate[i]['Indirizzo'] // Indirizzo (comprende il civico)
    testata[5] = ''  // manca
    testata[6] = objDiffideInviate[i]['Comune'] // Comune
    testata[7] = objDiffideInviate[i]['CAP'] // Cap
    testata[8] = objDiffideInviate[i]['Provincia'] 
    testata[9] = objDiffideInviate[i]['Nazione'] = 'Italia' // manca in tabella
    testata[10] = objDiffideInviate[i]['Riferimento pratica'] + '/' + objDiffideInviate[i]['Tipologia flusso']
    testata[11] = objDiffideInviate[i]['Data importazione'] // affido del (viene esportato su google doc in formato data)
    testata[12] = objDiffideInviate[i]['Codice cliente'] // CodiceCliente
    testata[13] = objDiffideInviate[i]['Dato fiscale'] // CodiceFiscale
    testata[14] = '' // Partita IVA
    testata[15] = 9 // IDAccountStatus
    testata[16] = '' // StatoCamerale
    testata[17] = objDiffideInviate[i]['Provenienza indirizzo'] // ProvenienzaIndirizzo
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
    
    
    
    
    address  = [,,,,,,,]
    address[0] = objDiffideInviate[i]['Codice cliente'] // CodiceCliente
    address[1] = '' // IDRecapito
    address[2] = objDiffideInviate[i]['Telefono'] // Numero Telefonico
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
    
    
    
    
    lettere  = [,,,,]
    lettere[0] = 'in uscita' // Tipologia
    lettere[1] = '#' + objDiffideInviate[i]['Data invio'] + '#' // CodiceFiscale // Data
    lettere[2] = 'Inviata' // Status
    lettere[3] = objDiffideInviate[i]['Codice cliente'] // CodiceCliente

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
    sheetExportLettere.getRange(2,1,rows,headers.length).clearContent()
    sheetExportLettere.getRange(2,1,lettereOuter.length,headers.length).setValues(lettereOuter)

  }
  
  return 
  
}

function exportDettagliFatture(objDiffideInviate){
  // esporta i dettagli fatture per Fabrizio
  Logger.log('exportDettagliFatture')
  
  var rows
  
  var data = sheetDettaglioFatture.getDataRange().getValues()
  var fattureSheet = ObjApp.rangeToObjectsNoCamel(data)
  
  var fattureOuter, fatturaInner, prop 
  var fattureOuter=[]
  for (var i=0; i<objDiffideInviate.length; i++){
     for (j=0; j<fattureSheet.length; j++){
         if (objDiffideInviate[i]['ID diffida'] == fattureSheet[j]['idDiffida']){
            fatturaInner = [,,,,,,,]     
            fatturaInner[0] = '' // IDFattura
            fatturaInner[1] = '' // Descrizione
            fatturaInner[2] = fattureSheet[j]['Importo scoperto']
            fatturaInner[3] = fattureSheet[j]['Data emissione']  // dataOraFattura
            fatturaInner[4] = fattureSheet[j]['Numero fattura'] // NumeroFattura
            fatturaInner[5] = 0 // IDAccount
            fatturaInner[6] = fattureSheet[j]['Codice cliente']// CodiceCliente     
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
  
  sheetExportDettagliFatture.getRange(2,1,fattureOuter.length,headers.length).setValues(fattureOuter)
  return
  
}
  

