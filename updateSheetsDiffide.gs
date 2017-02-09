
function updateSheets(objDiffideInviate, docFileUrl, docName){
   // aggiunge le diffide nello sheet diffide inviate
  Logger.log('updateSheetDiffide')
  Logger.log(objDiffideInviate)
  var arrayDiffideDaInviare = sheetDiffideDaInviare.getDataRange().getValues()
  var lastRow = sheetDiffideDaInviare.getLastRow()
  var headersDiffideDaInviare = arrayDiffideDaInviare[0]
  var colDiffideInviate = headersDiffideDaInviare.length
  var colDataStampa = headersDiffideDaInviare.indexOf('Data stampa')+1
  var colDataInvio = headersDiffideDaInviare.indexOf('Data invio')+1
  var colStato = headersDiffideDaInviare.indexOf('Stato')+1
    
  Logger.log(headersDiffideDaInviare)
  Logger.log(colDataStampa)
  Logger.log(colDataInvio)
  Logger.log(colStato)
  //var objDiffideDaInviare = ObjApp.rangeToObjectsNoCamel(range)
  Logger.log('objDiffideDaInviare prima')
  Logger.log(arrayDiffideDaInviare)
  arrayDiffideDaInviare.shift()
  var arrayDiffideInviate=[]
  var toDelete
  var inviate = 0

  for (var i in objDiffideInviate)  {
  Logger.log('i ' + i + ' - ' + objDiffideInviate[i]['ID diffida'])
    for (var j=0; j<arrayDiffideDaInviare.length; j++){
        Logger.log('j ' + j + ' - ' + arrayDiffideDaInviare[j][0])
          if (arrayDiffideDaInviare[j][0] == objDiffideInviate[i]['ID diffida']){
            // elimina la riga per la diffida inviata
            inviate++
            arrayDiffideInviate.push(arrayDiffideDaInviare[j])
            arrayDiffideInviate[inviate-1].pop()
            arrayDiffideInviate[inviate-1].push(objDiffideInviate[i]['Data stampa'])
            arrayDiffideInviate[inviate-1].push(objDiffideInviate[i]['Data invio'])
            arrayDiffideInviate[inviate-1].push(objDiffideInviate[i]['Stato'])
            
            sheetDiffideDaInviare.getRange(j+2,colDataStampa).setValue(objDiffideInviate[i]['Data stampa'])
            sheetDiffideDaInviare.getRange(j+2,colDataInvio).setValue(objDiffideInviate[i]['Data invio'])
            sheetDiffideDaInviare.getRange(j+2,colStato).setValue(objDiffideInviate[i]['Stato'])
         }
    } 
     
  }

Logger.log('arrayDiffideDaInviare')
Logger.log(arrayDiffideDaInviare)
Logger.log('arrayDiffideInviate')
Logger.log(arrayDiffideInviate)

  
  // aggiunge le diffide nello sheet diffide inviate
  for (var i=0; i<arrayDiffideInviate.length; i++){
    sheetDiffideInviate.appendRow(arrayDiffideInviate[i])
  }
  //var results = [arrayDiffideInviate.length, docFileUrl, docName]
  //Logger.log(results)
  //return results
  return
}