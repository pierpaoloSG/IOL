function retrieveAddresses(objDiffide) {
  Logger.log('Retrieve Labels')
  Logger.log(objDiffide)
  var arrayEtichetteDaStampare = objDiffide['diffide']
  var range = sheetDiffideDaInviare.getDataRange().getValues()
  var objDiffideDaInviare = ObjApp.rangeToObjectsNoCamel(range)
  Logger.log(objDiffideDaInviare)
  var objDataRow
  var arrayObjAddresses = []

  
  for (var i=0; i<arrayEtichetteDaStampare.length; i++)  {

      for (var record in objDiffideDaInviare){
        Logger.log(record.length)
        if (objDiffideDaInviare[record]['ID diffida'] ==  arrayEtichetteDaStampare[i]){
          // crea l'oggetto AddressesRecord per la diffida trovata
         var objAddressRecord = {
          'Riferimento pratica': objDiffideDaInviare[record]['Riferimento pratica'],
          'Tipologia flusso': objDiffideDaInviare[record]['Tipologia flusso'] ,
          'Ragione sociale': objDiffideDaInviare[record]['Ragione sociale'] ,
          'Indirizzo': objDiffideDaInviare[record]['Indirizzo'] ,
          'CAP':  objDiffideDaInviare[record]['CAP'] ,
          'Comune': objDiffideDaInviare[record]['Comune'] ,
          'Provincia': objDiffideDaInviare[record]['Provincia']
          }
          Logger.log(objAddressRecord)
          arrayObjAddresses.push(objAddressRecord)
        }
    } 
  }
  Logger.log('numero etichette da stampare ' + arrayObjAddresses.length)
  Logger.log('etichette da stampare ')
  Logger.log(arrayObjAddresses)
  //var results = printLabels(arrayObjAddresses)
  return arrayObjAddresses
}
