function querySheet(tipoFlusso){
  // filtra il foglio sheetDiffideDaInviare in base al tipoflusso
  //crea un querySheet
  var sheetName = 'Query '+ tipoFlusso
  try {ssDB.setActiveSheet(ssDB.getSheetByName(sheetName))
    var querySheet = ssDB.setActiveSheet(ssDB.getSheetByName(sheetName))
  }
  
  catch (e) {
    var sheetName = 'Query '+ tipoFlusso
    var querySheet = ssDB.insertSheet().setName(sheetName).activate()
  } 
  return querySheet
}

function deleteFogli(){
//cancella tutti i fogli che iniziano con la parola 'Foglio'
  var sheets = ssDB.getSheets()
  Logger.log(JSON.stringify(sheets))
  for (var i=0; i<sheets.length; i++){
        var name = sheets[i].getSheetName()
        Logger.log(name)
        var nameSlice = name.slice(0,6)
        Logger.log(nameSlice)
        if ( nameSlice == 'Foglio'){
           ssDB.deleteSheet(sheets[i]);
        }
  }
}