function readFilesAffidiFromFolder(){

  // legge il folder Files Affidi da Importare e scrive il contenuto nel foglio Files Affidi
  writeFilesToSheet(affidiDaImportareFolder, sheetFilesAffidi)
  var objFilesAffidi = ObjApp.rangeToObjectsNoCamel(sheetFilesAffidi.getDataRange().getValues())
  Logger.log(JSON.stringify(objFilesAffidi))
  return JSON.stringify(objFilesAffidi)
  
}
function readFilesReportFromFolder(){

  // legge il folder Files Affidi da Importare e scrive il contenuto nel foglio Files Affidi
  writeFilesToSheet(reportDaImportareFolder, sheetFilesReport)
  var objFilesAffidi = ObjApp.rangeToObjectsNoCamel(sheetFilesAffidi.getDataRange().getValues())
  Logger.log(JSON.stringify(objFilesAffidi))
  return JSON.stringify(objFilesAffidi)
  
}

function readDiffideFromFiles(arrayObjFilesAffidi){
  Logger.log('readDiffideDaInviareFromFiles')
  //legge i dati dei files Affidi da Importare 
  //var objFilesAffidi = ObjApp.rangeToObjectsNoCamel(sheetFilesAffidi.getDataRange().getValues())
  // legge l'ultimo ID sul foglio DiffideDaInviare
  var ssAffido, url 
  var lastRowDiffide = sheetDiffideDaInviare.getLastRow()
  //importa per ogni files i dati degli affidi su DB
  Logger.log('Numero Files di Affido da importare;  ' + arrayObjFilesAffidi.length )
  for (j=0; j<arrayObjFilesAffidi.length; j++){
      var url = arrayObjFilesAffidi[j].url
      ssAffido = SpreadsheetApp.openByUrl(url)
      Logger.log(ssAffido)
      Logger.log(ssAffido.getName())
    var results = readAffido(ssAffido)
    var objDiffideDaImportare = results[0]
    var dataImportazioneAffido = results[1]
    var erroriImportazioneDiffide = results[2]
    var logRateiFattura = results[3]
    Logger.log("errori di importazione \n" + JSON.stringify(erroriImportazioneDiffide))
    var fileState = updateFileState(url,dataImportazioneAffido) 
    }
    Logger.log(objDiffideDaImportare)
    return JSON.stringify(results)  
 } 

 
function readDiffideDaInviareFromSheet(){
    Logger.log('readDiffideDaInviareFromSheet')
    var objAllDiffideDaInviare = ObjApp.rangeToObjectsNoCamel(sheetDiffideDaInviare.getDataRange().getValues())
    Logger.log('objAllDiffideDaInviare')
    Logger.log(objAllDiffideDaInviare)
    Logger.log(JSON.stringify(objAllDiffideDaInviare))
    return JSON.stringify(objAllDiffideDaInviare)
}