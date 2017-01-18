// legge l'ID dello script
function getThisScriptId() {
  return DriveApp.getFileById('1DaQcz_hZwZjM8syYbIPF7cpT8hj6zhut1hZOlQsXbaWbk_73swYVv8cM');
}

function readDataFilesAffidi(){
  //legge i dati dal foglio Filed Affidi
  var arrayFilesAffidi = sheetFilesAffidi.getDataRange().getValues()
  Logger.log(arrayFilesAffidi)
  return arrayFilesAffidi
}