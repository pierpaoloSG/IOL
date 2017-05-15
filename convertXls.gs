/**
 * Convert Excel file to Sheets
 * @param {Blob} excelFile The Excel file blob data; Required
 * @param {String} filename File name on uploading drive; Required
 * @param {Array} arrParents Array of folder ids to put converted file in; Optional, will default to Drive root folder
 * @return {Spreadsheet} Converted Google Spreadsheet instance
 **/
 
 // la funzione convertAndProcess
 // Cerca tutti i file xls nella sourceFolder
 // converte ognuno dei file trovati in google sheet e li salva nella destFolder
 // modifica il contenuto del file per renderlo compatibile all'analisi con google data studio
 
 
 
function convertExcel2Sheets(excelFile, filename, arrParents) {
  
  var parents  = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
    Logger.log(parents)
  //if ( !parents.isArray ) parents = []; // make sure parents is an array, reset to empty array if not

  // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
  var uploadParams = {
    method:'post',
    contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
    contentLength: excelFile.getBytes().length,
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: excelFile.getBytes()
  };
  
  // Upload file to Drive root folder and convert to Sheets
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
    
  // Parse upload&convert response data (need this to be able to get id of converted sheet)
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());

  // Create payload (body) data for updating converted file's name and parent folder(s)
  var payloadData = {
    title: filename, 
    parents: []
  };
  if ( parents.length ) { // Add provided parent folder(s) id(s) to payloadData, if any
    for ( var i=0; i<parents.length; i++ ) {
      try {
        var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
        payloadData.parents.push({id: parents[i]});
      }
      catch(e){} // fail silently if no such folder id exists in Drive
    }
  }
  // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
  var updateParams = {
    method:'put',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };
  
  // Update metadata (filename and parent folder(s)) of converted sheet
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);
  
  return SpreadsheetApp.openById(fileDataResponse.id);
}

/**
 * Sample use of convertExcel2Sheets() for testing
 **/
 function convertAndProcess(sourceFolder,destFolders) {
  var xlsFiles = findAllExcelFilesInFolder()
  if (xlsFiles.length == 0){
     Browser.msgBox('Non ci sono file excel da importare')
     stop
  }
  Logger.log(xlsFiles)
  for (var n in xlsFiles){
      //var xlsId = "0BznBNzYR5OHDTnlGT2NxdERrYWs"; // ID of Excel file to convert
      //var xlsFile = DriveApp.getFileById(xlsId); // File instance of Excel file
      var xlsFile = xlsFiles[n]
      var xlsBlob = xlsFile.getBlob(); // Blob source of Excel file for conversion
      var xlsFilename = xlsFile.getName(); // File name to give to converted file; defaults to same as source file
      var ss = convertExcel2Sheets(xlsBlob, xlsFilename, destFolders);
      xlsFile.setTrashed(true)
      prepareSheetToDataStudio(ss)
    }
 }
 
function prepareSheetToDataStudio(ss){

  Logger.log('Prepare ' + ss.getName())
  
  var data = ss.getRange().getValues()
  var objData = ObjApp.rangeToObjectsNoCamel(data)
  var date
  objData.forEach(function(obj) {
        date = obj.DOInserimento
        Logger.log(date)
        obj.ImportoPagamenti = Utilities.formatDate(new Date(date), 'CET', 'yyyy/MM/dd');
    });
  Logger.log(objData)
    stop
}

function findAllExcelFilesInFolder(){
    
    var folder = DriveApp.getFolderById(sourceFolder)
    var xlsFiles = []
    var files = folder.getFiles()
    var file
      while (files.hasNext()){ 
          file = files.next() 
          Logger.log(file.getName())
          Logger.log(file.getMimeType())
          if (file.getMimeType() == 'application/vnd.ms-excel'){
            xlsFiles.push(file)
          }
    }
    
    return xlsFiles
 }
 
 
 function importXLS(){
  //var files = DriveApp.searchFiles('title contains ".xls"');// you can also use a folder as starting point and get the files in that folder... use only DriveApp method here.
  var folder = DriveApp.getFolderById(sourceFolder)
  var files = folder.getFiles()
  while(files.hasNext()){
    var xFile = files.next();
    var name = xFile.getName();
    if (name.indexOf('.xls')>-1){ // this check is not necessaey here because I get the files with a search but I left it in case you get the files differently...
      var ID = xFile.getId();
      var xBlob = xFile.getBlob();
      var newFile = { title : name+'_converted',
                     key : ID
                    }
      file = Drive.Files.insert(newFile, xBlob, {
        convert: true
      });
    }
  }
}