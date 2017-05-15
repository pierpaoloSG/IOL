function convertFiles(filesType) {
    Logger.log('convertFiles')
    
    switch (filesType) {
    
      case 'Report': 
           var sourceFolder = reportOriginaliFolder
           var sheet = sheetFilesReport
           break;
      case 'Affidi':
            var sourceFolder = affidiOriginaliFolder
            var sheet = sheetFilesAffidi
            break;
      default :
    
    }
    
    var files = sourceFolder.getFiles()
    var file, data
    var alreadyConverted = false
    var onFolder = 0
    var lastRow = sheet.getLastRow()
    var wroteOnSheet = lastRow -1 
    var row
    var numberOfFilesInFolder = countFilesInFolder(sourceFolder)
    //itera lungo i file trovati sulla folder
    while (files.hasNext()) {
            file = files.next();
            onFolder++
            alreadyOnSheet = false
             // imposta il valore di una nuova riga dello sheet
            //con i dati del file sulla folder
            var name = file.getName()

            data = [ 
              name,
              file.getDateCreated(),
              file.getUrl(),
              "Assegnato", //Stato
              new Date(), //Data assegnazione 
            ];
            
            Logger.log('DATA ************************************************')
            Logger.log(data)
                  
                  //verifica se il file presente sulla folder è già presente sullo sheet
                 // se non ci sono file già scritti sullo sheet salta al prossimo file sulla folder
              Logger.log('AlreadyOnSheet ' + alreadyOnSheet)
              Logger.log("wroteOnSheet " + wroteOnSheet)
              for (row=2; row<=wroteOnSheet+1; row++){  
              Logger.log('row ' + row)
                        // se il file della folder è gia sullo sheet pulisci colonna Badge
              Logger.log(data[3])
              Logger.log(sheet.getRange(row,4).getValue())
                        if (data[3] == sheet.getRange(row,4).getValue()){
                            // memorizza che il file era già sullo sheet
                            alreadyOnSheet = true
                            // pulisce eventuale 'new' su Badge
                            sheet.getRange(row,8).setValue('')
                        }
        
                  }
              if(!alreadyOnSheet && (numberOfFilesInFolder > wroteOnSheet)){
                  // scrive il file sullo sheet
                  data[6]= "new" // Badge
                  Logger.log('data ' + data)
                  sheet.appendRow(data);
                  wroteOnSheet++ 
              }          
     }
     var objFilesImported = ObjApp.rangeToObjectsNoCamel(sheet.getDataRange().getValues())
     Logger.log(JSON.stringify(objFilesImported))
     stop
     return JSON.stringify(objFilesImported)
     
}

