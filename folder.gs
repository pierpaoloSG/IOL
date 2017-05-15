function countFilesInFolder(folderID) {
 // Attenzione modificare la folder in caso di necessità
  // cerca gli spreadsheet presenti nella folder indicata
 Logger.log(folderID)
 var query = "trashed = false  and mimeType = 'application/vnd.google-apps.spreadsheet' and '" + folderID +"' in parents"
 Logger.log(query)

  var filesInFolder = Drive.Files.list({q: query});

  Logger.log('filesInFolder: ' + filesInFolder.items.length);
  var numberOfFiles = filesInFolder.items.length
  for (var i=0;i<filesInFolder.items.length;i++) {
    //Logger.log('key: ' + key);
    var thisItem = filesInFolder.items[i];
    Logger.log('value: ' + thisItem.title);
  };
  
  return numberOfFiles
};