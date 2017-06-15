function stampaLettera(infoLettera, dataInvio, objData) {
  
  var objTipoLettera = objData.filter(function (el) {
    return el['AccountsStatus'] == infoLettera.Status;
  });
  
  var objData = objTipoLettera
  
  Logger.log('numero diffide da inviare ' + objData.length)
  Logger.log('contestazioni da inviare ')
  Logger.log(objData.length)


// apre  il template 
  
  var templateURL = infoLettera.TemplateURL // attenzione ! l'URL deve contenere "/edit" 
  var docTemplate = DocumentApp.openByUrl(templateURL)
  
  // crea nuovo documento nella cartella
  var docName = infoLettera.Titolo + ' ' + Utilities.formatDate(new Date(), "CET", "dd-MM-yyyy' h.'HH:mm:ss")
  
  //apre il documento come file
  var fileTemplate = DriveApp.getFileById( docTemplate.getId())
  var docFile = cloneDoc(fileTemplate, docName)
  var docFileUrl = docFile.getUrl()
  var docFileId = docFile.getId()
  var mergedDoc = DocumentApp.openById(docFileId)
  
  //salva il file nella cartella
  DriveApp.getFolderById(infoLettera.FolderID).addFile( docFile )
  
  //rimuove il file creato nella cartella impostazioni
  DriveApp.getFolderById(IDSettingsFolder).removeFile(docFile)
  
  // pulisce il body
  mergedDoc.getBody().clear()
  
  // copia il body
  var bodyCopy = docTemplate.getActiveSection();

  var body = bodyCopy.copy();
 
  // sostituisce i placeholder nel documento
  body.replaceText("%Data invio%", dataInvio)    
  body.replaceText("%numeroPosizioni%", objData.length)
  
  // definisce l'intestazione della tabella
  var arrayPratiche = [['Rif. Pratica', 'Codice Cliente', 'Ragione Sociale']]

  // prepara il contenuto della tabella la tabella
  for (var record in objData) {
      var arrayPraticheInner = []

      arrayPraticheInner.push(objData[record]['CodicePratica'])
      arrayPraticheInner.push(objData[record]['CodiceCliente'])
      arrayPraticheInner.push(objData[record]['Account'])
      arrayPratiche.push(arrayPraticheInner)
    }
    
  // riempe la tabella con le pratiche
  
    var numChildren = body.getNumChildren();
    Logger.log('numChildren ' + numChildren)

    
    for (var c = 0; c < numChildren; c++) {
      var child = body.getChild(c);
      child = child.copy();
      // cerca l'elemento TABLE
      if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
        mergedDoc.appendHorizontalRule(child);
      } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
        mergedDoc.appendImage(child);
      } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        mergedDoc.appendParagraph(child);
      } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
        mergedDoc.appendListItem(child);
      } else if (child.getType() == DocumentApp.ElementType.TABLE) {
        
        Logger.log(arrayPratiche)
        if (typeof arrayPratiche != 'undefined'){
        // scrive l'header della table
        var table = mergedDoc.appendTable([arrayPratiche[0]]);
        var newRow, cellContent
        // formatta l'header della table
        for (var col=0; col<arrayPratiche[0].length; col++){
            var cell = table.getCell(0,col)
            var para = cell.getChild(0)
              // formatta l'header della table 
            cell.setAttributes(styleHighlight).setWidth(160).setPaddingBottom(0).setPaddingTop(0)
            para.setAttributes(styleHeaderTable)
        }
        // scrive e formatta le righe della table
        for (var row = 1; row<arrayPratiche.length; row++){
          newRow = table.appendTableRow()
          for (var col=0; col<arrayPratiche[row].length; col++){
                   cellContent = arrayPratiche[row][col]           
                   newRow.insertTableCell(col, cellContent)
                   var cell = table.getCell(row,col)
                   var para = cell.getChild(0)
                   cell.setWidth(160).setPaddingBottom(0).setPaddingTop(0)
                   para.setAttributes(styleCell)
              }
          } 
       }else 
         {
            Logger.log("Unknown element type: " + child);
         }
    
    }
   }

   //mergedDoc.getBody().appendParagraph('Distinti Saluti')
   
   // appone la firma 
   var fileImage = urlFirmaPGiacona 
   var fileID = fileImage.match(/[\w\_\-]{25,}/).toString();
   var blob   = DriveApp.getFileById(fileID).getBlob()
   var ratio = (1.130);
   var height = getHeight(400,ratio);
   var width = getWidth(height,ratio);

   Logger.log(height);
   Logger.log(width);
   
   mergedDoc.getBody().appendImage(blob).setWidth(width).setHeight(height).getParent().setAttributes(styleImage)
   
   // inserisce un'interruzione di pagina
   mergedDoc.appendPageBreak()
   
  // nomina il file 
  var d = new Date()
  var timeStamp = d.getTime();  // Number of ms since Jan 1, 1970
  var printDate = Utilities.formatDate(new Date(d), "CET", "dd-MM-yyyy' h.'HH:mm:ss")
  docName = 'Lettera Chiusura Contestazioni ' + printDate
  mergedDoc.setName(docName)
  
  // restituisce il risultato 
  var results = [objData.length, docFileUrl, docName]
    Logger.log('results \n'+ JSON.stringify(results))
  return results
}


