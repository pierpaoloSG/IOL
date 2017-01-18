// inserisce i dati inerenti ai Files affidi da importare in un oggetto 
function importFilesToDB(sheet){
  
  var objFilesAffidi = ObjApp.rangeToObjectsNoCamel(sheet.getDataRange().getValues())
  
  return objFilesAffidi
}

function importDataToSheet(data,sheet){
 
  // IMPORTA
  
  //elimina la prima e la seconda riga dell'array 2d che era vuota
  data.splice(0, 2); 
  
  //legge l'ultimo protocollo assegnato
  var lastRow = sheet.getLastRow()
  var lastProt = sheet.getRange(lastRow,1).getValue()
  Logger.log(lastProt)
  
  //protocolla i nuovi dati
  for (var i=0; i<data.length; i++){
    data[i].unshift(lastProt+i+1)
  }
  
  //importa i dati nel foglio
  sheet.getRange(lastRow+1,1,data.length, (data[0].length)).setValues(data)
  
}

// legge uno sheet e inserisce i dati in un oggetto
function grabObjectFromSheet(sheet){

  var data = sheet.getDataRange().getValues() 
  var headers = data[0]
  
  //elimina la prima e la seconda riga dell'array 2d che è vuota
  // in quanto negli sheets originali gli header sono composti da 2 righe
  data.splice(0,2);
  data.unshift(headers)
  
  var obj = ObjApp.rangeToObjects(data)
  return obj   
}
  
function createFilesInFolder(docName) {
  //crea il documento
  var doc = DocumentApp.create(docName)
  //apre il documento come file
  var docFile = DriveApp.getFileById( doc.getId() );
  //salva il file nella cartella
  DriveApp.getFolderById(IDFolder).addFile( docFile );
  //rimuove il file originario dalla root
  DriveApp.getRootFolder().removeFile(docFile);
}


function cloneDoc(fileDoc, copyTitle) {;
 var newFileDoc = fileDoc.makeCopy(copyTitle);
 Logger.log(fileDoc.getUrl())
 return newFileDoc
} 

function mergeDataToHtml(newDoc, objData){
  for (var record in objData){
    Logger.log(record)
    var text = 'name ' + objData.name + ' cognome' + objData.cognome
    newDoc.appendPageBreak()
  }
}

/**
* @param  {objData} Oggetto che contiene i dati da sostituire nei rispettivi campi 
 */
function replaceTextWithObject(fileDoc, objData){

  // itera su objData e per ogni oggetto interno richiama la funzione di sostituzione del testo
var i=1
  for (var record in objData){
    Logger.log(record)
    //crea un oggetto con i dati da unire al documento
    var objRecord = objData[record]
    // unisce i dati al documento
    Logger.log(objRecord)
    Logger.log("replace")
    var separator = '%'
     for (var key in objRecord){
       fileDoc.replaceText(separator + key + separator,objRecord[key])
       Logger.log(objRecord[key]) 
       // crea interruzione di pagina 
       fileDoc.appendPageBreak()
     }
     Logger.log(i)
     i++
  }
  
// chiude il ciclo
  
    var nameDoc = 'lotto'
    changeNameDoc(fileDoc, nameDoc)
}

/**
* @param  {doc} Oggetto Document di cui si vuole cambiare il nome
* @name nome da sostituire
 */
function changeNameDoc(doc, name){
 
  doc.setName(name)

}

function getHeight(length, ratio) {
  var height = ((length)/(Math.sqrt((Math.pow(ratio, 2)+1))));
  return Math.round(height);
}

function getWidth(length, ratio) {
  var width = ((length)/(Math.sqrt((1)/(Math.pow(ratio, 2)+1))));
  return Math.round(width);
}

 function comma(num){
    while (/(\d+)(\d{3})/.test(num.toString())){
        num = num.toString().replace(/(\d+)(\d{3})/, '$1'+'.'+'$2');
    }
    return num;
}



// aggiorna lo stato del file a "Affido importato" e scrive la data di importazione



function deleteData(){
  var sheet, lastCol 
  var sheets = [sheetFilesAffidi, sheetDiffideDaInviare,sheetDiffideInviate, sheetDettaglioFatture]

  for (i=0; i<sheets.length; i++){
    sheet = sheets[i]
    lastCol = sheet.getLastColumn()
    var data = sheet.getDataRange().getValues()
    Logger.log(sheet.getName() +" " + data.length)
    if (data.length >1){
      sheet.getRange(2, 1,(data.length-1),lastCol).clear()
    }
  }
}


function checkDuplicate(objDiffideDaInviare, sheet){
    Logger.log('checkDuplicate')
    Logger.log(objDiffideDaInviare)
    var arrayCodCliente = []
    for (var i=0; i<objDiffideDaInviare.length; i++){
        arrayCodCliente.push(objDiffideDaInviare[i]['Codice cliente'])
    }
    var err
    var colCodCliente = 3
    var lastRow = sheet.getLastRow()
    Logger.log(sheet.getName())
    var range = sheet.getRange(2, colCodCliente, lastRow, 1);
    var codClientePresenti = range.getValues();
    Logger.log('arrayCodCliente = ' + arrayCodCliente)
    Logger.log('codClientePresenti ' + codClientePresenti)
    var codClienteDuplicati = []
    for (var x=0; x<arrayCodCliente.length; x++){
        for (var i=0; i < codClientePresenti.length; i++) {
          if (codClientePresenti[i][0] == arrayCodCliente[x]) {
            codClienteDuplicati.push(arrayCodCliente[x])
          }    
        }
    }
   var arrayMessage = [
     "La pratica da importare è relativa ad un codice cliente già presente'" + codClienteDuplicati, 
     "Su " + arrayCodCliente.length + " pratiche da importare, sono stati riscontrati i seguenti " + codClienteDuplicati.length + " codici cliente già presenti: \n" + codClienteDuplicati
     ]
   Logger.log(codClienteDuplicati.length)
   Logger.log (arrayMessage[1])
   if (codClienteDuplicati.length ==1) {
       err = arrayMessage[0]
       throw err
   }else
     if (codClienteDuplicati.length>1){
        err = arrayMessage[1]
        throw err
     }
  else{ 
  return false
  }
}

// remove columns from 2d array
function removeEl(array, remIdx) {
 return array.map(function(arr) {
         return arr.filter(function(el,idx){return idx < remIdx});  // rimuove tutti gli elementi interni con index < di remIdx
        });
};


function getGoogleSpreadsheetAsExcel(url, sheet){
  
  try {
    
    var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + ssDBID + "&exportFormat=xlsx";
    var url = "https://docs.google.com/spreadsheets/d/1vcfSULfxMDQkbQC8WI8m-EKAuD2OvL-s92G6Nv9aG_c/edit#gid=908917144" + "&exportFormat=xlsx";
    
    var params = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };
    
    var blob = UrlFetchApp.fetch(url, params).getBlob();
    
    blob.setName(sheet.getName() + ".xlsx");
    
    MailApp.sendEmail("danielezappala@scenariopubblico.com", "Google Sheet to Excel", "The XLSX file is attached", {attachments: [blob]});
    
  } catch (f) {
    Logger.log(f.toString());
  }
}

function exportAsExcel(spreadsheetId) {
  var file = Drive.Files.get(spreadsheetId);
  var url = file.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  return response.getBlob();
}





