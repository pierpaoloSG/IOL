//FOLDERS

var affidiDaImportareFolderId = '0BwA0zaKu6qcrNjlpWFN4Y0NoQkU'
var affidiDaImportareFolder = DriveApp.getFolderById(affidiDaImportareFolderId)

var affidiOriginaliFolderId = '0BwA0zaKu6qcrZzlEZjFXVkpOeHM'
var affidiOriginaliFolder = DriveApp.getFolderById(affidiDaImportareFolderId)

// Spreadsheet
var urlDB = 'https://docs.google.com/spreadsheets/d/1GEUt7CDzgXyeXpttijjavcwSTxMD1uQlXdivcN2l8cA/edit'
var ssDB = SpreadsheetApp.openByUrl(urlDB)
var ssDBID = '1GEUt7CDzgXyeXpttijjavcwSTxMD1uQlXdivcN2l8cA'
var sheetImpostazioni = ssDB.getSheetByName('Impostazioni')
var sheetFilesAffidi = ssDB.getSheetByName('Files affidi')
var sheetDiffideDaInviare = ssDB.getSheetByName('Diffide da inviare')
var sheetDiffideInviate = ssDB.getSheetByName('Diffide inviate')
var sheetDettaglioFatture = ssDB.getSheetByName('Dettaglio fatture')
var sheetListaDiControllo = ssDB.getSheetByName('Lista di controllo')


// STAMPA DIFFIDE
IDTemplateDiffide = '1Eu98aL8kc3xlxurS6jBA5mX-2b8XInMwbWUi2sd_BPw'
IDTemplateEtichette = '1OH97tO91HxGMdekZpniGFyYH9K2y1LfKibF19ff5q9U'
IDTemplateListaDiControllo = '1umVg8hFXnWyN01M0n2KZZRO3jwDtf9ZdrTcvpZjKZfU'
IDSettingsFolder = '0BwA0zaKu6qcrcjZGZmgtNmZXVHc'
IDFolder = '0BwA0zaKu6qcrTnNfZk5CdEw1Yk0' // folder diffide stampate
urlFirma = 'https://drive.google.com/open?id=0BwA0zaKu6qcremtGTVpIU2stb2c'





