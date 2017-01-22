// ***** IOL DEV ******

//FOLDERS

var affidiDaImportareFolderId = '0BznBNzYR5OHDWGFPSUNKWlpkQTA'
var affidiDaImportareFolder = DriveApp.getFolderById(affidiDaImportareFolderId)

var affidiOriginaliFolderId = '0BznBNzYR5OHDVUpKVGYzTy0xeW8'
var affidiOriginaliFolder = DriveApp.getFolderById(affidiDaImportareFolderId)

// Spreadsheet
var urlDB = '//docs.google.com/spreadsheets/d/179OhQjvAQZ9QMYoXptBn2AcxSeWVmFtao3bOZAyXroE/edit'
var ssDB = SpreadsheetApp.openByUrl(urlDB)
var ssDBID = '179OhQjvAQZ9QMYoXptBn2AcxSeWVmFtao3bOZAyXroE'
var sheetImpostazioni = ssDB.getSheetByName('Impostazioni')
var sheetFilesAffidi = ssDB.getSheetByName('Files affidi')
var sheetDiffideDaInviare = ssDB.getSheetByName('Diffide da inviare')
var sheetDiffideInviate = ssDB.getSheetByName('Diffide inviate')
var sheetDettaglioFatture = ssDB.getSheetByName('Dettaglio fatture')
var sheetListaDiControllo = ssDB.getSheetByName('Lista di controllo')

//var sheetExportTestataPratiche = ssDB.getSheetByName('Export Testata Pratiche')
//var sheetExportDettagliFatture = ssDB.getSheetByName('Export Dettagli Fatture')
//var sheetExportRecapiti = ssDB.getSheetByName('Export Recapiti')
//var sheetExportLettere = ssDB.getSheetByName('Export Lettere')


// STAMPA DIFFIDE
IDTemplateDiffide = '19jxUUO7Jx6DZj8vTQ0IyzTEJbTW22HVgjlQB6UUWzOQ'
IDTemplateEtichette = '1oIORY45a0o-6g_giDDVzlATl1q_T9o_EolszHox5mxg'
IDTemplateListaDiControllo = '1HWlWgwnQkuViSpXQTM_1dp7rRXijA43FWf_3SmqCX_o'
IDSettingsFolder = '0BznBNzYR5OHDdW9yWFZYQzhWRmM'
IDFolder = '0BznBNzYR5OHDRThmWkFFUXZFcEE' // folder diffide stampate
urlFirma = 'https://drive.google.com/open?id=0BznBNzYR5OHDaFd3bXRKTHBNcEU'

//EXPORT
IDFolderExport = '0BznBNzYR5OHDZVFfMU85aUtqeGM'
var urlExportTestataPratiche = 'https://docs.google.com/a/scenariopubblico.com/spreadsheets/d/1tnz0XCFKv1svrVGSJ6Fro3i3XgcqaQlTgn1Fn69sCyM/edit'
var ssExportTestataPratiche = SpreadsheetApp.openByUrl(urlExportTestataPratiche)
var sheetExportTestataPratiche = ssExportTestataPratiche.getSheetByName('Export Testata Pratiche')

var urlExportRecapiti = 'https://docs.google.com/spreadsheets/d/1Igi3hX_j8zly0gF-rLdpXTMG6hVyLWg1fCc4g076eek/edit'
var ssExportRecapiti = SpreadsheetApp.openByUrl(urlExportRecapiti)
var sheetExportRecapiti = ssExportRecapiti.getSheetByName('Export Recapiti')

var urlExportLettere = 'https://docs.google.com/spreadsheets/d/1BQJMGeHWBhtSEYnjBv7hZk4KBgKT4gQs8d-3w5ueMGg/edit'
var ssExportLettere = SpreadsheetApp.openByUrl(urlExportLettere)
var sheetExportLettere = ssExportLettere.getSheetByName('Export Lettere')

var urlExportDettagliFatture = 'https://docs.google.com/spreadsheets/d/1HvZZkQE5sJ79o5nW38p_OJp6sT_Wpd0pf6za_94h8JM/edit'
var ssExportDettagliFatture = SpreadsheetApp.openByUrl(urlExportDettagliFatture)
var sheetExportDettagliFatture = ssExportDettagliFatture.getSheetByName('Export Dettagli Fatture')

/*
**** IOL release ******

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

//EXPORT
IDFolderExport = '0BznBNzYR5OHDZVFfMU85aUtqeGM'
*/



