

// ***** IOL DEV ******

//FOLDERS

// Affidi
var affidiDaImportareFolderId = '0BznBNzYR5OHDa2M5U21tUkJzYTA'
var affidiDaImportareFolder = DriveApp.getFolderById(affidiDaImportareFolderId)

var affidiOriginaliFolderId = '0BznBNzYR5OHDVGw3NTJsSTNIbVE'
var affidiOriginaliFolder = DriveApp.getFolderById(affidiDaImportareFolderId)

// Report
var reportOriginaliFolderId = '0BznBNzYR5OHDd0FsajNfNWZNUU0'
var reportOriginaliFolder = DriveApp.getFolderById(reportOriginaliFolderId)

var reportConvertitiFolderId = '0BznBNzYR5OHDcE5mbWxfaXdGdnc'
var reportConvertitiFolder = DriveApp.getFolderById(reportConvertitiFolderId)


// Spreadsheet
var urlDB = 'https://docs.google.com/spreadsheets/d/1yWfXqFBqcliyMmr4yNcbIEh_86Dhfp3QW9d5dltqKAA/edit'
var ssDB = SpreadsheetApp.openByUrl(urlDB)
var ssDBID = '1yWfXqFBqcliyMmr4yNcbIEh_86Dhfp3QW9d5dltqKAA'
var sheetImpostazioni = ssDB.getSheetByName('Impostazioni')
var sheetFilesAffidi = ssDB.getSheetByName('Files affidi')
var sheetFilesReport = ssDB.getSheetByName('Files report')
var sheetDiffideDaInviare = ssDB.getSheetByName('Diffide da inviare')
var sheetDettaglioFatture = ssDB.getSheetByName('Dettaglio fatture')
var sheetListaDiControllo = ssDB.getSheetByName('Lista di controllo')
var sheetInfoLettere = ssDB.getSheetByName('Info Lettere')


// STAMPA DIFFIDE
IDTemplateDiffide = '19jxUUO7Jx6DZj8vTQ0IyzTEJbTW22HVgjlQB6UUWzOQ'
IDTemplateEtichette = '1oIORY45a0o-6g_giDDVzlATl1q_T9o_EolszHox5mxg'
IDTemplateListaDiControllo = '1HWlWgwnQkuViSpXQTM_1dp7rRXijA43FWf_3SmqCX_o'
IDSettingsFolder = '0BznBNzYR5OHDekFneHBtUVFPMEE'
IDFolder = '0BznBNzYR5OHDV21nVW50ZW9FVjQ' // folder diffide stampate
urlFirma = 'https://drive.google.com/a/scenariopubblico.com/file/d/0BznBNzYR5OHDRUo3QjJXNnpRTmc/view?usp=sharing'

//EXPORT
IDFolderExport = '0BznBNzYR5OHDM0gxdGhzcTdOMzA'
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

//LETTERE

var urlFileEsportazione = 'https://docs.google.com/spreadsheets/d/1irJ6rGBFU84g7pwqZWMDTg-K7yrR6l3BlPCGKr-Muyg/edit'
var ssFileEsportazione = SpreadsheetApp.openByUrl(urlFileEsportazione)
var IDTemplateChiusuraContestazioni = '1d5ZGC_DbG9iUmzC40tMaH5HaO2cOkiv_CFnqe4Xtnj4'
urlFirmaPGiacona = 'https://drive.google.com/open?id=0BznBNzYR5OHDUjR0REUyb19PWk1aWmpVSWV0OFdOUDRrNzhR '


/*
//**** IOL release ******

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
IDFolderExport = '0BwA0zaKu6qcrUE1DVS1mbFVjeXc'
var urlExportTestataPratiche = 'https://docs.google.com/spreadsheets/d/1rpGvt_lv_d50_SIz6cApIzY_4DnLytf_Ln2YM5Tf2W0/edit'
var ssExportTestataPratiche = SpreadsheetApp.openByUrl(urlExportTestataPratiche)
var sheetExportTestataPratiche = ssExportTestataPratiche.getSheetByName('Export Testata Pratiche')

var urlExportRecapiti = 'https://docs.google.com/spreadsheets/d/1S4gLH1XAkhf4r5NtnWuGv2jBjfwUgOE43BTZQlfT06s/edit'
var ssExportRecapiti = SpreadsheetApp.openByUrl(urlExportRecapiti)
var sheetExportRecapiti = ssExportRecapiti.getSheetByName('Export Recapiti')

var urlExportLettere = 'https://docs.google.com/spreadsheets/d/15OBIzmRkNzwFXf1eDbGu6p4HMSDvJjWNd-k0HgU2BBk/edit'
var ssExportLettere = SpreadsheetApp.openByUrl(urlExportLettere)
var sheetExportLettere = ssExportLettere.getSheetByName('Export Lettere')

var urlExportDettagliFatture = 'https://docs.google.com/spreadsheets/d/1H_e68UFk60qG_RrNB0QrumJkvodXgRAO5_yF0SwC6_Y/edit'
var ssExportDettagliFatture = SpreadsheetApp.openByUrl(urlExportDettagliFatture)
var sheetExportDettagliFatture = ssExportDettagliFatture.getSheetByName('Export Dettagli Fatture')

*/


