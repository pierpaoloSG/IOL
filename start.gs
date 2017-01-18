function doGet() {

  return HtmlService
      .createTemplateFromFile('Index')
      .evaluate(); 
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}