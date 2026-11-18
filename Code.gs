// SATSET TECH Tournament Management System v10
// Backend Code.gs
function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'index';
  if (page === 'public' || page === 'viewer') {
    return HtmlService.createHtmlOutputFromFile('public-viewer');
  } else if (page === 'promotion') {
    return HtmlService.createHtmlOutputFromFile('promotion');
  }
  return HtmlService.createHtmlOutputFromFile('index');
}
