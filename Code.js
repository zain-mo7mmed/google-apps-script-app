function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Google Apps Script App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getServerTime() {
  return new Date().toLocaleString();
}
