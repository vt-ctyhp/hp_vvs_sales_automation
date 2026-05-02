function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function createTestReceipt() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([
    "Test Receipt",
    new Date()
  ]);
  
  return "Receipt created successfully!";
}