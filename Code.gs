var username,password,overwrite,frequency;

function initialize() {
  username = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B2").getValue();
  password = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B3").getValue();
  overwrite = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B4").getValue();
  frequency = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B5").getValue();
}
