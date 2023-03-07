var username = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B2").getValue();
var password = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B3").getValue();
var overwrite = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B4").getValue();

function update_config() {
  username = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B2").getValue();
  password = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B3").getValue();
  overwrite = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B4").getValue();
}

function update_running_status() {
  var is_running = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B8").getValue();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration").getRange("B17").setValue(is_running);
}

function onEdit(e) {
  if (e.range.getSheet().getName() == "Configuration") {
    if (e.range.getA1Notation() == "B2"
          || e.range.getA1Notation() == "B3"
          || e.range.getA1Notation() == "B4") {
      update_config()
    }
    if (e.range.getA1Notation() == "B8") {
      update_running_status()
    }
  }
}