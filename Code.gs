var config_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
var username = config_sheet.getRange("B2").getValue();
var password = config_sheet.getRange("B3").getValue();
var overwrite = config_sheet.getRange("B4").getValue();

function log_this(message) {
  var log_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity Log");
  var now = new Date();
  log_sheet.insertRowBefore(2);
  log_sheet.getRange("A2").setValue(now.toLocaleString());
  log_sheet.getRange("B2").setValue(message);
}

function update_config() {
  username = config_sheet.getRange("B2").getValue();
  password = config_sheet.getRange("B3").getValue();
  overwrite = config_sheet.getRange("B4").getValue();
  log_this("Config Values Updated");
}

function update_running_status() {
  var is_running = config_sheet.getRange("B8").getValue();
  config_sheet.getRange("B17").setValue(is_running);
  if(is_running) {
    var now = new Date();
    config_sheet.getRange("B13").setValue(now.toLocaleString());
    config_sheet.getRange("B14").setValue("");
    config_sheet.getRange("B15").setValue(0);
    config_sheet.getRange("B16").setValue(0);
    config_sheet.getRange("B18").setValue("? of ?");
  } else {
    config_sheet.getRange("B13").setValue("");
    config_sheet.getRange("B18").setValue("");
  }
  log_this("Running status changed to "+is_running);
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