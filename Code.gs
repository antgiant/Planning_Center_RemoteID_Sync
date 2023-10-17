var scriptProperties = PropertiesService.getScriptProperties();

var config = {
  api_version: "2023-02-15",
  username: "B2",
  password: "B3",
  is_running: "B6",
  last_check_time: "B7",
  total_created: "B8",
  left_to_create: "B9"
}
var config_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
var username = scriptProperties.getProperty('username');
var password = scriptProperties.getProperty('password');
Logger.log("Script Loaded and Config Values Set");

function log_this(message) {
  var log_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity Log");
  var now = new Date();
  log_sheet.insertRowBefore(2);
  log_sheet.getRange("A2").setValue(now.toLocaleString());
  log_sheet.getRange("B2").setValue(message);
  Logger.log(message);
}

function update_config() {
  Logger.log("Updating Config");
  username = config_sheet.getRange(config.username).getValue();
  password = config_sheet.getRange(config.password).getValue();

  if (username.replace(/([*])/gi,"") != "") { //Don't overwrite with redacted version
    scriptProperties.setProperty('username', username);
    log_this("Username updated");
    config_sheet.getRange(config.username).setValue(username.replace(/([a-z0-9])/gi,"*"));
  } else {
    username = scriptProperties.getProperty('username');
  }
  if (password.replace(/([*])/gi,"") != "") { //Don't overwrite with redacted version
    scriptProperties.setProperty('password', password);
    log_this("Password updated");
    config_sheet.getRange(config.password).setValue(password.replace(/([a-z0-9])/gi,"*"));
  } else {
    password = scriptProperties.getProperty('password');
  }
}

function update_running_status(is_running) {
  Logger.log("Starting update_running_status as "+is_running);
  config_sheet.getRange(config.is_running).setValue(is_running);
  log_this("Info Screen updated");
  if(is_running) {
    turn_on_sync();
    log_this("Running status turned on");
  } else {
    config_sheet.getRange(config.left_to_create).setValue("¿"+config_sheet.getRange(config.left_to_create).getValue().toString()+"?");
    turn_off_sync();
    log_this("Running status turned off");
  }
}

function onEdit(e) {
  if (e.range.getSheet().getName() == "Configuration") {
    if (e.range.getA1Notation() == config.username
          || e.range.getA1Notation() == config.password) {
      Logger.log("Updating Config due to edit of config values");
      update_config()
    }
  }
}

function installedOnEdit(e) {
  if (e.range.getSheet().getName() == "Configuration" && e.range.getA1Notation() == config.is_running) {
    log_this("One Time Sync triggered by clicking checkbox on Configuration");
    update_running_status(e.value)
  }
}

function onOpen() {
  Logger.log("Spreadsheet opened");
  update_config();
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Planning Center Sync')
      .addItem('Toggle One Time Sync (On/Off)', 'toggle')
      .addItem('Turn On Once a Day Sync', 'daily_sync_on')
      .addItem('Turn Off Once a Day Sync', 'daily_sync_off')
      .addToUi();
  Logger.log("Menu item added");
}

function addEditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    //Only add one of these
    if (triggers[i].getHandlerFunction() == "installedOnEdit" && triggers[i].getEventType() == "ON_EDIT"){
      return;
    }
  }
  Logger.log("Adding installedOnEdit trigger to make checkbox work");
  ScriptApp.newTrigger('installedOnEdit')
  .forSpreadsheet(SpreadsheetApp.getActive())
  .onEdit()
  .create();
}

function toggle() {
  addEditTrigger()
  Logger.log("Toggling Running status");
  if (!config_sheet.getRange(config.is_running).getValue()) {
    if (username.length == 64
      && password.length == 64) {
        log_this("Turning on Sync");
        update_running_status(true);
        SpreadsheetApp.getActive().toast('One Time Sync turned on');
        get_people_to_update();
      } else {
        var ui = SpreadsheetApp.getUi();
        log_this("No/Bad Application ID and/or Secret (aka Username and/or password)");
        ui.alert('Please Enter Application ID & Secret to Turn on Sync');
      }
  } else {
    update_running_status(false);
    SpreadsheetApp.getActive().toast('One Time Sync turned off');
  }
  Logger.log("Running status toggle complete");
}

function turn_on_sync() {
  Logger.log("Turning on repeating trigger");
  ScriptApp.newTrigger("get_people_to_update")
          .timeBased()
          .everyMinutes(10)
          .create();
  log_this("Turned on repeating process (trigger) that performs one time loading of people.");
}

function daily_sync_on() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    //Only add one of these
    if (triggers[i].getHandlerFunction() == "do_daily_sync"){
      return;
    }
  }
  Logger.log("Turning on Daily Sync");
  ScriptApp.newTrigger("do_daily_sync")
          .timeBased()
          .everyDays(1)
          .create();
  log_this("Daily Sync Turned On.");
}

function do_daily_sync() {
  get_people_to_update();
}

function turn_off_sync() {
  Logger.log("Clearing all repeating triggers");
  // clear any existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    //Don't delete the installedOnEdit or do_daily_sync trigger
    if (triggers[i].getHandlerFunction() != "do_daily_sync" && (triggers[i].getHandlerFunction() != "installedOnEdit" || triggers[i].getEventType() != "ON_EDIT")) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  log_this("Turned off repeating process (trigger) that performs one time loading of people.");
}

function daily_sync_off() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    //Don't delete the installedOnEdit tritter
    if (triggers[i].getHandlerFunction() == "do_daily_sync") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  log_this("Daily Sync Turned Off.");
}

function get_people_to_update() {
  var created_total = Number(config_sheet.getRange(config.total_created).getValue().toString().replace(/[^0-9]/gi,"").replace(/^$/,0));

  var login = {
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(username + ":" + password),
      "X-PCO-API-Version": config.api_version
    },
    'muteHttpExceptions': true
  };
  var total = 0;
  do {
    var now = new Date();
    var temp_url = "https://api.planningcenteronline.com/people/v2/people?order=created_at&per_page=10&where[remote_id]=&filter[ne]=organization_admins"
    Logger.log("Calling " + temp_url);
    try {
      var jsondata = UrlFetchApp.fetch(temp_url, login);
      var headers = jsondata.getAllHeaders();
      var responseCode = jsondata.getResponseCode();
  
      //If Retry-After is set API limits have been reached
      if (typeof headers["Retry-After"] === 'undefined') {
        Logger.log("No API delay requested by Planning Center loading data from JSON");
        var object = JSON.parse(jsondata.getContentText());
        total = object.meta.total_count;

        log_this("Processing Batch of "+object.data.length+" people with "+total+" people remaining.");
        config_sheet.getRange(config.last_check_time).setValue(now.toLocaleString());
        config_sheet.getRange(config.left_to_create).setValue(total);
        created_total = update_people(object.data, created_total);
        config_sheet.getRange(config.total_created).setValue(created_total);
        config_sheet.getRange(config.left_to_create).setValue(total - object.data.length);
        log_this("Batch Complete")
      } else {
        log_this("Planning Center API Limit reached. Delaying for "+headers["Retry-After"]+" seconds as requested by Planning Center API.");
        Utilities.sleep(headers["Retry-After"]*1000);
      }
      if (responseCode != 200 && responseCode != 429) {
        var responseBody = response.getContentText();
        log_this(Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody))
      }
    } catch (e) {
      log_this(e.toString());
    }
  } while (total > 0)
  //When complete turn off triggers
  update_running_status(false);
}

function update_people(data, current_count) {
  for (i = 0; i < data.length; i++) {
    var payload = {
      "data": {
        "type": data[i].type,
        "id": data[i].id,
        "attributes": {
          'remote_id': data[i].id
        }
      }
    }
    var options = {
      'method' : 'patch',
      'headers': {
                  Authorization: "Basic " + Utilities.base64Encode(username + ":" + password),
                  "X-PCO-API-Version": config.api_version
                },
      'contentType': 'application/json',
      // Convert the JavaScript object to a JSON string.
      'payload' : JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    try {
      var jsondata = UrlFetchApp.fetch(data[i].links.self, options);
      var headers = jsondata.getAllHeaders();
      var responseCode = jsondata.getResponseCode();
  
      //If Retry-After is set API limits have been reached
      if (typeof headers["Retry-After"] !== 'undefined') {
        log_this("Planning Center API Limit reached. Delaying for "+headers["Retry-After"]+" seconds as requested by Planning Center API.");
        Utilities.sleep(headers["Retry-After"]*1000);
        i--;
      }
      if (responseCode == 200) {
        Logger.log("Sucessfully updated "+data[i].attributes.name+" with Remote ID of ("+data[i].id+")");
      } else if (responseCode != 429) {
        var responseBody = response.getContentText();
        log_this(Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody))
      }
    } catch (e) {
      log_this(e.toString());
    }
  }
  return current_count + data.length;
}

function doPost(e) {
  log_this("Planning Center (Post) Webhook Called.");
  update_config();
  get_people_to_update();
}