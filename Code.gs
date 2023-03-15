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
  log_this("Config Values Updated (username:"+username.replace(/([a-z0-9])/gi,"*")+", password:"+password.replace(/([a-z0-9])/gi,"*")+", overwrite remote_id:"+overwrite+")");
}

function update_running_status(is_running) {
  config_sheet.getRange("B8").setValue(is_running);
  if(is_running) {
    var now = new Date();
    config_sheet.getRange("B9").setValue(now.toLocaleString());
    config_sheet.getRange("B10").setValue("");
    config_sheet.getRange("B11").setValue(0);
    config_sheet.getRange("B12").setValue(0);
    config_sheet.getRange("B13").setValue("? of ?");
    log_this("Info Screen updated");
    turn_on_sync();
    log_this("Running status turned on");
  } else {
    config_sheet.getRange("B9").setValue("");
    config_sheet.getRange("B13").setValue("");
    log_this("Info Screen updated");
    turn_off_sync();
    log_this("Running status turned off (Total People:"+config_sheet.getRange("B12").getValue()+" ,RemoteIDs Added:"+config_sheet.getRange("B11").getValue()+")");
  }
}

function onEdit(e) {
  if (e.range.getSheet().getName() == "Configuration") {
    if (e.range.getA1Notation() == "B2"
          || e.range.getA1Notation() == "B3"
          || e.range.getA1Notation() == "B4") {
      update_config()
    }
  }
}

function onOpen() {
  update_config();
  var ui = SpreadsheetApp.getUi();

    ui.createMenu('Planning Center Sync')
      .addItem('Toggle Sync (On/Off)', 'toggle')
        .addToUi();
}

function toggle() {
  var ui = SpreadsheetApp.getUi();
  if (!config_sheet.getRange("B8").getValue()) {
    if (config_sheet.getRange("B2").getValue().toString().length == 64
      && config_sheet.getRange("B3").getValue().toString().length == 64) {
        log_this("Turning on Sync");
        update_running_status(true);
        SpreadsheetApp.getActive().toast('Sync turned on');
      } else {
        log_this("No/Bad Application ID and/or Secret (aka Username and/or password)");
        ui.alert('Please Enter Application ID & Secret to Turn on Sync');
    }
  } else {
    update_running_status(false);
    SpreadsheetApp.getActive().toast('Sync turned off');
  }
  }

function turn_on_sync() {
  ScriptApp.newTrigger("get_people_to_update")
          .timeBased()
          .everyMinutes(10)
          .create();
  log_this("Turned on repeating process (trigger) that performs initial loading of people.");
}

function turn_off_sync() {
  // clear any existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  log_this("Turned off repeating process (trigger) that performs initial loading of people.");
}

function get_people_to_update() {
  log_this("Starting to load People into Data Sheet");
  var current_position = config_sheet.getRange("B13").getValue().replace(/([^ ]+)(.*)/gi,"$1").replace(/[^0-9]/gi,"").replace(/^$/,0);
  var total = config_sheet.getRange("B13").getValue().replace(/(^[^ ]+) of ([0-9]*)$/gi,"$2").replace(/[^0-9]/gi,"").replace(/^$/,0);
  var login = {headers: {Authorization: "Basic " + Utilities.base64Encode(username + ":" + password)}};
  do {
    var jsondata = UrlFetchApp.fetch("https://api.planningcenteronline.com/people/v2/people?per_page=100&offset="+current_position, login);
    var headers = jsondata.getAllHeaders();

    //If retry-after is set API limits have been reached
    if (typeof headers["retry-after"] === 'undefined') {
      var object = JSON.parse(jsondata.getContentText());
      current_position = load_people_to_data_sheet(object.data, current_position);
      if (total > object.meta.total_count) {
        //Total count has decreased. This means a record was deleted.
        // Move current position backwards by difference to ensure that no one is missed.
        
        current_position = current_position - (total - object.meta.total_count);
        if (current_position < 0) {
          //This is unlikely but possible in certian circumstances.
          current_position = 0;
        }
      }
      total = object.meta.total_count
      config_sheet.getRange("B13").setValue(current_position+" of "+total);
      log_this(current_position+" of "+total+" people loaded into Data Sheet")
    } else {
      log_this("Planning Center API Limit reached. Delaying for "+headers["retry-after"]+" seconds as requested by Planning Center API.");
      Utilities.sleep(headers["retry-after"]*1000);
    }
  } while (current_position < total)
  log_this("Completed First pass of loading People into Data Sheet");

  //Now do an immediate second pass to catch anyone added to system during this loading process
  // ****FIX ME*******
}

function load_people_to_data_sheet(data, current_count) {
  var data_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var last_row = (data_sheet.getRange("A1:A").getValues()).filter(String).length;
  var row = data_sheet.getRange("A"+(last_row+1)+":AD"+(last_row+data.length));
  var row_data = [];
  for (i = 0; i < data.length; i++) {
    row_data.push([
      data[i].id,
      data[i].attributes.name,
      data[i].attributes.given_name,
      data[i].attributes.first_name,
      data[i].attributes.nickname,
      data[i].attributes.middle_name,
      data[i].attributes.last_name,
      data[i].attributes.avatar,
      data[i].attributes.birthdate,
      data[i].attributes.anniversary,
      data[i].attributes.gender,
      data[i].attributes.demographic_avatar_url,
      data[i].attributes.grade,
      data[i].attributes.school_type,
      data[i].attributes.graduation_year,
      data[i].attributes.medical_notes,
      data[i].attributes.child,
      data[i].attributes.status,
      data[i].attributes.membership,
      data[i].attributes.inactivated_at,
      data[i].attributes.passed_background_check,
      data[i].attributes.created_at,
      data[i].attributes.updated_at,
      data[i].attributes.directory_status,
      data[i].attributes.people_permissions,
      data[i].attributes.can_create_forms,
      data[i].attributes.accounting_administrator,
      data[i].attributes.site_administrator,
      data[i].attributes.remote_id,
      data[i].links.self
    ]);
  }
  row.setValues(row_data);
  return current_count + data.length;
}

function erase_all_data() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data").getRange("A2:AD").clearContent();
}