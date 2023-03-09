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
    log_this("Running status turned on");
  } else {
    config_sheet.getRange("B13").setValue("");
    config_sheet.getRange("B18").setValue("");
    log_this("Running status turned off (Total People:"+config_sheet.getRange("B16").getValue()+" ,RemoteIDs Added:"+config_sheet.getRange("B15").getValue()+")");
  }
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

function get_people_to_update() {
  var current_position = config_sheet.getRange("B18").getValue().replace(/([^ ]+)(.*)/gi,"$1").replace(/[^0-9]/gi,"").replace(/^$/,0);
  var total = config_sheet.getRange("B18").getValue().replace(/(^[^ ]+) of ([0-9]*)$/gi,"$2").replace(/[^0-9]/gi,"").replace(/^$/,0);
  var login = {headers: {Authorization: "Basic " + Utilities.base64Encode(username + ":" + password)}};
  do {
    var jsondata = UrlFetchApp.fetch("https://api.planningcenteronline.com/people/v2/people?per_page=100&offset="+current_position, login);
    var headers = jsondata.getAllHeaders();

    //If retry-after is set API limits have been reached
    if (typeof headers["retry-after"] === 'undefined') {
      var object = JSON.parse(jsondata.getContentText());
      total = object.meta.total_count
      config_sheet.getRange("B18").setValue(current_position+" of "+total);
      current_position = load_people_to_data_sheet(object.data);
      config_sheet.getRange("B18").setValue(current_position+" of "+total);
    } else {
      Utilities.sleep(headers["retry-after"]*1000);
    }
  } while (current_position < total)
}

function load_people_to_data_sheet(data) {
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
  return last_row + data.length - 1;
}