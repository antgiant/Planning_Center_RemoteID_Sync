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
      var username = config_sheet.getRange(config.username).getValue();
      var password = config_sheet.getRange(config.password).getValue();
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
        log_this("Config Values Updated (username:"+username.replace(/([a-z0-9])/gi,"*")+", password:"+password.replace(/([a-z0-9])/gi,"*")+")");
      }
      
      function update_running_status(is_running) {
        Logger.log("Starting update_running_status as "+is_running);
        config_sheet.getRange(config.is_running).setValue(is_running);
        config_sheet.getRange(config.left_to_create).setValue("?");
        log_this("Info Screen updated");
        if(is_running) {
          turn_on_sync();
          log_this("Running status turned on");
        } else {
          config_sheet.getRange(config.left_to_create).setValue("?");
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
      
      function onOpen() {
        Logger.log("Spreadsheet opened");
        update_config();
        var ui = SpreadsheetApp.getUi();
      
        ui.createMenu('Planning Center Sync')
            .addItem('Toggle Sync (On/Off)', 'toggle')
            .addToUi();
        Logger.log("Menu item added");
      }
      
      function toggle() {
        Logger.log("Toggling Running status");
        var ui = SpreadsheetApp.getUi();
        if (!config_sheet.getRange(config.is_running).getValue()) {
          if (config_sheet.getRange(config.username).getValue().toString().length == 64
            && config_sheet.getRange(config.password).getValue().toString().length == 64) {
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
        Logger.log("Running status toggle complete");
      }
      
      function turn_on_sync() {
        // Logger.log("Turning on repeating trigger");
        // ScriptApp.newTrigger("get_people_to_update")
        //         .timeBased()
        //         .everyMinutes(10)
        //         .create();
        // log_this("Turned on repeating process (trigger) that performs initial loading of people.");
      }
      
      function turn_off_sync() {
        Logger.log("Clearing all repeating triggers");
        // clear any existing triggers
        var triggers = ScriptApp.getProjectTriggers();
        for (var i = 0; i < triggers.length; i++) {
          ScriptApp.deleteTrigger(triggers[i]);
        }
        log_this("Turned off repeating process (trigger) that performs initial loading of people.");
      }
      
      function get_people_to_update() {
        var created_total = config_sheet.getRange(config.total_created).getValue().toString().replace(/[^0-9]/gi,"").replace(/^$/,0);

        var login = {
          headers: {
            Authorization: "Basic " + Utilities.base64Encode(username + ":" + password),
            "X-PCO-API-Version": config.api_version
          }
        };
        do {
          var now = new Date();
          var temp_url = "https://api.planningcenteronline.com/people/v2/people?order=created_at&per_page=10&where[remote_id]=&filter[ne]=organization_admins"
          Logger.log("Calling " + temp_url);
          var jsondata = UrlFetchApp.fetch(temp_url, login);
          var headers = jsondata.getAllHeaders();
      
          //If retry-after is set API limits have been reached
          if (typeof headers["retry-after"] === 'undefined') {
            Logger.log("No API delay requested by Planning Center loading data from JSON");
            var object = JSON.parse(jsondata.getContentText());

            log_this("Processing Batch of "+object.data.length+" people with "+object.meta.total_count+" people remaining.");
            config_sheet.getRange(config.last_check_time).setValue(now.toLocaleString());
            config_sheet.getRange(config.left_to_create).setValue(object.meta.total_count);
            created_total = process_people(object.data, created_total);
            config_sheet.getRange(config.total_created).setValue(created_total);
            config_sheet.getRange(config.left_to_create).setValue(object.meta.total_count - object.data.length);
            log_this("Batch Complete")
          } else {
            log_this("Planning Center API Limit reached. Delaying for "+headers["retry-after"]+" seconds as requested by Planning Center API.");
            Utilities.sleep(headers["retry-after"]*1000);
          }
        } while (object.meta.total_count > 0)       
      }
      
      function process_people(data, current_count) {
        for (i = 0; i < data.length; i++) {
          temp = [
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
          ];
        }
        return current_count + data.length;
      }