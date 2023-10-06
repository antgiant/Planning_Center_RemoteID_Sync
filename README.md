# Planning Center RemoteID Sync
This is a tool to synchronize PeopleID and RemoteID in Planning Center to allow for simplified bulk updates in Planning Center via the built-in Import/Export Tool

## Warnings
* This tool does not check for existing RemoteIDs. If you have them, you must make sure they do not overlap with PeopleIDs. I have no idea what will happen if they overlap, but I doubt it will be good.
* Organization Admins will not get RemoteIDs. Planning Center blocks updating Organization Admins via the APIs that this uses.

## How to Setup
1. [Open this Google Sheet](https://docs.google.com/spreadsheets/d/13MgPxryby493Eo3ijk-xBSJ5tzCOYyil2R3Yvdc2vY4/edit?usp=sharing) and make a copy ![Make a Copy](img/make_a_copy.jpg)
1. Visit [Planning Center's Developer API Center](https://api.planningcenteronline.com/)
1. Click the New Personal Access API Token Button ![New Personal Access API Token Button](img/new_personal_access_api_token.jpg)
1. Give it whatever description you would like and make sure that the People drop down has 2023-02-15 selected in it. ![Selecting 2023-02-15](img/selecting_2023-02-15.jpg)
1. Submit it ![Submit Button](img/submit.jpg)
1. Copy the Application ID and Secret into the Google Sheet you copied in step 1.
1. Click the Planning Center Sync menu and select any option ![Planning Center Sync Menu Button](img/planning_center_sync_menu.jpg)
1. Select Continue ![Authorization Required pop up](img/authorization_required.jpg)
1. Login with your Google Account
1. Click the little tiny Advanced Link ![Advanced Link](img/advanced_link.jpg)
1. Click the Go to Planning Center Sync (Unsafe) link ![Go to Planning Center Sync (unsafe) link](img/go_to_planning_center_link.jpg)
1. Click Allow so that the spreadsheet can call the Planning Center API ![Permission Dialog screenshot](img/allow_access_to_API_calls.jpg)
1. Configuration is now complete you can choose your favorite way to run this tool. (Note the Activity Log Sheet will give you additional feedback on what the tool is doing)

## Optional Features
### Daily updates via scheduled task
* You can choose to have this run once a day to keep new records up to date

### Realtime updates via Webhook
If you wish to have this run in realtime when new people are added in Planning Center you can do that via a webhook. Please follow these instructions to turn on the Webhook option.
1. In the Google Sheet select Extensions -> Apps Script ![Apps Script Menu](img/start_apps_script.jpg)
1. In the top right select Deploy -> New deployment ![New Deployment Menu](img/select_new_deployment.jpg)
1. Click the Gear -> Web app ![Web app menu screenshot](img/select_web_app.jpg)
1. Give it a descriptive name ![Example name](img/web_app_name.jpg)
1. Change access to Anyone ![Permission Menu](img/web_app_access.jpg)
1. Click Deploy ![Deploy Button](img/web_app_deploy.jpg)
1. Copy the URL ![URL copy screenshot](img/web_app_link.jpg)
1. Click Done ![Done Button](img/web_app_done.jpg)
1. Go to the [Planning Center Webhooks API page](https://api.planningcenteronline.com/webhooks)
1. Click the Add button ![Webhook Add Button](img/webhook_add.jpg)
1. Paste in the Endpoint URL from Google ![Endpoint URL example](img/webhook_endpoint_url.jpg)
1. Select Created under People -> Person ![screenshot of selecting person created](img/webhook_select_person_created.jpg)
1. Click Save ![Save button](img/webhook_save.jpg)
1. That is all. Your Google Sheet will now update RemoteIDs in real time.