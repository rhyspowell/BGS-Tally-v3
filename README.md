# BGS-Tally-V3
A plugin to count BGS work and send to a google sheet
designed for use by multiple Cmdr's

Compatabile with Python 3 release

# Installation
Download the [latest release](https://github.com/tezw21/BGS-Tally-V3/releases/tag/3.0) of BGS Tally
 - In EDMC, on the Plugins settings tab press the “Open” button. This reveals the plugins folder where this app looks for plugins.
 - Open the .zip archive that you downloaded and move the folder contained inside into the plugins folder.

You will need to re-start EDMC for it to notice the new plugin.

# Usage
BGS Tally v3.0 counts all the BGS work you do for any faction, in any system.
It is highly reccomended that EDMC is started before ED is lauched as Data is recorded at startup and then when you dock at a station. Not doing this can result in missing data.
The data is now stored on a google sheet linked by following the directions below.
Once the credentials have been set multipe Cmdr's can download the JSON file and will be able to add their work to the day's totals
The plugin can paused in the BSG Tally v3.0 tab in settings.

Only four positive actions are counted at the moment. Mission inf + , Total trade profit sold to Faction controlled station, Cartographic data sold to Faction controlled station, Bounties issued by named Faction. These total during the session and reset at server tick.
On the server tick a new sheet is created by the first Cmdr to report after the tick. Previous sheets are left so a histroy can be kept. 

Users will be notified if there is a new version is avalible.

# Getting Google Sheet Credentials
This plugin requries a google service account linked to a google sheet.
If you want to have multiple cmdr's writing to the one sheet the credentials only need to be created once and copied into each cmdr's BGS tally folder

Create a Google Sheet, it only requires one sheet and the Plugin will create the rest as it goes

Enable API Access for a Project

Head to Google Developers Console and create a new project (or select the one you already have).
In the box labeled “Search for APIs and Services”, search for “Google Drive API” and enable it.
In the box labeled “Search for APIs and Services”, search for “Google Sheets API” and enable it.

Creating a Service Account

A service account is a special type of Google account intended to represent a non-human user that needs to authenticate and be authorized to access data in Google APIs [sic].
Since it’s a separate account, by default it does not have access to any spreadsheet until you share it with this account. Just like any other Google account.

Here’s how to get one:

Enable API Access for a Project if you haven’t done it yet.
Go to “APIs & Services > Credentials” and choose “Create credentials > Service account key”.
Fill out the form
Click “Create key”
Select “JSON” and click “Create”
You will automatically download a JSON file with credentials. It may look like this:

{
    "type": "service_account",
    "project_id": "api-project-XXX",
    "private_key_id": "2cd … ba4",
    "private_key": "-----BEGIN PRIVATE KEY-----\nNrDyLw … jINQh/9\n-----END PRIVATE KEY-----\n",
    "client_email": "473000000000-yoursisdifferent@developer.gserviceaccount.com",
    "client_id": "473 … hd.apps.googleusercontent.com",
    ...
}
Remeber the path to the downloaded credentials file. Also, in the next step you’ll need the value of client_email from this file.

Very important! Go to your spreadsheet and share it with a client_email from the step above. Just like you do with any other Google account. If you don’t do this, BGS tally can't write to the sheet.
Move the downloaded file to your BGS tally folder. You can get the folder from EDMC -> settings -> plugins
