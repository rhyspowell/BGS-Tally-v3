# BGS-Tally-V3
A plugin to count BGS work and send to a google sheet
Designed for use by multiple Cmdr's

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
I will detail how to do this soon.
For testing purposes the Credentials file for a test sheet is preloaded in the install.
Test sheet link - https://docs.google.com/spreadsheets/d/1uTkDn3ri2Xm4LZPh7ieYYGqf48y1sGQHu8QWa_mFXYI/edit?usp=sharing
