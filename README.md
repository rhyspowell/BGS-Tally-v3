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
I will detail how to do this soon.

# Market Sell
- Shows against the owner of the station (not always the same as the system)
- Doesn't include Black Market profit

Black Market sales are identified as ["BlackMarket":true]  in the event line, eg,

```json
{ "timestamp":"2021-05-25T12:20:21Z", "event":"MarketSell", "MarketID":3227073536, "Type":"slaves", "Count":50, "SellPrice":15800, "TotalSale":790000, "AvgPricePaid":11166, "IllegalGoods":true, "BlackMarket":true }
```

Unlike normal market sales, which show as;

```json
{ "timestamp":"2021-05-24T22:00:56Z", "event":"MarketSell", "MarketID":3227322112, "Type":"gold", "Count":1, "SellPrice":9744, "TotalSale":9744, "AvgPricePaid":0 }
```

In both cases the MarketID identifies the station, and thus the owner of the station;

```json
{ "timestamp":"2021-05-24T21:47:50Z", "event":"Docked", "StationName":"Jones City", "StationType":"Coriolis", "Taxi":false, "Multicrew":false, "StarSystem":"Akba Atacab", "SystemAddress":6405843096290, "MarketID":3227322112, "StationFaction":{ "Name":"Fatal Shadows", "FactionState":"Boom" }, "StationGovernment":"$government_Democracy;", "StationGovernment_Localised":"Democracy", "StationServices":[ "dock", "autodock", "blackmarket", "commodities", "contacts", "exploration", "missions", "outfitting", "crewlounge", "rearm", "refuel", "repair", "shipyard", "tuning", "engineer", "missionsgenerated", "flightcontroller", "stationoperations", "powerplay", "searchrescue", "materialtrader", "stationMenu", "shop", "livery", "socialspace" ], "StationEconomy":"$economy_Refinery;", "StationEconomy_Localised":"Refinery", "StationEconomies":[ { "Name":"$economy_Refinery;", "Name_Localised":"Refinery", "Proportion":0.880000 }, { "Name":"$economy_Industrial;", "Name_Localised":"Industrial", "Proportion":0.120000 } ], "DistFromStarLS":31.633158 }
```

