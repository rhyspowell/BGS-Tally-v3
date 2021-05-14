import logging
import myNotebook as nb
import sys
import json
import requests
from config import config, appname
from theme import theme
import webbrowser
import os.path
import gspread
from gspread_formatting import *
from os import path

try:
    # Python 2
    import Tkinter as tk
    import ttk
except ModuleNotFoundError:
    # Python 3
    import tkinter as tk
    from tkinter import ttk


this = sys.modules[__name__]  # For holding module globals
this.VersionNo = "4.2.9"
this.FactionNames = []
this.TodayData = {}
this.YesterdayData = {}
this.DataIndex = 0
this.Status = "Active"
this.TickTime = ""
this.State = tk.IntVar()
# this.cred = 'client_secret.json' #google sheet service account cred's path to file

# This could also be returned from plugin_start3()
plugin_name = os.path.basename(os.path.dirname(__file__))

# A Logger is used per 'found' plugin to make it easy to include the plugin's
# folder name in the logging output format.
# NB: plugin_name here *must* be the plugin's folder name as per the preceding
#     code, else the logger won't be properly set up.
logger = logging.getLogger(f"{appname}.{plugin_name}")

# If the Logger has handlers then it was already set up by the core code, else
# it needs setting up here.
if not logger.hasHandlers():
    level = logging.INFO  # So logger.info(...) is equivalent to print()

    logger.setLevel(level)
    logger_channel = logging.StreamHandler()
    logger_formatter = logging.Formatter(
        f"%(asctime)s - %(name)s - %(levelname)s - %(module)s:%(lineno)d:%(funcName)s: %(message)s"
    )
    logger_formatter.default_time_format = "%Y-%m-%d %H:%M:%S"
    logger_formatter.default_msec_format = "%s.%03d"
    logger_channel.setFormatter(logger_formatter)
    logger.addHandler(logger_channel)

logger.info("BGS Tally has started")


def plugin_prefs(parent, cmdr, is_beta):
    logger.debug("plugin_prefs function")
    """
    Return a TK Frame for adding to the EDMC settings dialog.
    """

    frame = nb.Frame(parent)
    nb.Label(frame, text="BGS Tally v" + this.VersionNo).grid(column=0, sticky=tk.W)
    """
   reset = nb.Button(frame, text="Reset Counter").place(x=0 , y=290)
   """
    nb.Checkbutton(
        frame,
        text="Make BGS Tally Active",
        variable=this.Status,
        onvalue="Active",
        offvalue="Paused",
    ).grid()
    return frame


def prefs_changed(cmdr, is_beta):
    logger.debug("prefs_changed function")
    """
    Save settings.
    """
    this.StatusLabel["text"] = this.Status.get()


def check_version():
    try:
        logger.debug("get version from api")
        response = requests.get(
            "https://api.github.com/repos/rhyspowell/BGS-Tally-v3/releases/latest"
        )  # check latest version
        latest = response.json()
        this.GitVersion = latest["tag_name"][1:]
        if this.GitVersion > this.VersionNo:
            return True
    except KeyError:
        logger.error("Failed to get latest version from the github api")
        this.GitVersion = "Connection Error"
    return False


def plugin_start3(plugin_dir):
    logger.debug("Plugin_start function")
    """
    Load this plugin into EDMC
    """

    this.Dir = plugin_dir
    try:
        this.cred = os.path.join(this.Dir, "client_secret.json")
    except FileNotFoundError:
        logger.error("missing client secret file for gspread")

    file = os.path.join(this.Dir, "Today Data.txt")

    if path.exists(file):
        with open(file) as json_file:
            this.TodayData = json.load(json_file)
            z = len(this.TodayData)
            for i in range(1, z + 1):
                x = str(i)
                this.TodayData[i] = this.TodayData[x]
                del this.TodayData[x]

    file = os.path.join(this.Dir, "Yesterday Data.txt")

    if path.exists(file):
        with open(file) as json_file:
            this.YesterdayData = json.load(json_file)
            z = len(this.YesterdayData)
            for i in range(1, z + 1):
                x = str(i)
                this.YesterdayData[i] = this.YesterdayData[x]
                del this.YesterdayData[x]

    this.LastTick = tk.StringVar(value=config.get("XLastTick"))
    this.TickTime = tk.StringVar(value=config.get("XTickTime"))
    this.Status = tk.StringVar(value=config.get("XStatus"))
    this.DataIndex = tk.IntVar(value=config.get("xIndex"))
    this.StationFaction = tk.StringVar(value=config.get("XStation"))

    #  tick check and counter reset
    response = requests.get(
        "https://elitebgs.app/api/ebgs/v5/ticks"
    )  # get current tick and reset if changed
    tick = response.json()
    this.CurrentTick = tick[0]["_id"]
    this.TickTime = tick[0]["time"]
    logger.info(this.LastTick.get())
    logger.info(this.CurrentTick)
    if this.LastTick.get() != this.CurrentTick:
        this.LastTick.set(this.CurrentTick)
        this.YesterdayData = this.TodayData
        this.TodayData = {}
        print("Tick auto reset happened")
    # create google sheet
    Google_sheet_int()

    return "BGS Tally v3"


def plugin_stop():
    """
    EDMC is closing
    """
    save_data()

    logger.info("Farewell cruel world!")


def plugin_app(parent):
    """
    Create a frame for the EDMC main window
    """
    logger.debug("plugin_app function")
    this.frame = tk.Frame(parent)

    Title = tk.Label(this.frame, text="BGS Tally v" + this.VersionNo)
    Title.grid(row=0, column=0, sticky=tk.W)

    update = check_version()
    if update:
        title2 = tk.Label(
            this.frame, text="New version available", fg="blue", cursor="hand2"
        )
        title2.grid(
            row=0,
            column=1,
            sticky=tk.W,
        )
        title2.bind(
            "<Button-1>",
            lambda e: webbrowser.open_new(
                "https://github.com/rhyspowell/BGS-Tally-v3/releases"
            ),
        )

    tk.Button(this.frame, text="Data Today", command=display_data).grid(
        row=1, column=0, padx=3
    )
    tk.Button(this.frame, text="Data Yesterday", command=display_yesterdaydata).grid(
        row=1, column=1, padx=3
    )
    tk.Label(this.frame, text="Status:").grid(row=2, column=0, sticky=tk.W)
    tk.Label(this.frame, text="Last Tick:").grid(row=3, column=0, sticky=tk.W)
    this.StatusLabel = tk.Label(this.frame, text=this.Status.get())
    this.StatusLabel.grid(row=2, column=1, sticky=tk.W)
    this.TimeLabel = tk.Label(this.frame, text=tick_format(this.TickTime)).grid(
        row=3, column=1, sticky=tk.W
    )

    return this.frame


def faction_processing(entry):
    FactionNames = []
    FactionStates = {"Factions": []}
    z = 0
    for i in entry["Factions"]:
        if i["Name"] != "Pilots' Federation Local Branch":
            FactionNames.append(i["Name"])
            FactionStates["Factions"].append(
                {
                    "Faction": i["Name"],
                    "Happiness": i["Happiness_Localised"],
                    "States": [],
                }
            )

            try:
                for x in i["ActiveStates"]:
                    FactionStates["Factions"][z]["States"].append({"State": x["State"]})
            except KeyError:
                FactionStates["Factions"][z]["States"].append({"State": "None"})
            z += 1
    logging.debug(FactionStates)
    return FactionNames, FactionStates


def docked():
    pass


def missioncompleted():
    pass


def explorationdata():
    pass


def bounty():
    pass


def combatbond():
    pass


def marketsell():
    pass


def missionaccepted():
    logger.debug("Mission Accepted")
    pass


def missionfailed():
    logger.debug("Mission failed")
    pass


def missionabandoned():
    logger.debug("Mission Abandoned")
    pass


def missionsstartup():
    logger.debug("Missions at startup")
    pass


def ussdrop():
    logger.debug("USS Drop")
    pass


def journal_entry(cmdr, is_beta, system, station, entry, state):
    logger.debug("journal_entry function")

    if this.Status.get() != "Active":
        print("Paused")
        return

    if (
        entry["event"] == "Location"
        or entry["event"] == "FSDJump"
        or entry["event"] == "CarrierJump"
    ):  # get factions
        this.FactionNames = []
        this.FactionStates = {"Factions": []}
        this.FactionNames, this.FactionStates = faction_processing(entry)

    if entry["event"] == "Docked":  # enter system and faction named

        this.StationFaction.set(
            entry["StationFaction"]["Name"]
        )  # set controlling faction name

        #  tick check and counter reset
        response = requests.get(
            "https://elitebgs.app/api/ebgs/v5/ticks"
        )  # get current tick and reset if changed
        tick = response.json()
        this.CurrentTick = tick[0]["_id"]
        this.TickTime = tick[0]["time"]
        print(this.LastTick.get())
        print(this.CurrentTick)
        print(this.TickTime)
        if this.LastTick.get() != this.CurrentTick:
            this.LastTick.set(this.CurrentTick)
            this.YesterdayData = this.TodayData
            this.TodayData = {}
            this.TimeLabel = tk.Label(this.frame, text=tick_format(this.TickTime)).grid(
                row=3, column=1, sticky=tk.W
            )
            theme.update(this.frame)
            print("Tick auto reset happened")
        # set up new sheet at tick reset
        Google_sheet_int()

        x = len(this.TodayData)
        if x >= 1:
            for y in range(1, x + 1):
                if entry["StarSystem"] == this.TodayData[y][0]["System"]:
                    this.DataIndex.set(y)
                    print("system in data")

                    Sheet_Insert_New_System(y)
                    return
            this.TodayData[x + 1] = [
                {
                    "System": entry["StarSystem"],
                    "SystemAddress": entry["SystemAddress"],
                    "Factions": [],
                }
            ]
            this.DataIndex.set(x + 1)
            z = len(this.FactionNames)
            for i in range(0, z):
                this.TodayData[x + 1][0]["Factions"].append(
                    {
                        "Faction": this.FactionNames[i],
                        "MissionPoints": 0,
                        "TradeProfit": 0,
                        "Bounties": 0,
                        "CartData": 0,
                    }
                )
        else:
            this.TodayData = {
                1: [
                    {
                        "System": entry["StarSystem"],
                        "SystemAddress": entry["SystemAddress"],
                        "Factions": [],
                    }
                ]
            }
            z = len(this.FactionNames)
            this.DataIndex.set(1)
            for i in range(0, z):
                this.TodayData[1][0]["Factions"].append(
                    {
                        "Faction": this.FactionNames[i],
                        "MissionPoints": 0,
                        "TradeProfit": 0,
                        "Bounties": 0,
                        "CartData": 0,
                    }
                )
        Sheet_Insert_New_System(x + 1)

    if entry["event"] == "MissionCompleted":  # get mission influence value
        fe = entry["FactionEffects"]
        print("mission completed")
        for i in fe:
            fe3 = i["Faction"]
            print(fe3)
            fe4 = i["Influence"]
            for x in fe4:
                fe6 = x["SystemAddress"]
                inf = len(x["Influence"])
                for y in this.TodayData:
                    if fe6 == this.TodayData[y][0]["SystemAddress"]:
                        t = len(this.TodayData[y][0]["Factions"])
                        system = this.TodayData[y][0]["System"]

                        for z in range(0, t):
                            if fe3 == this.TodayData[y][0]["Factions"][z]["Faction"]:
                                this.TodayData[y][0]["Factions"][z][
                                    "MissionPoints"
                                ] += inf
                                Sheet_Commit_Data(system, z, "Mission", inf)
        save_data()

    if (
        entry["event"] == "SellExplorationData"
        or entry["event"] == "MultiSellExplorationData"
    ):  # get carto data value
        t = len(this.TodayData[this.DataIndex.get()][0]["Factions"])
        for z in range(0, t):
            if (
                this.StationFaction.get()
                == this.TodayData[this.DataIndex.get()][0]["Factions"][z]["Faction"]
            ):
                this.TodayData[this.DataIndex.get()][0]["Factions"][z][
                    "CartData"
                ] += entry["TotalEarnings"]
                system = this.TodayData[this.DataIndex.get()][0]["System"]
                index = this.DataIndex.get()
                data = entry["TotalEarnings"]
                Sheet_Commit_Data(system, index, "Expo", data)
        save_data()

    if (
        entry["event"] == "RedeemVoucher" and entry["Type"] == "bounty"
    ):  # bounties collected
        """
        example bounty
        { "timestamp":"2021-05-06T23:08:25Z", "event":"RedeemVoucher", "Type":"bounty", "Amount":7656828, "Factions":[ { "Faction":"Fatal Shadows", "Amount":7656828 } ] }
        """
        logger.debug("Running bounty")
        faction = entry["Factions"][0]["Faction"]
        amount = entry["Factions"][0]["Amount"]

        t = len(this.TodayData[this.DataIndex.get()][0]["Factions"])

        for x in range(0, t):
            if (
                faction
                == this.TodayData[this.DataIndex.get()][0]["Factions"][x]["Faction"]
            ):
                this.TodayData[this.DataIndex.get()][0]["Factions"][x][
                    "Bounties"
                ] += amount
                system = this.TodayData[this.DataIndex.get()][0]["System"]
                logger.debug("DataIndex: " + str(this.DataIndex))
                logger.debug(this.DataIndex.get)
                index = this.DataIndex.get()
                logger.debug("Index: " + str(index))
                Sheet_Commit_Data(system, index, "Bounty", amount)
        save_data()

    if entry["event"] == "RedeemVoucher" and entry["Type"] == "CombatBond":
        """
        example combat bound
        { "timestamp":"2021-05-06T23:08:19Z", "event":"RedeemVoucher", "Type":"CombatBond", "Amount":9098729, "Faction":"Fatal Shadows" }
        """
        logger.debug("Sheet Commit Data")
        sh = gspread.service_account(filename=this.cred).open("BSG Tally Store")
        worksheet = sh.worksheet(this.TickTime)
        try:
            current_value = int(worksheet.get('R1C8').first())
        except Exception as e:
            logger.error("Cell value error: " + str(e))
            current_value = 0
        Total = current_value + entry["Amount"]
        worksheet.update_cell('R1C8', Total)

    if entry["event"] == "MarketSell":  # Trade Profit
        t = len(this.TodayData[this.DataIndex.get()][0]["Factions"])
        for z in range(0, t):
            if (
                this.StationFaction.get()
                == this.TodayData[this.DataIndex.get()][0]["Factions"][z]["Faction"]
            ):
                cost = entry["Count"] * entry["AvgPricePaid"]
                profit = entry["TotalSale"] - cost
                this.TodayData[this.DataIndex.get()][0]["Factions"][z][
                    "TradeProfit"
                ] += profit
                system = this.TodayData[this.DataIndex.get()][0]["System"]
                index = this.DataIndex.get()
                data = profit
                Sheet_Commit_Data(system, index, "Trade", data)
        save_data()

    if entry["event"] == "MissionAccepted":  # mission accpeted
        missionaccepted()

    if entry["event"] == "MissionFailed":  # mission failed
        missionfailed()

    if entry["event"] == "MissionAbandoned":
        missionabandoned()

    if entry["event"] == "Missions":  # missions on startup
        missionsstartup()

    if entry["event"] == "USSDrop":
        ussdrop()


def human_format(num):
    num = float("{:.3g}".format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return "{}{}".format(
        "{:f}".format(num).rstrip("0").rstrip("."), ["", "K", "M", "B", "T"][magnitude]
    )


def display_data():
    form = tk.Toplevel(this.frame)
    form.title("BGS Tally v" + this.VersionNo + " - Data Today")
    form.geometry("500x280")
    # tk.Label(this.frame, text="BGS Tally v" + this.VersionNo)

    tab_parent = ttk.Notebook(form)

    for i in this.TodayData:
        tab = ttk.Frame(tab_parent)
        tab_parent.add(tab, text=this.TodayData[i][0]["System"])
        FactionLabel = tk.Label(tab, text="Faction")
        MPLabel = tk.Label(tab, text="Misson Points")
        TPLabel = tk.Label(tab, text="Trade Profit")
        BountyLabel = tk.Label(tab, text="Bounties")
        CDLabel = tk.Label(tab, text="Cart Data")

        FactionLabel.grid(row=0, column=0)
        MPLabel.grid(
            row=0,
            column=1,
        )
        TPLabel.grid(row=0, column=2)
        BountyLabel.grid(row=0, column=3)
        CDLabel.grid(row=0, column=4)
        z = len(this.TodayData[i][0]["Factions"])
        for x in range(0, z):
            FactionName = tk.Label(
                tab, text=this.TodayData[i][0]["Factions"][x]["Faction"]
            )
            FactionName.grid(row=x + 1, column=0, sticky=tk.W)
            Missions = tk.Label(
                tab, text=this.TodayData[i][0]["Factions"][x]["MissionPoints"]
            )
            Missions.grid(row=x + 1, column=1)
            Trade = tk.Label(
                tab,
                text=human_format(this.TodayData[i][0]["Factions"][x]["TradeProfit"]),
            )
            Trade.grid(row=x + 1, column=2)
            Bounty = tk.Label(
                tab, text=human_format(this.TodayData[i][0]["Factions"][x]["Bounties"])
            )
            Bounty.grid(row=x + 1, column=3)
            CartData = tk.Label(
                tab, text=human_format(this.TodayData[i][0]["Factions"][x]["CartData"])
            )
            CartData.grid(row=x + 1, column=4)
    tab_parent.pack(expand=1, fill="both")


def display_yesterdaydata():
    form = tk.Toplevel(this.frame)
    form.title("BGS Tally v" + this.VersionNo + " - Data Yesterday")
    form.geometry("500x280")

    tab_parent = ttk.Notebook(form)

    for i in this.YesterdayData:
        tab = ttk.Frame(tab_parent)
        tab_parent.add(tab, text=this.YesterdayData[i][0]["System"])
        FactionLabel = tk.Label(tab, text="Faction")
        MPLabel = tk.Label(tab, text="Misson Points")
        TPLabel = tk.Label(tab, text="Trade Profit")
        BountyLabel = tk.Label(tab, text="Bounties")
        CDLabel = tk.Label(tab, text="Cart Data")

        FactionLabel.grid(row=0, column=0)
        MPLabel.grid(
            row=0,
            column=1,
        )
        TPLabel.grid(row=0, column=2)
        BountyLabel.grid(row=0, column=3)
        CDLabel.grid(row=0, column=4)
        z = len(this.YesterdayData[i][0]["Factions"])
        for x in range(0, z):
            FactionName = tk.Label(
                tab, text=this.YesterdayData[i][0]["Factions"][x]["Faction"]
            )
            FactionName.grid(row=x + 1, column=0, sticky=tk.W)
            Missions = tk.Label(
                tab, text=this.YesterdayData[i][0]["Factions"][x]["MissionPoints"]
            )
            Missions.grid(row=x + 1, column=1)
            Trade = tk.Label(
                tab,
                text=human_format(
                    this.YesterdayData[i][0]["Factions"][x]["TradeProfit"]
                ),
            )
            Trade.grid(row=x + 1, column=2)
            Bounty = tk.Label(
                tab,
                text=human_format(this.YesterdayData[i][0]["Factions"][x]["Bounties"]),
            )
            Bounty.grid(row=x + 1, column=3)
            CartData = tk.Label(
                tab,
                text=human_format(this.YesterdayData[i][0]["Factions"][x]["CartData"]),
            )
            CartData.grid(row=x + 1, column=4)
    tab_parent.pack(expand=1, fill="both")


def tick_format(ticktime):
    datetime1 = ticktime.split("T")
    x = datetime1[0]
    z = datetime1[1]
    y = x.split("-")
    if y[1] == "01":
        month = "Jan"
    elif y[1] == "02":
        month = "Feb"
    elif y[1] == "03":
        month = "March"
    elif y[1] == "04":
        month = "April"
    elif y[1] == "05":
        month = "May"
    elif y[1] == "06":
        month = "June"
    elif y[1] == "07":
        month = "July"
    elif y[1] == "08":
        month = "Aug"
    elif y[1] == "09":
        month = "Sep"
    elif y[1] == "10":
        month = "Oct"
    elif y[1] == "11":
        month = "Nov"
    elif y[1] == "12":
        month = "Dec"
    date1 = y[2] + " " + month
    time1 = z[0:5]
    datetimetick = time1 + " UTC " + date1
    return datetimetick


def save_data():
    config.set("XLastTick", this.CurrentTick)
    config.set("XTickTime", this.TickTime)
    config.set("XStatus", this.Status.get())
    config.set("XIndex", this.DataIndex.get())
    config.set("XStation", this.StationFaction.get())

    file = os.path.join(this.Dir, "Today Data.txt")
    with open(file, "w") as outfile:
        json.dump(this.TodayData, outfile)

    file = os.path.join(this.Dir, "Yesterday Data.txt")
    with open(file, "w") as outfile:
        json.dump(this.YesterdayData, outfile)


def Google_sheet_int():
    # start google sheet data store
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("BSG Tally Store")
    try:
        worksheet = sh.worksheet(this.TickTime)
    except:
        worksheet = sh.add_worksheet(title=this.TickTime, rows="100", cols="20")
        worksheet.update("A1", "# of Systems")
        worksheet.update("B1", 0)
        set_column_width(worksheet, "A", 300)


def Sheet_Insert_New_System(index):
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("BSG Tally Store")
    worksheet = sh.worksheet(this.TickTime)
    FactionName = []
    system = this.TodayData[index][0]["System"]
    try:
        cell = worksheet.find(system)
    except gspread.exceptions.CellNotFound:
        z = len(this.TodayData[index][0]["Factions"])
        for x in range(0, z):
            FactionName.append([this.TodayData[index][0]["Factions"][x]["Faction"]])
        no_of_systems = int(worksheet.acell("B1").value)
        if no_of_systems == 0:
            no_of_systems += 1
            worksheet.update("B1", no_of_systems)
            # worksheet.update('A2:E3', [['System', system],['Faction', 'Mission +', 'Trade', 'Bounties', 'Carto Data']])
            worksheet.batch_update(
                [
                    {
                        "range": "A2:E3",
                        "values": [
                            ["System", system],
                            ["Faction", "Mission +", "Trade", "Bounties", "Carto Data"],
                        ],
                    },
                    {"range": "A4:A11", "values": FactionName},
                    {
                        "range": "B4:E11",
                        "values": [
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                        ],
                    },
                ]
            )
        else:
            row = no_of_systems * 11 + 2
            no_of_systems += 1
            worksheet.update("B1", no_of_systems)
            range1 = "A" + str(row) + ":E" + str(row + 1)
            range2 = "A" + str(row + 2) + ":A" + str(row + 10)
            range3 = "B" + str(row + 2) + ":E" + str(row + 10)
            worksheet.batch_update(
                [
                    {
                        "range": range1,
                        "values": [
                            ["System", system],
                            ["Faction", "Mission +", "Trade", "Bounties", "Carto Data"],
                        ],
                    },
                    {"range": range2, "values": FactionName},
                    {
                        "range": range3,
                        "values": [
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                            [0, 0, 0, 0],
                        ],
                    },
                ]
            )


def Sheet_Commit_Data(system, index, event, data):
    logger.debug("Sheet Commit Data")
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("BSG Tally Store")
    worksheet = sh.worksheet(this.TickTime)
    cell1 = worksheet.find(system)
    logger.debug("Cell1 is " + str(cell1))
    # Increase the value here by 1 as numbers start from 0
    FactionRow = cell1.row + 3 + index
    logger.debug("factions: " + this.FactionNames)
    logger.debug("Faction row: " + str(FactionRow))
    if event == "Mission":
        cell = worksheet.cell(FactionRow, 2).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 2, Total)

    if event == "Expo":
        cell = worksheet.cell(FactionRow, 5).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 5, Total)

    if event == "Bounty":
        cell = worksheet.cell(FactionRow, 4).value
        logger.debug("bounty cell" + str(cell))
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 4, Total)

    if event == "Trade":
        cell = worksheet.cell(FactionRow, 3).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 3, Total) 
