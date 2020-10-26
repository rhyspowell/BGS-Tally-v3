import myNotebook as nb
import sys
import json
import requests
from config import config
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
this.VersionNo = "2.1.1"
this.FactionNames = []
this.TodayData = {}
this.MissionLog = {}
this.DataIndex = 0
this.Status = "Active"
this.TickTime = ""
this.State = tk.IntVar()
this.cred = ''  # google sheet service account cred's path to file


def plugin_prefs(parent, cmdr, is_beta):
    """
   Return a TK Frame for adding to the EDMC settings dialog.
   """

    frame = nb.Frame(parent)
    nb.Label(frame, text="BGS Tally v" + this.VersionNo).grid(column=0, sticky=tk.W)
    """
   reset = nb.Button(frame, text="Reset Counter").place(x=0 , y=290)
   """
    nb.Checkbutton(frame, text="Make BGS Tally Active", variable=this.Status, onvalue="Active",
                   offvalue="Paused").grid()

    return frame


def prefs_changed(cmdr, is_beta):
    """
   Save settings.
   """
    this.StatusLabel["text"] = this.Status.get()


def plugin_start(plugin_dir):
    """
   Load this plugin into EDMC
   """
    this.Dir = plugin_dir
    this.cred = os.path.join(this.Dir, "service_account.json")
    file = os.path.join(this.Dir, "Today Data.txt")

    if path.exists(file):
        with open(file) as json_file:
            this.TodayData = json.load(json_file)
            z = len(this.TodayData)
            for i in range(1, z + 1):
                x = str(i)
                this.TodayData[i] = this.TodayData[x]
                del this.TodayData[x]

    this.LastTick = tk.StringVar(value=config.get("XLastTick"))
    this.TickTime = tk.StringVar(value=config.get("XTickTime"))
    this.Status = tk.StringVar(value=config.get("XStatus"))
    this.DataIndex = tk.IntVar(value=config.get("xIndex"))
    this.StationFaction = tk.StringVar(value=config.get("XStation"))

    # this.LastTick.set("12")

    response = requests.get('https://api.github.com/repos/tezw21/BGS-Tally/releases/latest')  # check latest version
    latest = response.json()
    this.GitVersion = latest['tag_name']
    #  tick check and counter reset
    response = requests.get('https://elitebgs.app/api/ebgs/v4/ticks')  # get current tick and reset if changed
    tick = response.json()
    this.CurrentTick = tick[0]['_id']
    this.TickTime = tick[0]['time']
    print(this.LastTick.get())
    print(this.CurrentTick)
    if this.LastTick.get() != this.CurrentTick:
        this.LastTick.set(this.CurrentTick)
        this.TodayData = {}
        print("Tick auto reset happened")
    # create google sheet
    Google_sheet_int()

    return "BGS Tally v2"


def plugin_start3(plugin_dir):
    return plugin_start(plugin_dir)


def plugin_stop():
    """
    EDMC is closing
    """
    save_data()

    print("Farewell cruel world!")


def plugin_app(parent):
    """    Create a frame for the EDMC main window    """
    this.frame = tk.Frame(parent)

    Title = tk.Label(this.frame, text="BGS Tally v" + this.VersionNo)
    Title.grid(row=0, column=0, sticky=tk.W)
    if version_tuple(this.GitVersion) > version_tuple(this.VersionNo):
        title2 = tk.Label(this.frame, text="New version available", fg="blue", cursor="hand2")
        title2.grid(row=0, column=1, sticky=tk.W,)
        title2.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/tezw21/BGS-Tally/releases"))

    tk.Button(this.frame, text='Data Today', command=display_data).grid(row=1, column=0, padx=3)
    tk.Button(this.frame, text='Mission Log', command=display_MissionLog).grid(row=1, column=1, padx=3)
    tk.Label(this.frame, text="Status:").grid(row=2, column=0, sticky=tk.W)
    tk.Label(this.frame, text="Last Tick:").grid(row=3, column=0, sticky=tk.W)
    this.StatusLabel = tk.Label(this.frame, text=this.Status.get())
    this.StatusLabel.grid(row=2, column=1, sticky=tk.W)
    this.TimeLabel = tk.Label(this.frame, text=tick_format(this.TickTime)).grid(row=3, column=1, sticky=tk.W)

    return this.frame


def journal_entry(cmdr, is_beta, system, station, entry, state):
    if this.Status.get() != "Active":
        print('Paused')
        return

    if entry['event'] == 'Docked':

        this.StationFaction.set(entry['StationFaction']['Name'])  # set station controlling faction name

        #  tick check and counter reset
        response = requests.get('https://elitebgs.app/api/ebgs/v4/ticks')  # get current tick and reset if changed
        tick = response.json()
        this.CurrentTick = tick[0]['_id']
        this.TickTime = tick[0]['time']
        print(this.LastTick.get())
        print(this.CurrentTick)
        print(this.TickTime)
        if this.LastTick.get() != this.CurrentTick:
            this.LastTick.set(this.CurrentTick)
            this.TodayData = {}
            this.TimeLabel = tk.Label(this.frame, text=tick_format(this.TickTime)).grid(row=3, column=1, sticky=tk.W)
            theme.update(this.frame)
            print("Tick auto reset happened")
        # set up new sheet at tick reset
        Google_sheet_int()
        # TodayData creation
        x = len(this.TodayData)
        if (x >= 1):
            for y in range(1, x+1):
                if entry['StarSystem'] == this.TodayData[y][0]['System']:
                    this.DataIndex.set(y)
                    print('system in data')
                    print(this.DataIndex.get())
                    Sheet_Insert_New_System(y)
                    return
        this.TodayData[x+1] = [{'System': entry['StarSystem'], 'SystemAddress': entry['SystemAddress'], 'Factions': []}]
        this.DataIndex.set(x+1)
        z = len(this.FactionNames)
        for i in range(0, z):
            inf = this.FactionStates['Factions'][i]['Influence'] * 100
            inf = round(inf, 2)
            this.TodayData[x+1][0]['Factions'].append({'Faction': this.FactionNames[i], 'INF': inf, 'State': 0, 'MissionPoints': 0, 'MissionFailed': 0, 'TradeProfit': 0, 'Bounties': 0, 'Bonds': 0, 'CartData': 0, 'Murders': 0})
        else:
            this.TodayData = {1: [{'System': entry['StarSystem'], 'SystemAddress': entry['SystemAddress'], 'Factions': []}]}
            z = len(this.FactionNames)
            this.DataIndex.set(1)
            for i in range(0, z):
                inf = this.FactionStates['Factions'][i]['Influence'] * 100
                inf = round(inf, 2)
                this.TodayData[1][0]['Factions'].append({'Faction': this.FactionNames[i], 'INF': inf, 'State': 0, 'MissionPoints': 0, 'MissionFailed': 0, 'TradeProfit': 0, 'Bounties': 0, 'Bonds': 0, 'CartData': 0, 'Murders': 0})
        Sheet_Insert_New_System(x+1)

    if entry['event'] == 'Location':  # get factions at startup
        this.FactionNames = []
        this.FactionStates = {'Factions': []}
        z = 0
        for i in entry['Factions']:
            if i['Name'] != "Pilots' Federation Local Branch":
                this.FactionNames.append(i['Name'])
                # get faction States and Influence
                this.FactionStates['Factions'].append({'Faction': i['Name'],
                                                      'Influence': i['Influence'],
                                                       'States': []})

                try:
                    for x in i['ActiveStates']:
                        this.FactionStates['Factions'][z]['States'].append({'State': x['State']})
                except KeyError:
                    this.FactionStates['Factions'][z]['States'].append({'State': 'None'})
                z += 1

    if entry['event'] == 'FSDJump':  # get factions at jump
        this.FactionNames = []
        this.FactionStates = {'Factions': []}
        z = 0
        for i in entry['Factions']:
            if i['Name'] != "Pilots' Federation Local Branch":
                this.FactionNames.append(i['Name'])
                this.FactionStates['Factions'].append(
                    {'Faction': i['Name'], 'Influence': i['Influence'],
                     'States': []})

                try:
                    for x in i['ActiveStates']:
                        this.FactionStates['Factions'][z]['States'].append({'State': x['State']})
                except KeyError:
                    this.FactionStates['Factions'][z]['States'].append({'State': 'None'})
                z += 1

    if entry['event'] == 'CommitCrime' and entry['CrimeType'] == 'murder':  # crime murder needs tested
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in entry['Faction']:
            for x in range(0, t):
                if z['Faction'] == this.TodayData[this.DataIndex.get()][0]['Faction'][x]['Faction']:
                    this.TodayData[this.DataIndex.get()][0]['Faction'][x]['Murders'] += 1
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = this.DataIndex.get()
                    data = 1
                    Sheet_Commit_Data(system, index, 'Murders', data)
        save_data()

    if entry['event'] == 'MissionCompleted':  # get mission influence value
        fe = entry['FactionEffects']
        print("mission completed")
        for i in fe:
            fe3 = i['Faction']
            print(fe3)
            fe4 = i['Influence']
            for x in fe4:
                fe6 = x['SystemAddress']
                inf = len(x['Influence'])
                for y in this.TodayData:
                    if fe6 == this.TodayData[y][0]['SystemAddress']:
                        t = len(this.TodayData[y][0]['Factions'])
                        system = this.TodayData[y][0]['System']

                    for z in range(0, t):
                        if fe3 == this.TodayData[y][0]['Factions'][z]['Faction']:
                            this.TodayData[y][0]['Factions'][z]['MissionPoints'] += inf
                            Sheet_Commit_Data(system, z, 'Mission', inf)
        save_data()

    if entry['event'] == 'SellExplorationData' or entry['event'] == "MultiSellExplorationData":  # get carto data value
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in range(0, t):
            if this.StationFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                this.TodayData[this.DataIndex.get()][0]['Factions'][z]['CartData'] += entry['TotalEarnings']
                system = this.TodayData[this.DataIndex.get()][0]['System']
                index = this.DataIndex.get()
                data = entry['TotalEarnings']
                Sheet_Commit_Data(system, index, 'Expo', data)
        save_data()

    if entry['event'] == 'RedeemVoucher' and entry['Type'] == 'bounty':  # bounties collected
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in entry['Factions']:
            for x in range(0, t):
                if z['Faction'] == this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Faction']:
                    this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Bounties'] += z['Amount']
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = this.DataIndex.get()
                    data = z['Amount']
                    Sheet_Commit_Data(system, index, 'Bounty', data)
        save_data()

    if entry['event'] == 'RedeemVoucher' and entry['Type'] == 'bond':  # bonds collected
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in entry['Factions']:
            for x in range(0, t):
                if z['Faction'] == this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Faction']:
                    this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Bonds'] += z['Amount']
                    system = this.TodayData[this.DataIndex.get()][0]['System']
                    index = this.DataIndex.get()
                    data = z['Amount']
                    Sheet_Commit_Data(system, index, 'Bonds', data)
        save_data()

    if entry['event'] == 'MarketSell':  # Trade Profit
        t = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for z in range(0, t):
            if this.StationFaction.get() == this.TodayData[this.DataIndex.get()][0]['Factions'][z]['Faction']:
                cost = entry['Count'] * entry['AvgPricePaid']
                profit = entry['TotalSale'] - cost
                this.TodayData[this.DataIndex.get()][0]['Factions'][z]['TradeProfit'] += profit
                system = this.TodayData[this.DataIndex.get()][0]['System']
                index = this.DataIndex.get()
                data = profit
                Sheet_Commit_Data(system, index, 'Trade', data)
        save_data()


def version_tuple(version):
    try:
        ret = tuple(map(int, version.split(".")))
    except:
        ret = (0,)
    return ret


def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])


def display_data():
    form = tk.Toplevel(this.frame)
    form.title("BGS Tally v" + this.VersionNo + " - Data Today")
    form.geometry("1250x560")
    # tk.Label(this.frame, text="BGS Tally v" + this.VersionNo)

    tab_parent = ttk.Notebook(form)

    for i in this.TodayData:
        tab = ttk.Frame(tab_parent)
        tab_parent.add(tab, text=this.TodayData[i][0]['System'])
        FactionLabel = tk.Label(tab, text="Faction")
        INFLabel = tk.Label(tab, text="INF")
        STLabel = tk.Label(tab, text="State")
        MPLabel = tk.Label(tab, text="Mission Points")
        MFLabel = tk.Label(tab, text="Mission Failed")
        TPLabel = tk.Label(tab, text="Trade Profit")
        BountyLabel = tk.Label(tab, text="Bounties")
        BondLabel = tk.Label(tab, text="Bonds")
        CDLabel = tk.Label(tab, text="Cart Data")
        CMLabel = tk.Label(tab, text="Murders")

        FactionLabel.grid(row=0, column=0)
        INFLabel.grid(row=0, column=1)
        STLabel.grid(row=0, column=2)
        MPLabel.grid(row=0, column=3, )
        MFLabel.grid(row=0, column=4)
        TPLabel.grid(row=0, column=5)
        BountyLabel.grid(row=0, column=6)
        BondLabel.grid(row=0, column=7)
        CDLabel.grid(row=0, column=8)
        CMLabel.grid(row=0, column=9)
        z = len(this.TodayData[i][0]['Factions'])
        for x in range(0, z):
            FactionName = tk.Label(tab, text=this.TodayData[i][0]['Factions'][x]['Faction'])
            FactionName.grid(row=x + 1, column=0, sticky=tk.W)
            INF = tk.Label(tab, text=this.TodayData[i][0]['Factions'][x]['INF'])
            INF.grid(row=x + 1, column=1)
            State = tk.Label(tab, text=this.TodayData[i][0]['Factions'][x]['State'])
            State.grid(row=x + 1, column=2)
            Missions = tk.Label(tab, text=this.TodayData[i][0]['Factions'][x]['MissionPoints'])
            Missions.grid(row=x + 1, column=3)
            MissionsF = tk.Label(tab, text=this.TodayData[i][0]['Factions'][x]['MissionFailed'])
            MissionsF.grid(row=x + 1, column=4)
            Trade = tk.Label(tab, text=human_format(this.TodayData[i][0]['Factions'][x]['TradeProfit']))
            Trade.grid(row=x + 1, column=5)
            Bounty = tk.Label(tab, text=human_format(this.TodayData[i][0]['Factions'][x]['Bounties']))
            Bounty.grid(row=x + 1, column=6)
            Bonds = tk.Label(tab, text=human_format(this.TodayData[i][0]['Factions'][x]['Bonds']))
            Bonds.grid(row=x + 1, column=7)
            CartData = tk.Label(tab, text=human_format(this.TodayData[i][0]['Factions'][x]['CartData']))
            CartData.grid(row=x + 1, column=8)
            Murders = tk.Label(tab, text=human_format(this.TodayData[i][0]['Factions'][x]['Murders']))
            Murders.grid(row=x + 1, column=9)
    tab_parent.pack(expand=1, fill='both')


def display_MissionLog():
    form = tk.Toplevel(this.frame)
    form.title("BGS Tally v" + this.VersionNo + " - MissionLog")
    form.geometry("1000x560")

    tab_parent = ttk.Notebook(form)

    for i in this.MissionLog:
        tab = ttk.Frame(tab_parent)
        tab_parent.add(tab, text=this.MissionLog[i][0]['System'])
        factionlabel = tk.Label(tab, text="Faction")
        milabel = tk.Label(tab, text="Mission INF")
        mdlabel = tk.Label(tab, text="Mission ID")

        factionlabel.grid(row=0, column=0)
        milabel.grid(row=0, column=1, )
        mdlabel.grid(row=0, column=2)
        z = len(this.MissionLog[i][0]['Factions'])
        for x in range(0, z):
            factionname = tk.Label(tab, text=this.MissionLog[i][0]['Factions'][x]['Faction'])
            factionname.grid(row=x + 1, column=0, sticky=tk.W)
            missionsinf = tk.Label(tab, text=this.MissionLog[i][0]['Factions'][x]['MissionINF'])
            missionsinf.grid(row=x + 1, column=1)
            missionsid = tk.Label(tab, text=this.MissionLog[i][0]['Factions'][x]['MissionID'])
            missionsid.grid(row=x + 1, column=2)
    tab_parent.pack(expand=1, fill='both')


def tick_format(ticktime):
    datetime1 = ticktime.split('T')
    x = datetime1[0]
    z = datetime1[1]
    y = x.split('-')
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
    datetimetick = time1 + ' UTC ' + date1
    return (datetimetick)


def save_data():
    config.set('XLastTick', this.CurrentTick)
    config.set('XTickTime', this.TickTime)
    config.set('XStatus', this.Status.get())
    config.set('xIndex', this.DataIndex.get())
    config.set('XStation', this.StationFaction.get())

    file = os.path.join(this.Dir, "Today Data.txt")
    with open(file, 'w') as outfile:
        json.dump(this.TodayData, outfile)

    file = os.path.join(this.Dir, "Mission Log.txt")
    with open(file, 'w') as outfile:
        json.dump(this.MissionLog, outfile)


def Google_sheet_int():
    # start google sheet data store
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("BSG Tally Store")
    try:
        worksheet = sh.worksheet(this.TickTime)
    except:
        worksheet = sh.add_worksheet(title=this.TickTime, rows="2000", cols="20")
        worksheet.update('A1', '# of Systems')
        worksheet.update('B1', 0)
        set_column_width(worksheet, 'A', 300)


def Sheet_Insert_New_System(index):
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("BSG Tally Store")
    worksheet = sh.worksheet(this.TickTime)
    FactionName = []
    FactionINF = []
    system = this.TodayData[this.DataIndex.get()][0]['System']

    try:
        cell = worksheet.find(system)
    except gspread.exceptions.CellNotFound:
        z = len(this.TodayData[this.DataIndex.get()][0]['Factions'])
        for x in range(0, z):
            FactionName.append([this.TodayData[this.DataIndex.get()][0]['Factions'][x]['Faction']])
            FactionINF.append([this.TodayData[this.DataIndex.get()][0]['Factions'][x]['INF']])
        no_of_systems = int(worksheet.acell('B1').value)
        if no_of_systems == 0:
            no_of_systems += 1
            worksheet.update('B1', no_of_systems)
            # worksheet.update('A2:J3', [['System', system],['Faction', 'Mission +', 'Trade', 'Bounties',
            # 'Carto Data']])
            worksheet.batch_update([{'range': 'A2:J3', 'values': [['System', system],
                                                                  ['Faction', 'INF', 'State', 'Mission +',
                                                                   'Mission Failed', 'Trade', 'Bounties', 'Bonds',
                                                                   'Carto Data', 'Murders']]},
                                    {'range': 'A4:A11', 'values': FactionName},
                                    {'range': 'B4:B11', 'values': FactionINF},
                                    {'range': 'C4:J11',
                                     'values': [[0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0]]}])
        else:
            row = no_of_systems * 11 + 2
            no_of_systems += 1
            worksheet.update('B1', no_of_systems)
            range1 = 'A' + str(row) + ':J' + str(row + 1)
            range2 = 'A' + str(row + 2) + ':A' + str(row + 10)
            range3 = 'B' + str(row + 2) + ':B' + str(row + 10)
            range4 = 'C' + str(row + 2) + ':J' + str(row + 10)
            worksheet.batch_update([{'range': range1, 'values': [['System', system],
                                                                 ['Faction', 'INF', 'State', 'Mission +',
                                                                  'Mission Failed', 'Trade', 'Bounties', 'Bonds',
                                                                  'Carto Data', 'Murders']]},
                                    {'range': range2, 'values': FactionName},
                                    {'range': range3, 'values': FactionINF},
                                    {'range': range4,
                                     'values': [[0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0],
                                                [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0]]}])


def Sheet_Commit_Data(system, index, event, data):
    gc = gspread.service_account(filename=this.cred)
    sh = gc.open("BSG Tally Store")
    worksheet = sh.worksheet(this.TickTime)
    cell1 = worksheet.find(system)
    FactionRow = cell1.row + 2 + index
    if event == "INF":
        cell = worksheet.cell(FactionRow, 2).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 2, Total)

    if event == "State":
        cell = worksheet.cell(FactionRow, 3).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 3, Total)

    if event == "Mission":
        cell = worksheet.cell(FactionRow, 4).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 4, Total)

    if event == "MissionFailed":
        cell = worksheet.cell(FactionRow, 5).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 5, Total)

    if event == "Trade":
        cell = worksheet.cell(FactionRow, 6).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 6, Total)

    if event == "Bounty":
        cell = worksheet.cell(FactionRow, 7).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 7, Total)

    if event == "Bonds":
        cell = worksheet.cell(FactionRow, 8).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 8, Total)

    if event == "Expo":
        cell = worksheet.cell(FactionRow, 9).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 9, Total)

    if event == "Murders":
        cell = worksheet.cell(FactionRow, 10).value
        Total = int(cell) + data
        worksheet.update_cell(FactionRow, 10, Total)
