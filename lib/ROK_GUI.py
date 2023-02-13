import glob
import os
import sqlite3 as sq
import sys
import threading
import datetime
import pandas as pd
from pathlib import Path

import PySimpleGUI as sg
import csv
import yaml

from pywinauto import application


class pyROK_GUI:

    def __init__(self):

        if sys.platform.startswith('win'):
            import ctypes
            # Make sure Pyinstaller icons are still grouped
            if sys.argv[0].endswith('.exe') == False:
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    u'TODO.pyROK.GUI')

        self.window: sg.Window = None
        self.WSDirectory: Path = ""
        self.KDList: list[int] = []
        self.kingdom: int = 0
        self.kingdomDir: Path = ""
        self.Project: str = ""
        self.ProjectList: list[str] = []
        self.ProjectDir: Path = ""
        self.ProjectScreenshotsDir: Path = ""
        self.ProjectConfigFile: Path = ""
        self.acquire_power: bool = False
        self.acquire_alliance: bool = False
        self.acquire_playerNane: bool = False
        self.acquire_T4: bool = False
        self.acquire_T5: bool = False
        self.acquire_dead: bool = False
        self.acquire_rssAssist: bool = False
        self.progressValue: int = 0
        self.configDone: threading.Event = threading.Event()
        self.close: threading.Event = threading.Event()
        self.status: bool = True
        self.players_to_acquire = 0
        self.previous_dates: list[str] = []
        self.defaultValues: dict[str:str | bool] = {'-DIR-': "Choose a folder...",
                                                    '-KD-': "Type or select a kingdom number",
                                                    '-PROJECT-': "Type or select a project",
                                                    'dates': "Select a poll date to delete",
                                                    '-PLAYER_NAME-': True,
                                                    '-PLAYER_POWER-': True,
                                                    '-PLAYER_T5-': True,
                                                    '-PLAYER_RSS-': True,
                                                    '-ALLIANCE_NAME-': True,
                                                    '-PLAYER_T4-': True,
                                                    '-PLAYER_DEAD-': True,
                                                    '-PLAYERS_TO_ACQUIRE-': 1000,
                                                    }
        col1 = [[sg.Text("working directory :")],
                [sg.Text("Kingdom :")],
                [sg.Text("Project :")]]
        col2 = [[sg.Input(default_text=self.defaultValues['-DIR-'], expand_x=True,
                          disabled=True, key='-DIR-', change_submits=True),
                 sg.FolderBrowse(target='-DIR-', enable_events=True, key='dir select')],
                [sg.Combo(values=self.KDList, key='-KD-', expand_x=True,
                          default_value=self.defaultValues['-KD-'], disabled=True),
                 sg.Button("Submit", enable_events=True, key='KD_valid')],
                [sg.Combo(values=self.ProjectList, key='-PROJECT-', expand_x=True,
                          default_value=self.defaultValues['-PROJECT-'], disabled=True),
                 sg.Button("Submit", enable_events=True, key='Project_valid')]]
        col3 = [[sg.Checkbox("Acquire Player Name", key='-PLAYER_NAME-',
                             default=self.defaultValues['-PLAYER_NAME-'])],
                [sg.Checkbox("Acquire Player Power", key='-PLAYER_POWER-',
                             default=self.defaultValues['-PLAYER_POWER-'])],
                [sg.Checkbox("Acquire Player T5 kills", key='-PLAYER_T5-',
                             default=self.defaultValues['-PLAYER_T5-'])],
                [sg.Checkbox("Acquire Player Rss assist", key='-PLAYER_RSS-',
                             default=self.defaultValues['-PLAYER_RSS-'])]]
        col4 = [[sg.Checkbox("Acquire Alliance Name", key='-ALLIANCE_NAME-',
                             default=self.defaultValues['-ALLIANCE_NAME-'])],
                [sg.Checkbox("Acquire Player T4 kills", key='-PLAYER_T4-',
                             default=self.defaultValues['-PLAYER_T4-'])],
                [sg.Checkbox("Acquire Player dead", key='-PLAYER_DEAD-',
                             default=self.defaultValues['-PLAYER_DEAD-'])],
                [sg.Text("Players to acquire : 1000", key='-PLAYERS_TO_ACQUIRE_disp-')]]
        col5 = [[sg.Slider(range=(100, 1000), default_value=self.defaultValues['-PLAYERS_TO_ACQUIRE-'], key='-PLAYERS_TO_ACQUIRE-',
                           orientation='vertical', resolution=100, disable_number_display=True, enable_events=True)]]

        row1 = [[sg.Button("Clean screenshots", disabled=True), sg.Button(
            "Delete previous poll", disabled=True), sg.Button("Generate CSV", disabled=True)]]
        row2 = [[sg.Combo(values=self.previous_dates, key='dates', size=(30, None), disabled=True,
                          default_value=self.defaultValues['dates'],), sg.Button('Delete target poll', disabled=True)]]
        row3 = [[sg.Button("Start pull", disabled=True), sg.Button("Reset parameters", key="Reset parameters"),
                 sg.Button("Exit", key="Exit")]]
        self.layout = [[sg.Column(col1), sg.Column(col2)],
                       [sg.Frame("Parameters", [
                                 [sg.Column(col3), sg.Column(col4), sg.Column(col5)]])],
                       [sg.ProgressBar(key='-PROGRESS-', max_value=1000,
                                       visible=False, expand_x=True, size=(50, 10))],
                       [sg.Frame("", [[sg.Column(row1, vertical_alignment='center', justification='center',  k='-C-')],
                                      [sg.Column(row2, vertical_alignment='center', justification='center',  k='-C-')]])],
                       [sg.Column(row3, vertical_alignment='center', justification='center',  k='-C-')]]

    def DirSelected(self, values):
        self.WSDirectory = Path(values['-DIR-'])
        if os.path.isdir(self.WSDirectory):
            self.KDList = []
            for file in os.listdir(self.WSDirectory):
                d = os.path.join(self.WSDirectory, file)
                if os.path.isdir(d):
                    self.KDList.append(file)
            self.window['-KD-'].update(value=self.defaultValues['-KD-'],
                                       values=self.KDList, disabled=False)

    def KDValidated(self, values):
        self.kingdom = values['-KD-'].replace(" ", "_").replace("#", "_").replace("%", "_").replace("&", "_").replace("{", "_").replace("}", "_").replace("\\", "_").replace("$", "_").replace("!", "_").replace("\'", "_").replace(
            "\"", "_").replace(":", "_").replace("@", "_").replace("<", "_").replace(">", "_").replace("*", "_").replace("?", "_").replace("/", "_").replace("+", "_").replace("`", "_").replace("|", "_").replace("=", "_")
        self.kingdomDir = Path(str(self.WSDirectory) + "\\" + self.kingdom)
        if not os.path.exists(self.kingdomDir):
            os.makedirs(self.kingdomDir)
            self.KDList.append(self.kingdom)
            self.window['-KD-'].update(value=self.kingdom, values=self.KDList)
        self.ProjectList = []
        for file in os.listdir(self.kingdomDir):
            d = os.path.join(self.kingdomDir, file)
            if os.path.isdir(d):
                self.ProjectList.append(file)
        self.window['-PROJECT-'].update(value=self.defaultValues['-PROJECT-'],
                                        values=self.ProjectList, disabled=False)

    def ProjectValidated(self, values):
        self.Project = values['-PROJECT-'].replace(" ", "_").replace("#", "_").replace("%", "_").replace("&", "_").replace("{", "_").replace("}", "_").replace("\\", "_").replace("$", "_").replace("!", "_").replace("\'", "_").replace(
            "\"", "_").replace(":", "_").replace("@", "_").replace("<", "_").replace(">", "_").replace("*", "_").replace("?", "_").replace("/", "_").replace("+", "_").replace("`", "_").replace("|", "_").replace("=", "_")
        self.ProjectDir = Path(str(self.kingdomDir) + "\\" + self.Project)
        self.ProjectConfigFile = Path(str(self.ProjectDir) + "\\Config.yaml")
        if not os.path.exists(self.ProjectDir):
            os.makedirs(self.ProjectDir)
            self.ProjectList.append(self.Project)
            self.window['-PROJECT-'].update(value=self.Project,
                                            values=self.ProjectList)
        if "Config.yaml" in os.listdir(self.ProjectDir):
            with open(self.ProjectConfigFile, 'r') as f:
                config = yaml.safe_load(f)
                for param in self.defaultValues.keys():
                    if param not in ['-DIR-', '-KD-', '-PROJECT-']:
                        try:
                            self.window[param].update(
                                value=config[param], disabled=True)
                        except:
                            pass
        table = self.generate_sql_commands()[1]
        connection = sq.connect('rok.db')
        mycursor = connection.cursor()
        mycursor.execute(
            f"SELECT DISTINCT date from {table} order by date ASC")
        dates = [x[0] for x in mycursor]
        self.window['dates'].update(values=dates)
        self.window["Start pull"].update(disabled=False)
        self.window["Clean screenshots"].update(disabled=False)
        self.window["Delete previous poll"].update(disabled=False)
        self.window["dates"].update(disabled=False)
        self.window["Delete target poll"].update(disabled=False)
        self.window["Generate CSV"].update(disabled=False)

    def updateBar(self, inc=1):
        self.progressValue = self.progressValue+inc
        self.window['-PROGRESS-'].update_bar(self.progressValue)

    def Reset(self):

        if os.path.isdir(self.ProjectDir):
            if not any(os.scandir(self.ProjectDir)):
                os.rmdir(self.ProjectDir)

        if os.path.isdir(self.kingdomDir):
            if not any(os.scandir(self.kingdomDir)):
                os.rmdir(self.kingdomDir)

        for param in self.defaultValues.keys():
            if param not in ['-KD-', '-PROJECT-']:
                self.window[param].update(
                    self.defaultValues[param], disabled=False)
            else:
                self.window[param].update(
                    self.defaultValues[param], disabled=True)
        for key in ["Clean screenshots", "Delete previous poll", "dates", "Delete target poll", "Start pull", 'Generate CSV']:
            self.window[key].update(disabled=True)

    def pull(self, values):

        self.acquire_power = values['-PLAYER_POWER-']
        self.acquire_alliance = values['-ALLIANCE_NAME-']
        self.acquire_playerNane = values['-PLAYER_NAME-']
        self.acquire_T4 = values['-PLAYER_T4-']
        self.acquire_T5 = values['-PLAYER_T5-']
        self.acquire_dead = values['-PLAYER_DEAD-']
        self.acquire_rssAssist = values['-PLAYER_RSS-']
        self.players_to_acquire = values['-PLAYERS_TO_ACQUIRE-']

        with open(self.ProjectConfigFile, 'w') as f:
            config = {param: values[param]
                      for param in self.defaultValues.keys()}
            yaml.safe_dump(config, f)

        for param in self.defaultValues.keys():
            try:
                self.window[param].update(disabled=True)
            except:
                pass

        self.ProjectScreenshotsDir = Path(
            str(self.ProjectDir) + "\\ScreenShots")
        if "Config.yaml" not in os.listdir(self.ProjectDir):
            file = open(self.ProjectConfigFile, 'a+')
            file.close()
        if not os.path.exists(self.ProjectScreenshotsDir):
            os.makedirs(self.ProjectScreenshotsDir)
        self.window['-PROGRESS-'].update(visible=True, max=2 *
                                         self.players_to_acquire, current_count=self.progressValue)
        self.configDone.set()
        for value in ['dir select', 'KD_valid', 'Project_valid', 'Start pull', 'Reset parameters', "Clean screenshots", "Delete previous poll", "dates", "Delete target poll", 'Generate CSV']:
            self.window[value].update(disabled=True)
        app = application.Application()
        appli = 'Rise of Kingdoms'
        app.connect(title_re=".*%s.*" % appli)
        app_dialog = app.top_window_()
        app_dialog.Minimize()
        app_dialog.Restore()

    def handle_event(self, event, values):

        if event == '-DIR-':
            self.DirSelected(values)

        if event == 'KD_valid' and not values['-KD-'] == self.defaultValues['-KD-']:
            self.KDValidated(values)

        if event == 'Project_valid' and not values['-PROJECT-'] == self.defaultValues['-PROJECT-']:
            self.ProjectValidated(values)

        if event == 'Reset parameters':
            self.Reset()

        if event == '-PLAYERS_TO_ACQUIRE-':
            self.window['-PLAYERS_TO_ACQUIRE_disp-'].update(
                f"Players to acquire : {int(values['-PLAYERS_TO_ACQUIRE-'])}")

        if event == 'Start pull':
            self.pull(values)

        if event == 'Clean screenshots':
            if os.path.isdir(str(self.ProjectDir) + "\\ScreenShots"):
                for filename in glob.iglob(str(self.ProjectDir) + "\\ScreenShots" + '**/*.png',    recursive=True):
                    os.remove(filename)
                sg.Popup('Sucessfully deleted content', icon=os.path.dirname(
                    __file__)+'\\pyROK.ico', title='Validation')

        if event == 'Delete previous poll':
            self.remove_last_pull(values)

        if event == 'Delete target poll':
            self.remove_target_pull(values)

        if event == 'Generate CSV':
            self.generate_CSV()

        if '_clickReset' in event:
            self.window[event.split('_click')[0]].update('')

    def returnParams(self):
        params = {'ProjectScreenshotsDir': self.ProjectScreenshotsDir,
                  'acquire_playerNane': self.acquire_playerNane,
                  'acquire_alliance': self.acquire_alliance,
                  'acquire_power': self.acquire_power,
                  'acquire_T4': self.acquire_T4,
                  'acquire_T5': self.acquire_T5,
                  'acquire_dead': self.acquire_dead,
                  'acquire_rssAssist': self.acquire_rssAssist,
                  'player_to_acquire': self.players_to_acquire
                  }
        return params

    def pyROK_run(self):

        self.window = sg.Window('pyROK - Stat tool',
                                self.layout, finalize=True)
        self.window.set_icon(icon=os.path.dirname(__file__)+'\\pyROK.ico')
        self.window['-KD-'].bind('<ButtonPress>', '_clickReset')
        self.window['-PROJECT-'].bind('<ButtonPress>', '_clickReset')

        while True:
            event, values = self.window.read()

            if event:
                self.handle_event(event, values)

            if event in (sg.WIN_CLOSED, 'Exit'):
                self.status = False
                break
        self.close.set()
        self.window.close()

    def generate_sql_commands(self, database: str = 'rok.db') -> tuple[str, str, str]:
        file = open(database, 'a+')
        file.close()
        connection = sq.connect(database)
        mycursor = connection.cursor()
        table_name: str = "rok_"+self.kingdom+"_"+self.Project
        # create table if not exist
        mycursor.execute(
            f"CREATE TABLE IF NOT EXISTS {table_name} (_id INTEGER PRIMARY KEY, date VARCHAR(255), player_id INTEGER, name VARCHAR(255), alliance VARCHAR(255), power INTEGER, T4 INTEGER, T5 INTEGER, dead INTEGER, rssAssist INTEGER)")
        mycursor.execute(f"PRAGMA table_info({table_name})")
        connection.commit()
        # create write request
        write_request = f"INSERT INTO {table_name} (`date`, `player_id`,`name`, `alliance`, `power`, `T4`, `T5`, `dead`, `rssAssist`) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
        connection.close()
        return((write_request, table_name))

    def remove_last_pull(self, values):
        table = self.generate_sql_commands()[1]
        connection = sq.connect('rok.db')
        mycursor = connection.cursor()
        mycursor.execute(f"SELECT date from {table} order by date ASC")
        dates = [x[0] for x in mycursor]
        if len(dates) != 0:
            date = dates[-1].replace("{", "").replace("}", "")
            mycursor.execute(f"DELETE FROM {table} where date = \'{date}\'")
            connection.commit()
            sg.Popup('Sucessfully deleted content', icon=os.path.dirname(
                __file__)+'\\pyROK.ico', title='Validation')
        connection.close()

    def remove_target_pull(self, values):
        date = values['dates']
        table = self.generate_sql_commands()[1]
        connection = sq.connect('rok.db')
        mycursor = connection.cursor()
        mycursor.execute(
            f"SELECT DISTINCT date from {table} order by date ASC")
        dates = [x[0] for x in mycursor]
        if date in dates:
            mycursor.execute(f"DELETE FROM {table} where date = \'{date}\'")
            connection.commit()
            sg.Popup('Sucessfully deleted content', icon=os.path.dirname(
                __file__)+'\\pyROK.ico', title='Validation')
        connection.close()

    def generate_CSV(self):
        generate_time = str(
            datetime.datetime.now().strftime("%Y_%m_%d__%H_%M_%S"))
        csv_file = str(self.WSDirectory) + r'\rok_' + self.kingdom + \
            '_' + self.Project + '_' + generate_time + '.csv'
        connection = sq.connect('rok.db')
        default = [-1, -1, -1, -1, -1]
        mycursor = connection.cursor()
        table = self.generate_sql_commands()[1]
        mycursor.execute(
            f"SELECT DISTINCT date from {table} order by date ASC")
        dates = [x[0] for x in mycursor]
        header = ['Alliance', 'Player ID', 'Player Nickname'] + ['Power',
                                                                 'T4 eliminations', 'T5 eliminations', 'dead', 'Rss Assist']*len(dates)
        mycursor.execute(
            f"SELECT DISTINCT player_id from {table} order by player_id ASC")
        players = [x[0] for x in mycursor]
        with open(csv_file, 'a+', newline='', encoding='utf-8') as file:  # create new .csv file
            writer = csv.writer(file)
            writer.writerow(header)
            for player in players:
                current = ['NONE', -1, 'NONE']+default*len(dates)
                mycursor.execute(
                    f'SELECT date, player_id, name, alliance, power, T4, T5, dead, rssAssist FROM {table} WHERE player_id = {player} ORDER BY date ASC')
                player_data = [x for x in mycursor]
                for data in player_data:
                    index = dates.index(data[0])
                    current[index*5+3: (index+1)*5+3] = data[4:]
                if current[3:8] == default:
                    current[3:8] = player_data[0][4:]
                if current[-5:] == default:
                    current[-5:] = player_data[-1][4:]
                current[0] = player_data[0][3]
                current[1] = player
                current[2] = player_data[-1][2]
                writer.writerow(current)
        dataFrame = pd.read_csv(csv_file)
        dataFrame.sort_values(
            ["Power"], axis=0, ascending=False, inplace=True, na_position='first')
        dataFrame.to_csv(csv_file, index=False)
        sg.Popup(f'Sucessfully generated CSV in {self.WSDirectory}', icon=os.path.dirname(
            __file__)+'\\pyROK.ico', title='Validation')

    def done_pulling(self):

        self.window['-PROGRESS-'].update(visible=False, max=2 *
                                         self.players_to_acquire, current_count=self.progressValue)
        self.configDone.set()
        for value in ['dir select', '-KD-', '-PROJECT-', 'KD_valid', 'Project_valid', 'Start pull', 'Reset parameters', "Clean screenshots", "Delete previous poll", "dates", "Delete target poll", 'Generate CSV']:
            self.window[value].update(disabled=False)


if __name__ == '__main__':
    pyROK = pyROK_GUI()
    pyROK.pyROK_run()
