import io
import shutil
import stat
import subprocess
import sys
import time
from datetime import datetime, timedelta
from enum import IntEnum
import os.path
from typing import Literal
import zipfile

import PySimpleGUI as sg
import json
import re
import pyperclip
from configparser import ConfigParser
import requests
import random

from gmailAPI import GmailAPI

class PhraseType(IntEnum):
    INVOICE = 0
    SUBJECT = 1
    DATES = 2

class InvoiceApp:
    # Get theme List
    with open('theme_list.txt', 'r') as file:
        content = file.read()
        
    theme_list = content.split()
    print(theme_list)

    # Default values
    DEFAULT_THEME = 'LightGreen2'
    KEEP_ON_TOP = True
    OUTSIDE_TERM_TIME_MSG = '<OUTSIDE TERM TIME!>'
    STARTING_WINDOW_X = 585
    STARTING_WINDOW_Y = 427
    APP_URL = 'https://github.com/WolfyCodeK/student-invoice-template-app/raw/main/StudentInvoiceExecutable.zip'
    APP_VERSION_URL = 'https://raw.githubusercontent.com/WolfyCodeK/student-invoice-template-app/main/app_version'

    # Term Dates
    today = datetime.now()
    current_date = today

    # Debug code for when out of term time
    # year = int(sg.popup_get_text('YEAR', size= 10, keep_on_top=KEEP_ON_TOP))
    # month = int(sg.popup_get_text('MONTH', size= 10, keep_on_top=KEEP_ON_TOP))
    # day = int(sg.popup_get_text('DAY', size= 10, keep_on_top=KEEP_ON_TOP))
    # currentDate = datetime(year, month, day)

    autumn_first_half = [datetime(2023, 9, 4), datetime(2023, 10, 21), '1st', 'autumn']
    autumn_second_half = [datetime(2023, 10, 30), datetime(2023, 12, 23), '2nd', 'autumn']

    spring_first_half = [datetime(2024, 1, 8), datetime(2024, 2, 10), '1st', 'spring']
    spring_second_half  = [datetime(2024, 2, 19), datetime(2024, 3, 30), '2nd', 'spring']

    summer_first_half = [datetime(2024, 4, 15), datetime(2024, 5, 25), '1st', 'summer']
    summer_second_half = [datetime(2024, 5, 3), datetime(2024, 6, 24), '2nd', 'summer']

    term_list = [autumn_first_half, autumn_second_half, spring_first_half, spring_second_half, summer_first_half, summer_second_half]

    weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November','December']

    # Custom Fonts
    text_font = ('Lucida Console', 13)
    
    # Colours
    READONLY_BACKGROUND_COLOUR = '#FF6961'

    # File Paths
    TEMPLATES_PATH = 'templates.json'
    CONFIG_PATH = 'config.ini'

    # Config Sections
    PREFERENCES_SECTION = 'Preferences'
    STATE_SECTION = 'State'
    META_DATA_SECTION = 'Meta-Data'
    
    # Config Values
    THEME = 'Theme'
    EMAIL_MODE = 'email-mode'
    CURRENT_TEMPLATE = 'current-template'
    LAST_WINDOW_X = 'last-win-x'
    LAST_WINDOW_Y = 'last-win-y'
    APP_VERSION_TITLE = 'app-version'
    APP_LATEST_AVAILABLE_VERSION_TITLE = 'latest-available-version'

    # Resource Paths
    RESOURCE_DIR = 'res/'
    BLANK_ICO = 'res/Blank.ico'
    MAIN_ICO = 'res/Email.ico'
    
    # Window Titles
    EDIT_TEMPLATE_TITLE = 'Edit Template'

    # Window Sizes
    MAIN_WIDTH = 750
    MAIN_HEIGHT = 225
    SELECT_WIDTH = 450
    SELECT_HEIGHT = 350
    SETTINGS_WIDTH = 600
    SETTINGS_HEIGHT = 200

    WIN_OFFSET_X = 550 / 2
    WIN_OFFSET_Y = 225 / 2

    # Element Sizes
    MAIN_PADDING = 15
    SELECT_PADDING = 10
    SETTINGS_PADDING = 10

    # Button Values
    EDIT_BUTTON = 'Edit'
    EXIT_BUTTON = 'Exit'
    NEW_TEMPLATE_BUTTON = 'New Template'
    SAVE_BUTTON = 'Save & Close'
    DELETE_BUTTON = 'DELETE'
    DRAFT_BUTTON = 'Draft'
    DRAFT_ALL_BUTTON = 'Draft All'
    BODY_BUTTON = 'Body'
    SUBJECT_BUTTON = 'Subject'
    SETTINGS_BUTTON = 'Settings'
    UPDATE_BUTTON = 'Update'
    
    template_edit_buttons_list = [DELETE_BUTTON, DRAFT_BUTTON, DRAFT_ALL_BUTTON, BODY_BUTTON, SUBJECT_BUTTON, EDIT_BUTTON]

    # Email Modes
    CLIPBOARD = 'Clipboard'
    AUTO_DRAFT = 'Auto Draft'

    # Combo Values
    NAMES_COMBOBOX = 'Names'
    THEME_COMBOBOX = '<Theme>'

    # Template Name
    RECIPIENT_INPUT = 'Recipient'
    
    # Template Info
    COST_INPUT = 'Cost'
    INSTRUMENT_INPUT = 'Instrument'
    DAY_INPUT = 'Day'
    STUDENT_INPUT = 'Students'
    INFO_VALUES_LIST = [COST_INPUT, INSTRUMENT_INPUT, DAY_INPUT, STUDENT_INPUT]
    
    INPUT_SIZE = 15
    THEME_INPUT_SIZE = 21

    SINGLE_DRAFT_THREAD_END_KEY = '-SINGLE DRAFT THREAD-'
    
    ALL_DRAFT_THREAD_END_KEY = '-ALL DRAFT THREAD-'

    instruments_list = ['piano', 'drum', 'guitar', 'vocal', 'music', 'singing', 'bass guitar', 'classical guitar']

    def __init__(self):
        # Get the absolute path of this file
        self.current_file_path = os.path.abspath(__file__)

        self.parent_path = os.path.dirname(self.current_file_path)

        # Navigate two directories back to get the top level path
        self.top_level_path = os.path.dirname(self.parent_path)
        
        # Get app version
        app_version_path = f'{self.parent_path}\\lib\\app_version'

        try:
            # Read the content and remove leading/trailing whitespace
            with open(app_version_path, 'r') as f:
                app_version = f.read().strip()  
                f.close()

        except FileNotFoundError:
            InvoiceApp.log_message('ERROR', f'File "{app_version_path}" not found')
            sys.exit()
        
        # Create resources directory if it does not exist
        if not os.path.exists(self.RESOURCE_DIR):
            os.makedirs(self.RESOURCE_DIR)
        
        # Create templates file if it does not exist
        if not os.path.isfile(self.TEMPLATES_PATH):
            with open(self.TEMPLATES_PATH, 'w') as f:
                f.write('{}')
                f.close()          
    
        # Create config file if it does not exist
        if not os.path.isfile(self.CONFIG_PATH):
            with open(self.CONFIG_PATH, 'w') as f:
                f.write(f'[{self.PREFERENCES_SECTION}]\n')
                f.write(f'{self.THEME} = ' + self.DEFAULT_THEME + '\n')
                f.write(f'{self.EMAIL_MODE} = {self.CLIPBOARD}\n')
            
                f.write(f'\n[{self.STATE_SECTION}]\n')
                f.write(f'{self.CURRENT_TEMPLATE} = \n')
                f.write(f'{self.LAST_WINDOW_X} = {self.STARTING_WINDOW_X}\n')
                f.write(f'{self.LAST_WINDOW_Y} = {self.STARTING_WINDOW_Y}\n')
                
                f.write(f'\n[{self.META_DATA_SECTION}]\n')
                f.write(f'{self.APP_VERSION_TITLE} = {app_version}\n')
                f.write(f'{self.APP_LATEST_AVAILABLE_VERSION_TITLE} = \n')
                f.close()
            
        # Create config parser
        self.config = ConfigParser()
        self.config.read(self.CONFIG_PATH)       
        
        # Fetch latest version and set number in config
        self.config.set(self.META_DATA_SECTION, self.APP_LATEST_AVAILABLE_VERSION_TITLE, self.fetch_latest_available_version_number())     

        self.config.set(self.META_DATA_SECTION, self.APP_VERSION_TITLE, self.get_current_app_version())     
        
        self.save_config()   
            
        sg.theme(self.get_theme())
        
        # Create gmail API object
        self.gmail_API = GmailAPI()
        
    @staticmethod
    def isFloat(value):
        try:
            float(value)
            return True
        except ValueError:
            return False
    
    @staticmethod
    def log_message(log_level: Literal['ERROR', 'WARNING', 'INFO'], log_msg: str):
        with open('error_log.txt', 'a') as f:
            f.write(f'{datetime.now().strftime("%d-%m-%Y %H:%M:%S")}: [{log_level}] - {log_msg}\n')
            f.close()
            
    @staticmethod
    def is_newer_version_available(current_version, new_version):
        v1_parts = list(map(int, current_version.split('.')))
        v2_parts = list(map(int, new_version.split('.')))

        for i in range(len(v1_parts)):
            if v2_parts[i] > v1_parts[i]:
                return True
            elif v2_parts[i] < v1_parts[i]:
                return False

        return False
    
    def run(self):
        repaired = False
        
        # Templates file data validation
        with open(self.TEMPLATES_PATH, 'r') as fr:
            json_data = dict(json.load(fr))
            
            name_list = list(json_data)
            new_json_data = dict()
            
            needs_repairing = False
            
            # Check if templates file needs repairing
            for name in name_list:
                new_json_data[name] = {}
        
                name_data = dict(json_data[name])
                
                if sorted(list(name_data.keys())) != sorted(self.INFO_VALUES_LIST):
                    needs_repairing = True
            
            if needs_repairing:
                if self.display_message_box("It seems that the information in your templates might be outdated or damaged. To fix this issue, would you like to try repairing the data?", 'yn', None) == 'Yes':
                    
                    for name in name_list:
                        new_json_data[name] = {}
            
                        name_data = dict(json_data[name])
                        
                        # Check if there's the correct number of keys for the first name in the template
                        if sorted(list(name_data.keys())) != sorted(self.INFO_VALUES_LIST):    
                        
                            # Cost
                            if name_data.get(self.COST_INPUT) == None:
                                for info in name_data:
                                    if InvoiceApp.isFloat(name_data[info]):
                                        new_json_data[name][self.COST_INPUT] = json_data[name][info]  
                            else:
                                new_json_data[name][self.COST_INPUT] = json_data[name][self.COST_INPUT]
                                        
                            # Instrument
                            if name_data.get(self.INSTRUMENT_INPUT) == None:
                                for info in name_data:
                                    if name_data[info] in self.instruments_list:
                                        new_json_data[name][self.INSTRUMENT_INPUT] = json_data[name][info]
                            else:
                                new_json_data[name][self.INSTRUMENT_INPUT] = json_data[name][self.INSTRUMENT_INPUT]
                                        
                            # Day
                            if name_data.get(self.DAY_INPUT) == None:
                                for info in name_data:
                                    if name_data[info] in self.weekdays:
                                        new_json_data[name][self.DAY_INPUT] = json_data[name][info]
                            else:
                                new_json_data[name][self.DAY_INPUT] = json_data[name][self.DAY_INPUT]
                                        
                            # Students
                            if name_data.get(self.STUDENT_INPUT) == None:
                                # Know Student key alternatives
                                new_json_data[name][self.STUDENT_INPUT] = json_data[name]['Student']
                            else:
                                new_json_data[name][self.STUDENT_INPUT] = json_data[name][self.STUDENT_INPUT]
                        else:
                            new_json_data[name] = json_data[name]
                            
                    with open(self.TEMPLATES_PATH, 'w') as fw:
                            fw.write(json.dumps(new_json_data, sort_keys=True))   
                            fw.close()
                            
                    repaired = True
                        
                else:
                    self.display_message_box("Cannot proceed with invalid template data\nApp is terminating!", 'er', None)
                        
                    time.sleep(3)
                    sys.exit()

        self.main_window(repaired)
                        
    def fetch_latest_available_version_number(self):
        latest_available_version = self.get_current_app_version()
        
        try:
            response = requests.get(self.APP_VERSION_URL)
                
            latest_available_version = response.content.decode().replace('\n', '')
            InvoiceApp.log_message('INFO', f'Latest version tag fetched from github -> {latest_available_version}')
                
        except Exception as e:
            self.display_message_box('There was a problem fetching the latest update. Try checking your internet connection.' + e, 'er', None)
            
        return latest_available_version

    def get_theme(self):
        return self.config.get(self.PREFERENCES_SECTION, self.THEME)
        
    def get_email_mode(self):
        return self.config.get(self.PREFERENCES_SECTION, self.EMAIL_MODE)
    
    def get_current_template(self):
        return self.config.get(self.STATE_SECTION, self.CURRENT_TEMPLATE)  
    
    def get_current_app_version(self):
        return self.config.get(self.META_DATA_SECTION, self.APP_VERSION_TITLE)
    
    def get_latest_available_app_version(self):
        return self.config.get(self.META_DATA_SECTION, self.APP_LATEST_AVAILABLE_VERSION_TITLE)
        
    def get_phrases(self, start_date: datetime, end_date: datetime, half: str, term: str):
            
        body_phrase = str(self.num_to_weekday(start_date.isoweekday())) + ' ' + str(start_date.day) + self.get_day_suffix(start_date.day) + str(self.num_to_month(start_date.month)) + ' to and including ' + str(self.num_to_weekday(end_date.isoweekday())) + ' ' + str(end_date.day) + self.get_day_suffix(end_date.day) + str(self.num_to_month(end_date.month))
        
        current_term = [
            half + ' half ' + term + ' term ' + str(start_date.year), 
            half + ' Half ' + term.capitalize() + ' Term ' + str(end_date.year), body_phrase
        ]
        
        return current_term

    def get_day_suffix(self, day: int):
        match(day):
            case 1:
                return 'st '
            case 2:
                return 'nd '
            case 3:
                return 'rd '
            case _:
                return 'th '

    def get_term_length_in_weeks(self, start_date, end_date):
        first_monday = (start_date - timedelta(days=start_date.weekday()))
        last_monday = (end_date - timedelta(days=end_date.weekday()))   
        
        return timedelta(weeks=(((last_monday - first_monday).days / 7) + 1))

    def which_term(self, date: datetime, day: str):
        current_term = [self.OUTSIDE_TERM_TIME_MSG, self.OUTSIDE_TERM_TIME_MSG, self.OUTSIDE_TERM_TIME_MSG] 

        for term in self.term_list:
            if (date >= term[0] and date <= term[1]):
                start_date = self.next_day_in_week(term[0], day)
                date_gap = self.get_term_length_in_weeks(term[0], term[1])
                end_date = start_date + date_gap
        
                current_term = self.get_phrases(start_date, end_date - timedelta(weeks=1), term[2], term[3])
                
                return current_term, int(date_gap.days / 7)
            
        return None 
            
    def check_select_fields_are_not_empty(self, values: dict):
        return len(str(values[self.RECIPIENT_INPUT])) == 0 or len(str(values[self.COST_INPUT])) == 0 or len(str(values[self.INSTRUMENT_INPUT])) == 0 or len(str(values[self.STUDENT_INPUT])) == 0

    def num_to_weekday(self, num: int):
        return self.weekdays[num-1]

    def num_to_month(self, num: int):
        return self.months[num-1]

    def next_day_in_week(self, date: datetime, target_day: str):    
        day_found = False
        search_day = date
        
        while day_found == False:
            if (self.num_to_weekday(search_day.isoweekday()) == target_day):
                day_found = True
            else:
                search_day = search_day + timedelta(days=1)
            
        return search_day

    def toggle_buttons_disabled(self, window: sg.Window, button_list: str, disabled: bool):
        for button in button_list:
            window[button].update(disabled=disabled)

    def get_names_list(self):
        with open(self.TEMPLATES_PATH, 'r') as f:
            json_data = json.load(f)
            names_list = list(json_data.keys())
            names_list = sorted(names_list, key=str.lower)
            f.close()
            
        return names_list

    def get_subject(self, name: str):
        with open(self.TEMPLATES_PATH, 'r') as f:
            json_data = json.load(f)
        
        day = str(json_data[name][self.DAY_INPUT])
        instrument = str(json_data[name][self.INSTRUMENT_INPUT])
        term_data = self.which_term(self.current_date, day)
        
        if term_data is not None:
            phrases, _ = term_data
            return f"Invoice for {instrument.title()} Lessons {phrases[PhraseType.SUBJECT]}"
        
        return f"Invoice for <?> Lessons <?>"
        
    def get_body(self, name: str):
        with open(self.TEMPLATES_PATH, 'r') as f:
            json_data = json.load(f)
            f.close()
        
        cost_of_lessons = json_data[name][self.COST_INPUT]
        instrument = json_data[name][self.INSTRUMENT_INPUT]
        day = json_data[name][self.DAY_INPUT]
        students = json_data[name][self.STUDENT_INPUT]
        
        term_data = self.which_term(self.current_date, day)
        
        if term_data is not None:
            phrases, num_of_lessons = term_data
            
            total_cost = '%.2f' % (round(float(int(num_of_lessons) * float(cost_of_lessons)), 2))
    
            num_of_lessons_phrase = f"There are {num_of_lessons} sessions"
            
            if int(num_of_lessons) == 1:
                num_of_lessons_phrase = f"There is {num_of_lessons} session"
        
            return f"Hi {name},\n\nHere is my invoice for {students}'s {instrument} lessons {phrases[PhraseType.INVOICE]}.\n--------\n{num_of_lessons_phrase} this {phrases[PhraseType.INVOICE]} from {phrases[PhraseType.DATES]}.\n\n{num_of_lessons} x £{cost_of_lessons} = £{total_cost}\n\nThank you\n--------\n\nKind regards\nRobert"
        
        eMsg = '<?>'
        
        return f">>>>> WARNING: OUTSIDE TERM TIME\n>>>>> INFORMATION CANNOT BE APPLIED\n\nHi {name},\n\nHere is my invoice for {eMsg}'s {eMsg} lessons {eMsg}.\n--------\n{eMsg} this {eMsg} from {eMsg}.\n\n{eMsg} x £{eMsg} = £{eMsg}\n\nThank you\n--------\n\nKind regards\nRobert"

    def display_message_box(self, title: str, box_type: Literal['yn', 'qm', 'pu', 'er'], window: sg.Window, close_duration: int = 2):
        """Display a message to the user inside a box.

        Args:
            title (str): The text displayed inside the pop up message.
            box_type (Literal['yn', 'qm', 'pu', 'er']): [ys -> Yes or No], [qm -> Quick Message], [pu -> Pop Up], [er -> Error Message]

        Returns:
            (str | None): The string of the button clicked or None if no button was clicked.
        """
        # If possible save window location before displaying message box
        if window is not None:
            self.save_win_location(window)
        
        if box_type == 'yn':
            choice = sg.popup_yes_no(title, title='', font=self.text_font, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, location=(self.get_last_win_x() + self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y))
    
            return choice
        
        elif box_type == 'qm':
            sg.popup_quick_message(title, font=self.text_font, title='', icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, background_color="Black", text_color="White", location=(self.get_last_win_x() + self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y), auto_close_duration=close_duration) 
            
            return None
        
        elif box_type == 'pu':
            choice = sg.popup(title, title='', font=self.text_font, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, location=(self.get_last_win_x() + self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y))
            
            return choice
        elif box_type == 'er':
            sg.popup(title, font=self.text_font, title='', icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, background_color="Red", text_color="White", location=(self.get_last_win_x() + self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y)) 
            
            return None
        
    def toggle_clipboard_visible(self, window: sg.Window, clipboard_visible: bool):
            window[self.DRAFT_BUTTON].update(visible=not clipboard_visible)
            window[self.DRAFT_ALL_BUTTON].update(visible=not clipboard_visible)
            window[self.SUBJECT_BUTTON].update(visible=clipboard_visible)
            window[self.BODY_BUTTON].update(visible=clipboard_visible)
            
    def save_config(self):
        with open(self.CONFIG_PATH, 'w') as config_file:
            self.config.write(config_file)
            config_file.close()       
            
    def create_draft_for_template(self, name: str):
        self.gmail_API.gmail_create_draft(self.get_subject(name), self.get_body(name))
        
    def get_last_win_x(self):
        return int(self.config.get(self.STATE_SECTION, self.LAST_WINDOW_X))
    
    def get_last_win_y(self):
        return int(self.config.get(self.STATE_SECTION, self.LAST_WINDOW_Y))
    
    def save_win_location(self, window: sg.Window):
        x, y = window.CurrentLocation()
                
        self.config.set(self.STATE_SECTION, self.LAST_WINDOW_X, str(x)) 
        self.config.set(self.STATE_SECTION, self.LAST_WINDOW_Y, str(y))
        
        self.save_config()
    
    def settings_window(self):    
        """Create the settings window.

        Returns:
            bool: True if user saves the settings.
        """
        saved = False
            
        layout = [
                    [
                        sg.Text(self.THEME_COMBOBOX, font=self.text_font, pad=self.SETTINGS_PADDING), 
                        sg.Combo(self.theme_list, font=self.text_font, size=self.THEME_INPUT_SIZE, readonly=True, key=self.THEME_COMBOBOX, default_value=self.get_theme()),
                        sg.Button('Randomise', font=self.text_font, pad=self.SETTINGS_PADDING)
                    ],
                    [
                        sg.Text('<Email Mode>', font=self.text_font, pad=self.SETTINGS_PADDING), 
                        sg.Combo([self.CLIPBOARD, self.AUTO_DRAFT], font=self.text_font, size=self.THEME_INPUT_SIZE, readonly=True, key='Email Mode', default_value=self.get_email_mode()),
                    ],
                    [
                        sg.VPush()
                    ],
                    [
                        sg.Button(self.SAVE_BUTTON, font=self.text_font),
                        sg.Push(), 
                        sg.Button(self.EXIT_BUTTON, font=self.text_font)
                    ]
            ]
        
        window = sg.Window(self.SETTINGS_BUTTON, layout, element_justification='l', size=(self.SETTINGS_WIDTH, self.SETTINGS_HEIGHT), modal=True, icon=self.MAIN_ICO, keep_on_top=self.KEEP_ON_TOP, location=(self.get_last_win_x() + self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y))
        
        # Make initial window read
        window.read(timeout=1)
        
        # Event Loop
        while True:
            event, values = window.read()
            
            if event == self.EXIT_BUTTON or event == sg.WIN_CLOSED:
                break
            if event == 'Randomise':
                window[self.THEME_COMBOBOX].update(value=random.choice(self.theme_list))
                
            if event == self.SAVE_BUTTON:
                self.config.set(self.PREFERENCES_SECTION, self.THEME, values[self.THEME_COMBOBOX])
                self.config.set(self.PREFERENCES_SECTION, self.EMAIL_MODE, values['Email Mode'])
                
                self.save_config()
                
                saved = True
                
                break
            
        window.close()
        
        return saved

    def selected_template_window(self, is_new_template: bool, name: str = ''):
        with open(self.TEMPLATES_PATH, 'r') as f:
            json_data = json.load(f)
            f.close()
        
        if is_new_template:
            recipient_default = ''
            recipient_disabled = False
            cost_default = ''
            instrument_default = ''
            day_default = ''
            student_default = ''
        else:
            recipient_default = name
            recipient_disabled = True
            cost_default = json_data[name][self.COST_INPUT]   
            instrument_default = json_data[name][self.INSTRUMENT_INPUT]
            day_default = json_data[name][self.DAY_INPUT]
            student_default = json_data[name][self.STUDENT_INPUT] 
        
        layout = [
                    [
                        sg.Text('Recipient', font=self.text_font, pad=self.SELECT_PADDING), 
                        sg.Input(size=self.INPUT_SIZE*2, font=self.text_font, key=self.RECIPIENT_INPUT, default_text=recipient_default, disabled=recipient_disabled, disabled_readonly_background_color=self.READONLY_BACKGROUND_COLOUR)
                    ],
                    [
                        sg.Text('Cost of lesson  £', font=self.text_font, pad=self.SELECT_PADDING), 
                        sg.Input(size=self.INPUT_SIZE, font=self.text_font, key=self.COST_INPUT, default_text=cost_default)
                    ],
                    [   sg.Text('Instrument', font=self.text_font, pad=self.SELECT_PADDING), 
                        sg.Combo(values=sorted(self.instruments_list), size=self.INPUT_SIZE*2,
                        font=self.text_font, key=self.INSTRUMENT_INPUT, default_value=instrument_default, readonly=True)
                    ],
                    [
                        sg.Text('Start Day', font=self.text_font, pad=self.SELECT_PADDING),
                        sg.Combo(values=self.weekdays, size=self.INPUT_SIZE, font=self.text_font, key=self.DAY_INPUT, default_value=day_default, readonly=True)
                    ],
                    [   
                        sg.Text('Student(s)', font=self.text_font)
                    ], 
                    [
                        sg.Multiline(size=(self.INPUT_SIZE*2, 2), font=self.text_font, key=self.STUDENT_INPUT, default_text=student_default)
                    ],
                    [
                        sg.VPush()
                    ],
                    [
                        sg.Button(self.SAVE_BUTTON, font=self.text_font), 
                        sg.Push(), 
                        sg.Button(self.EXIT_BUTTON, font=self.text_font)
                    ]
                ]
        
        if is_new_template:
            template_title = self.NEW_TEMPLATE_BUTTON
        else:
            template_title = self.EDIT_TEMPLATE_TITLE
        
        window = sg.Window(template_title, layout, element_justification='l', size=(self.SELECT_WIDTH, self.SELECT_HEIGHT), modal=True, icon=self.MAIN_ICO, keep_on_top=self.KEEP_ON_TOP, location=(self.get_last_win_x()+ self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y))
        
        # Make initial window read
        window.read(timeout=1)
        
        # Event Loop
        while True:
            event, values = window.read()
            
            if event == self.EXIT_BUTTON or event == sg.WIN_CLOSED:
                break
            if event == self.SAVE_BUTTON: 
                # Input error checking
                if re.search('\d', values[self.RECIPIENT_INPUT]):
                    self.display_message_box('Recipient name cannot contain numbers!', 'pu', window)
                    
                elif re.search('\D', str(values[self.COST_INPUT]).replace('.', '')):
                    self.display_message_box('Cost of lesson cannot contain characters!', 'pu', window)
                    
                elif re.search('\d', values[self.STUDENT_INPUT]):
                    self.display_message_box('Students names cannot contain numbers!', 'pu', window)
                    
                elif self.check_select_fields_are_not_empty(values):
                    self.display_message_box('All fields must be completed!', 'pu', window)
                    
                else:
                    if (name in json_data) and is_new_template:
                        self.display_message_box('Template with that name already exists!', 'pu', window)
                        
                    else:
                        name = values[self.RECIPIENT_INPUT]
                        
                        info = {
                            self.COST_INPUT : '%.2f' % (round(float(values[self.COST_INPUT]), 2)),
                            self.INSTRUMENT_INPUT : values[self.INSTRUMENT_INPUT],
                            self.DAY_INPUT: values[self.DAY_INPUT],
                            self.STUDENT_INPUT : values[self.STUDENT_INPUT]
                        }
                        
                        json_data[name] = info

                        with open(self.TEMPLATES_PATH, 'w') as f:
                            f.write(json.dumps(json_data, sort_keys=True))
                            f.close()
                        break        
        window.close()
        
        return name

    def get_main_window(self):
        if (self.get_email_mode() == self.CLIPBOARD):
            sub_body_enabled = True
            draft_enabled = False
        else:
            draft_enabled = True
            sub_body_enabled = False
            
        current_template = self.get_current_template()
        
        # Check if no template is currentely selected
        if current_template == "":
            support_buttons_disabled = True
        else:
            support_buttons_disabled = False
            
        if InvoiceApp.is_newer_version_available(self.get_current_app_version(), self.get_latest_available_app_version()):
            update_tooltip = f' Latest update v{self.get_latest_available_app_version()} is available '
            tooltip_disabled = False
        else:
            update_tooltip = None
            tooltip_disabled = True
        
        support_buttons = [
                            [
                                sg.Button(self.DELETE_BUTTON, font=self.text_font, disabled=support_buttons_disabled),
                                sg.Push(),
                                sg.Button(self.NEW_TEMPLATE_BUTTON, font=self.text_font), 
                                sg.Button(self.SETTINGS_BUTTON, font=self.text_font),
                                sg.Button(self.EXIT_BUTTON, font=self.text_font)
                            ]       
                        ]

        layout = [  
                    [
                        sg.Text(f'Build: v{self.get_current_app_version()}'),
                        sg.Push(),
                        sg.Button(self.UPDATE_BUTTON, font=self.text_font, disabled=tooltip_disabled, tooltip=update_tooltip)
                    ],
                    [
                        [
                            sg.Text('<Templates>', font=self.text_font),
                            sg.Combo(self.get_names_list(), enable_events=True, default_value=current_template, pad=self.MAIN_PADDING, key=self.NAMES_COMBOBOX, size=15, font=self.text_font, readonly=True),
                            sg.Button(self.EDIT_BUTTON, font=self.text_font, disabled=support_buttons_disabled),
                            sg.Button(self.DRAFT_BUTTON, font=self.text_font, disabled=support_buttons_disabled, visible=draft_enabled),
                            sg.Button(self.DRAFT_ALL_BUTTON, font=self.text_font, disabled=support_buttons_disabled, visible=draft_enabled),
                            sg.Button(self.SUBJECT_BUTTON, font=self.text_font, disabled=support_buttons_disabled, visible=sub_body_enabled),
                            sg.Button(self.BODY_BUTTON, font=self.text_font, disabled=support_buttons_disabled, visible=sub_body_enabled)
                        ]
                    ],
                    [
                        sg.VPush()
                    ],
                    [
                        sg.Column(support_buttons, element_justification='right',expand_x=True)
                    ]
                ]

        window = sg.Window('Invoice Templates', layout, element_justification='l', size=(self.MAIN_WIDTH, self.MAIN_HEIGHT), icon=self.MAIN_ICO, keep_on_top=self.KEEP_ON_TOP, location=(self.get_last_win_x(), self.get_last_win_y()), enable_close_attempted_event=True)
        
        # Make initial window read
        window.read(timeout=1)
        
        return window

    def main_window(self, repaired: bool = False):        
        window = self.get_main_window()  
        
        if repaired:
            self.display_message_box("Template data repaired successfully!", 'qm', window)
        
        # Event Loop
        while True:
            event, values = window.read()
            
            if event == self.EXIT_BUTTON or event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT:
                self.config.set(self.STATE_SECTION, self.CURRENT_TEMPLATE, str(values[self.NAMES_COMBOBOX]))
                
                self.save_win_location(window)
                
                self.save_config()
                
                break

            if event == self.NAMES_COMBOBOX:
                self.toggle_buttons_disabled(window, self.template_edit_buttons_list, False)        
                    
            if event == self.DELETE_BUTTON:                
                if self.display_message_box('Are you sure you want to delete this template?', 'yn', window) == 'Yes':
                    with open(self.TEMPLATES_PATH, 'r') as f:
                        json_data = json.load(f)
                        f.close()
                        
                    del json_data[values[self.NAMES_COMBOBOX]]   
                    
                    with open(self.TEMPLATES_PATH, 'w') as f:
                        f.write(json.dumps(json_data))   
                        f.close()   
                    
                    window[self.NAMES_COMBOBOX].update(values=self.get_names_list())                  
                        
                    self.toggle_buttons_disabled(window, self.template_edit_buttons_list, True)
                    
                    self.display_message_box('Template Deleted!', 'qm', window)
                
            if event == self.EDIT_BUTTON:
                self.save_win_location(window)
                
                self.selected_template_window(False, values[self.NAMES_COMBOBOX])
                
            if event == self.NEW_TEMPLATE_BUTTON:
                self.save_win_location(window)
                
                name = self.selected_template_window(True)
                
                if name != '':
                    window[self.NAMES_COMBOBOX].update(value=name, values=self.get_names_list())
                    
                    self.toggle_buttons_disabled(window, self.template_edit_buttons_list, False)
            
            if event == self.DRAFT_BUTTON:
                self.display_message_box('Sending...', 'qm', window)
                
                window.start_thread(lambda: self.create_draft_for_template(values[self.NAMES_COMBOBOX]), self.SINGLE_DRAFT_THREAD_END_KEY)
                
            if event == self.DRAFT_ALL_BUTTON:                
                if self.display_message_box('Are you sure you want to draft all templates?', 'yn', window) == 'Yes':

                    self.display_message_box('Sending...', 'qm', window)

                    for name in self.get_names_list():
                        window.start_thread(lambda: self.create_draft_for_template(name), self.ALL_DRAFT_THREAD_END_KEY)
                
            if event == self.SINGLE_DRAFT_THREAD_END_KEY:
                self.display_message_box('Draft Sent!', 'qm', window)
                
            if event == self.ALL_DRAFT_THREAD_END_KEY:
                self.display_message_box('All Drafts Sent!', 'qm', window)
                
            if event == self.SUBJECT_BUTTON:
                pyperclip.copy(self.get_subject(values[self.NAMES_COMBOBOX]))
                
                self.display_message_box('Copied Subject!', 'qm', window)
                
            if event == self.BODY_BUTTON:
                pyperclip.copy(self.get_body(values[self.NAMES_COMBOBOX]))
                
                self.display_message_box('Copied Body!', 'qm', window) 

            if event == self.SETTINGS_BUTTON:
                self.save_win_location(window)
                
                # Reload window if settings were saved
                if self.settings_window():
                    self.config.set(self.STATE_SECTION, self.CURRENT_TEMPLATE, str(values[self.NAMES_COMBOBOX]))
                
                    self.save_config()
                    
                    sg.theme(self.get_theme())
                    window.close()
                    window = self.get_main_window()
                    
                    self.display_message_box('Settings Saved!', 'qm', window)
                    
            if event == self.UPDATE_BUTTON:
                if self.display_message_box(f'Are you sure you want to update to version v{self.get_latest_available_app_version()}', 'yn', window) == 'Yes':
                    self.display_message_box('Updating...', 'qm', window, 5)
                    
                    # Attempt to get response from download url
                    response = requests.get(self.APP_URL, stream=True)
                    
                    if response.status_code == 200:
                        download_content = io.BytesIO(response.content)
                    else:
                        self.log_message('ERROR', f'Failed to download the file. Status code: {response.status_code}')
                        self.display_message_box('Failed to update!', 'er', window)
                        
                    with zipfile.ZipFile(download_content, 'r') as zip_ref:
                        zip_ref.extractall(self.top_level_path)
                    
                    latest_app_name = f'StudentInvoice-{self.get_latest_available_app_version()}'
                    latest_app_directory_path = self.top_level_path + f'\\{latest_app_name}'
                    
                    # Copy over credential files and any relevant encryption files
                    if os.path.exists(f'{self.parent_path}\\lib\\credentials.json') and os.path.exists(f'{self.parent_path}\\lib\\key.key'):
                        shutil.copy(
                            f'{self.parent_path}\\lib\\credentials.json',
                            f'{latest_app_directory_path}\\lib'
                        )
                        
                        shutil.copy(
                            f'{self.parent_path}\\lib\\key.key',
                            f'{latest_app_directory_path}\\lib'
                        )
                        
                        if os.path.exists(f'{self.parent_path}\\lib\\token.json'):
                            shutil.copy(
                                f'{self.parent_path}\\lib\\token.json',
                                f'{latest_app_directory_path}\\lib'
                            )

                    # Copy over templates file to new version
                    if os.path.exists(f'{self.parent_path}\\templates.json'):
                        shutil.copy(
                            f'{self.parent_path}\\templates.json',
                            latest_app_directory_path
                        )
                        
                    # Make this directory deletable for when new app is installed
                    os.chmod(self.parent_path, stat.S_IWRITE)      

                    # Set directory to new app version
                    os.chdir(latest_app_directory_path)
                    
                    self.log_message('INFO', f'Launching {latest_app_name}.exe')
                    
                    # Launch new app version
                    subprocess.run([f'{self.parent_path}\\lib\\launch_executable.bat', f'{latest_app_directory_path}\\{latest_app_name}.exe', self.get_current_app_version()])

                    sys.exit()
                
        window.close()