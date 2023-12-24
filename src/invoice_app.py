from math import floor
import random
import sys
import time
import PySimpleGUI as sg
import json
import re
import pyperclip
from datetime import datetime, timedelta
from enum import IntEnum
from gmailAPI import GmailAPI
import os.path
from configparser import ConfigParser

class PhraseType(IntEnum):
    INVOICE = 0
    SUBJECT = 1
    DATES = 2

class InvoiceApp:
    # Theme List
    theme_list = ['Black', 'BlueMono', 'BluePurple', 'BrightColors', 'BrownBlue', 'Dark', 'Dark2', 'DarkAmber', 'DarkBlack', 'DarkBlack1', 'DarkBlue', 'DarkBlue1', 'DarkBlue10', 'DarkBlue11', 'DarkBlue12', 'DarkBlue13', 'DarkBlue14', 'DarkBlue15', 'DarkBlue16', 'DarkBlue17', 'DarkBlue2', 'DarkBlue3', 'DarkBlue4', 'DarkBlue5', 'DarkBlue6', 'DarkBlue7', 'DarkBlue8', 'DarkBlue9', 'DarkBrown', 'DarkBrown1', 'DarkBrown2', 'DarkBrown3', 'DarkBrown4', 'DarkBrown5', 'DarkBrown6', 'DarkGreen', 'DarkGreen1', 'DarkGreen2', 'DarkGreen3', 'DarkGreen4', 'DarkGreen5', 'DarkGreen6', 'DarkGrey', 'DarkGrey1', 'DarkGrey2', 'DarkGrey3', 'DarkGrey4', 'DarkGrey5', 'DarkGrey6', 'DarkGrey7', 'DarkPurple', 'DarkPurple1', 'DarkPurple2', 'DarkPurple3', 'DarkPurple4', 'DarkPurple5', 'DarkPurple6', 'DarkRed', 'DarkRed1', 'DarkRed2', 'DarkTanBlue', 'DarkTeal', 'DarkTeal1', 'DarkTeal10', 'DarkTeal11', 'DarkTeal12', 'DarkTeal2', 'DarkTeal3', 'DarkTeal4', 'DarkTeal5', 'DarkTeal6', 'DarkTeal7', 'DarkTeal8', 'DarkTeal9', 'Default', 'Default1', 'DefaultNoMoreNagging', 'Green', 'GreenMono', 'GreenTan', 'HotDogStand', 'Kayak', 'LightBlue', 'LightBlue1', 'LightBlue2', 'LightBlue3', 'LightBlue4', 'LightBlue5', 'LightBlue6', 'LightBlue7', 'LightBrown', 'LightBrown1', 'LightBrown10', 'LightBrown11', 'LightBrown12', 'LightBrown13', 'LightBrown2', 'LightBrown3', 'LightBrown4', 'LightBrown5', 'LightBrown6', 'LightBrown7', 'LightBrown8', 'LightBrown9', 'LightGray1', 'LightGreen', 'LightGreen1', 'LightGreen10', 'LightGreen2', 'LightGreen3', 'LightGreen4', 'LightGreen5', 'LightGreen6', 'LightGreen7', 'LightGreen8', 'LightGreen9', 'LightGrey', 'LightGrey1', 'LightGrey2', 'LightGrey3', 'LightGrey4', 'LightGrey5', 'LightGrey6', 'LightPurple', 'LightTeal', 'LightYellow', 'Material1', 'Material2', 'NeutralBlue', 'Purple', 'Reddit', 'Reds', 'SandyBeach', 'SystemDefault', 'SystemDefault1', 'SystemDefaultForReal', 'Tan', 'TanBlue', 'TealMono', 'Topanga']

    # Default values
    DEFAULT_THEME = 'LightGreen2'
    KEEP_ON_TOP = True
    OUTSIDE_TERM_TIME_MSG = '<OUTSIDE TERM TIME!>'
    STARTING_WINDOW_X = 585
    STARTING_WINDOW_Y = 427
    APP_VERSION = 'v0.3.0'

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

    SINGLE_DRAFT_THREAD_END_KEY = ('-SINGLE DRAFT THREAD-', '-SINGLE DRAFT THREAD ENDED-')
    
    ALL_DRAFT_THREAD_END_KEY = ('-ALL DRAFT THREAD-', '-ALL DRAFT THREAD ENDED-')

    instruments_list = ['piano', 'drum', 'guitar', 'vocal', 'music', 'singing', 'bass guitar', 'classical guitar']

    def __init__(self):
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
                f.write(f'{self.APP_VERSION_TITLE} = {self.APP_VERSION}')
                f.close()
            
        # Create config parser
        self.config = ConfigParser()
        self.config.read(self.CONFIG_PATH)               
            
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
                if self.display_pop_up_message("It seems that the information in your templates might be outdated or damaged. To fix this issue, would you like to try repairing the data?", True) == 'Yes':
                    
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
                    self.display_pop_up_message("Cannot proceed with invalid template data\nApp is terminating!", False)
                        
                    time.sleep(3)
                    sys.exit()

        self.main_window(repaired)
                        
    def get_theme(self) -> str:
        return self.config.get(self.PREFERENCES_SECTION, self.THEME)
        
    def get_email_mode(self) -> str:
        return self.config.get(self.PREFERENCES_SECTION, self.EMAIL_MODE)
    
    def get_current_template(self) -> str:
        return self.config.get(self.STATE_SECTION, self.CURRENT_TEMPLATE)  
        
    def get_phrases(self, start_date: datetime, end_date: datetime, half: str, term: str) -> list:
            
        body_phrase = str(self.num_to_weekday(start_date.isoweekday())) + ' ' + str(start_date.day) + self.get_day_suffix(start_date.day) + str(self.num_to_month(start_date.month)) + ' to and including ' + str(self.num_to_weekday(end_date.isoweekday())) + ' ' + str(end_date.day) + self.get_day_suffix(end_date.day) + str(self.num_to_month(end_date.month))
        
        current_term = [
            half + ' half ' + term + ' term ' + str(start_date.year), 
            half + ' Half ' + term.capitalize() + ' Term ' + str(end_date.year), body_phrase
        ]
        
        return current_term

    def get_day_suffix(self, day: int) -> str:
        match(day):
            case 1:
                return 'st '
            case 2:
                return 'nd '
            case 3:
                return 'rd '
            case _:
                return 'th '

    def get_term_length_in_weeks(self, start_date, end_date) -> timedelta:
        first_monday = (start_date - timedelta(days=start_date.weekday()))
        last_monday = (end_date - timedelta(days=end_date.weekday()))   
        
        return timedelta(weeks=(((last_monday - first_monday).days / 7) + 1))

    def which_term(self, date: datetime, day: str) -> tuple:
        current_term = [self.OUTSIDE_TERM_TIME_MSG, self.OUTSIDE_TERM_TIME_MSG, self.OUTSIDE_TERM_TIME_MSG] 

        for term in self.term_list:
            if (date >= term[0] and date <= term[1]):
                start_date = self.next_day_in_week(term[0], day)
                date_gap = self.get_term_length_in_weeks(term[0], term[1])
                end_date = start_date + date_gap
        
                current_term = self.get_phrases(start_date, end_date - timedelta(weeks=1), term[2], term[3])
                
                return current_term, int(date_gap.days / 7)
            
    def check_select_fields_are_not_empty(self, values: dict) -> bool:
        return len(str(values[self.RECIPIENT_INPUT])) == 0 or len(str(values[self.COST_INPUT])) == 0 or len(str(values[self.INSTRUMENT_INPUT])) == 0 or len(str(values[self.STUDENT_INPUT])) == 0

    def num_to_weekday(self, num: int) -> str:
        return self.weekdays[num-1]

    def num_to_month(self, num: int) -> str:
        return self.months[num-1]

    def next_day_in_week(self, date: datetime, target_day: str) -> datetime:    
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

    def get_names_list(self) -> list[str]:
        with open(self.TEMPLATES_PATH, 'r') as f:
            json_data = json.load(f)
            names_list = list(json_data.keys())
            names_list = sorted(names_list, key=str.lower)
            f.close()
            
        return names_list

    def get_subject(self, name: str) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
            json_data = json.load(f)
        
        day = str(json_data[name][self.DAY_INPUT])
        instrument = str(json_data[name][self.INSTRUMENT_INPUT])
        phrases, _ = self.which_term(self.current_date, day)
        
        return f"Invoice for {instrument.title()} Lessons {phrases[PhraseType.SUBJECT]}"
    
    def get_body(self, name: str) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
            json_data = json.load(f)
            f.close()
        
        cost_of_lessons = json_data[name][self.COST_INPUT]
        instrument = json_data[name][self.INSTRUMENT_INPUT]
        day = json_data[name][self.DAY_INPUT]
        students = json_data[name][self.STUDENT_INPUT]
        
        phrases, num_of_lessons = self.which_term(self.current_date, day)
        
        total_cost = '%.2f' % (round(float(int(num_of_lessons) * float(cost_of_lessons)), 2))
    
        num_of_lessons_phrase = f"There are {num_of_lessons} sessions"
        
        if int(num_of_lessons) == 1:
            num_of_lessons_phrase = f"There is {num_of_lessons} session"
    
        return f"Hi {name},\n\nHere is my invoice for {students}'s {instrument} lessons {phrases[PhraseType.INVOICE]}.\n--------\n{num_of_lessons_phrase} this {phrases[PhraseType.INVOICE]} from {phrases[PhraseType.DATES]}.\n\n{num_of_lessons} x £{cost_of_lessons} = £{total_cost}\n\nThank you\n--------\n\nKind regards\nRobert"

    def display_pop_up_message(self, title: str, question: bool) -> None|str:
        if question:
            choice = sg.popup_yes_no(title, title='', font=self.text_font, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, location=(self.get_last_win_x() + self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y))
    
            return choice
        else:
            sg.popup_quick_message(title, font=self.text_font, title='', icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, background_color="Black", text_color="White", location=(self.get_last_win_x() + self.WIN_OFFSET_X, self.get_last_win_y() + self.WIN_OFFSET_Y)) 
            
            return None
        
    def toggle_clipboard_visible(self, window: sg.Window, clipboard_visible: bool) -> None:
            window[self.DRAFT_BUTTON].update(visible=not clipboard_visible)
            window[self.DRAFT_ALL_BUTTON].update(visible=not clipboard_visible)
            window[self.SUBJECT_BUTTON].update(visible=clipboard_visible)
            window[self.BODY_BUTTON].update(visible=clipboard_visible)
            
    def save_config(self):
        with open(self.CONFIG_PATH, 'w') as config_file:
            self.config.write(config_file)
            config_file.close()       
            
    def create_draft_for_template(self, name: str):
        self.display_pop_up_message('Sending...', False)
        
        self.gmail_API.gmail_create_draft(self.get_subject(name), self.get_body(name))
        
    def get_last_win_x(self) -> int:
        return int(self.config.get(self.STATE_SECTION, self.LAST_WINDOW_X))
    
    def get_last_win_y(self) -> int:
        return int(self.config.get(self.STATE_SECTION, self.LAST_WINDOW_Y))
    
    def save_win_location(self, window: sg.Window):
        x, y = window.CurrentLocation()
                
        self.config.set(self.STATE_SECTION, self.LAST_WINDOW_X, str(x)) 
        self.config.set(self.STATE_SECTION, self.LAST_WINDOW_Y, str(y))
        
        self.save_config()
    
    def settings_window(self) -> bool:    
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
        window.read(timeout=0)
        
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

    def selected_template_window(self, is_new_template: bool, name: str = '') -> str:
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
        window.read(timeout=0)
        
        # Event Loop
        while True:
            event, values = window.read()
            
            if event == self.EXIT_BUTTON or event == sg.WIN_CLOSED:
                break
            if event == self.SAVE_BUTTON: 
                # Input error checking
                if re.search('\d', values[self.RECIPIENT_INPUT]):
                    sg.popup('Recipient name cannot contain numbers!', title='', font=self.text_font, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
                elif re.search('\D', str(values[self.COST_INPUT]).replace('.', '')):
                    sg.popup('Cost of lesson cannot contain characters!', title='', font=self.text_font, icon=self.BLANK_ICO,
                    keep_on_top=self.KEEP_ON_TOP)
                elif re.search('\d', values[self.STUDENT_INPUT]):
                    sg.popup('Students names cannot contain numbers!', title='', font=self.text_font, icon=self.BLANK_ICO,keep_on_top=self.KEEP_ON_TOP)
                elif self.check_select_fields_are_not_empty(values):
                    sg.popup('All fields must be completed!', title='', font=self.text_font, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
                else:
                    if (name in json_data) and is_new_template:
                        sg.popup('Template with that name already exists!', title='', font=self.text_font, icon=self.BLANK_ICO,keep_on_top=self.KEEP_ON_TOP)
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

    def get_main_window(self) -> sg.Window:
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
        window.read(timeout=0)
        
        return window

    def main_window(self, repaired: bool = False) -> None:        
        window = self.get_main_window()  
        
        if repaired:
            self.display_pop_up_message("Template data repaired successfully!", False)
        
        # Event Loop
        while True:
            event, values = window.read()
            
            if event == self.EXIT_BUTTON or event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT:
                self.config.set(self.STATE_SECTION, self.CURRENT_TEMPLATE, str(values[self.NAMES_COMBOBOX]))
                
                self.save_config()
                
                break

            if event == self.NAMES_COMBOBOX:
                self.toggle_buttons_disabled(window, self.template_edit_buttons_list, False)        
                    
            if event == self.DELETE_BUTTON:                
                if self.display_pop_up_message('Are you sure you want to delete this template?', True) == 'Yes':
                    with open(self.TEMPLATES_PATH, 'r') as f:
                        json_data = json.load(f)
                        f.close()
                        
                    del json_data[values[self.NAMES_COMBOBOX]]   
                    
                    with open(self.TEMPLATES_PATH, 'w') as f:
                        f.write(json.dumps(json_data))   
                        f.close()   
                    
                    window[self.NAMES_COMBOBOX].update(values=self.get_names_list())                  
                        
                    self.toggle_buttons_disabled(window, self.template_edit_buttons_list, True)
                    
                    self.display_pop_up_message('Template Deleted!', False)
                
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
                window.start_thread(lambda: self.create_draft_for_template(values[self.NAMES_COMBOBOX]), self.SINGLE_DRAFT_THREAD_END_KEY)
                
            if event == self.DRAFT_ALL_BUTTON:                
                if self.display_pop_up_message('Are you sure you want to draft all templates?', True) == 'Yes':

                    for name in self.get_names_list():
                        window.start_thread(lambda: self.create_draft_for_template(name), self.ALL_DRAFT_THREAD_END_KEY)
                
            if event[0] == self.SINGLE_DRAFT_THREAD_END_KEY[0]:
                self.display_pop_up_message('Draft Sent!', False)
                
            if event[0] == self.ALL_DRAFT_THREAD_END_KEY[0]:
                self.display_pop_up_message('All Drafts Sent!', False)
                
            if event == self.SUBJECT_BUTTON:
                pyperclip.copy(self.get_subject(values[self.NAMES_COMBOBOX]))
                
                self.display_pop_up_message('Copied Subject!', False)
                
            if event == self.BODY_BUTTON:
                pyperclip.copy(self.get_body(values[self.NAMES_COMBOBOX]))
                
                self.display_pop_up_message('Copied Body!', False) 

            if event == self.SETTINGS_BUTTON:
                self.save_win_location(window)
                
                # Reload window if settings were saved
                if self.settings_window():
                    self.config.set(self.STATE_SECTION, self.CURRENT_TEMPLATE, str(values[self.NAMES_COMBOBOX]))
                
                    self.save_config()
                    
                    sg.theme(self.get_theme())
                    window.close()
                    window = self.get_main_window()
                    
                    self.display_pop_up_message('Settings Saved!', False)
                
        window.close()