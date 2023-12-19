from math import floor
import random
import PySimpleGUI as sg
import json
import re
import pyperclip
from datetime import datetime, timedelta
from enum import IntEnum
from gmailAPI import gmail_create_draft
import os.path
from configparser import ConfigParser

class PhraseType(IntEnum):
    INVOICE = 0
    SUBJECT = 1
    DATES = 2

class InvoiceApp:
    # Theme List
    themeList = ['Black', 'BlueMono', 'BluePurple', 'BrightColors', 'BrownBlue', 'Dark', 'Dark2', 'DarkAmber', 'DarkBlack', 'DarkBlack1', 'DarkBlue', 'DarkBlue1', 'DarkBlue10', 'DarkBlue11', 'DarkBlue12', 'DarkBlue13', 'DarkBlue14', 'DarkBlue15', 'DarkBlue16', 'DarkBlue17', 'DarkBlue2', 'DarkBlue3', 'DarkBlue4', 'DarkBlue5', 'DarkBlue6', 'DarkBlue7', 'DarkBlue8', 'DarkBlue9', 'DarkBrown', 'DarkBrown1', 'DarkBrown2', 'DarkBrown3', 'DarkBrown4', 'DarkBrown5', 'DarkBrown6', 'DarkGreen', 'DarkGreen1', 'DarkGreen2', 'DarkGreen3', 'DarkGreen4', 'DarkGreen5', 'DarkGreen6', 'DarkGrey', 'DarkGrey1', 'DarkGrey2', 'DarkGrey3', 'DarkGrey4', 'DarkGrey5', 'DarkGrey6', 'DarkGrey7', 'DarkPurple', 'DarkPurple1', 'DarkPurple2', 'DarkPurple3', 'DarkPurple4', 'DarkPurple5', 'DarkPurple6', 'DarkRed', 'DarkRed1', 'DarkRed2', 'DarkTanBlue', 'DarkTeal', 'DarkTeal1', 'DarkTeal10', 'DarkTeal11', 'DarkTeal12', 'DarkTeal2', 'DarkTeal3', 'DarkTeal4', 'DarkTeal5', 'DarkTeal6', 'DarkTeal7', 'DarkTeal8', 'DarkTeal9', 'Default', 'Default1', 'DefaultNoMoreNagging', 'Green', 'GreenMono', 'GreenTan', 'HotDogStand', 'Kayak', 'LightBlue', 'LightBlue1', 'LightBlue2', 'LightBlue3', 'LightBlue4', 'LightBlue5', 'LightBlue6', 'LightBlue7', 'LightBrown', 'LightBrown1', 'LightBrown10', 'LightBrown11', 'LightBrown12', 'LightBrown13', 'LightBrown2', 'LightBrown3', 'LightBrown4', 'LightBrown5', 'LightBrown6', 'LightBrown7', 'LightBrown8', 'LightBrown9', 'LightGray1', 'LightGreen', 'LightGreen1', 'LightGreen10', 'LightGreen2', 'LightGreen3', 'LightGreen4', 'LightGreen5', 'LightGreen6', 'LightGreen7', 'LightGreen8', 'LightGreen9', 'LightGrey', 'LightGrey1', 'LightGrey2', 'LightGrey3', 'LightGrey4', 'LightGrey5', 'LightGrey6', 'LightPurple', 'LightTeal', 'LightYellow', 'Material1', 'Material2', 'NeutralBlue', 'Purple', 'Reddit', 'Reds', 'SandyBeach', 'SystemDefault', 'SystemDefault1', 'SystemDefaultForReal', 'Tan', 'TanBlue', 'TealMono', 'Topanga']

    # Default values
    DEFAULT_THEME = 'DarkAmber'
    KEEP_ON_TOP = True
    OUTSIDE_TERM_TIME_MSG = '<OUTSIDE TERM TIME!>'
    STARTING_WINDOW_X = 585
    STARTING_WINDOW_Y = 427

    # Term Dates
    today = datetime.now()
    currentDate = today

    # Debug code for when out of term time
    # year = int(sg.popup_get_text('YEAR', size= 10, keep_on_top=KEEP_ON_TOP))
    # month = int(sg.popup_get_text('MONTH', size= 10, keep_on_top=KEEP_ON_TOP))
    # day = int(sg.popup_get_text('DAY', size= 10, keep_on_top=KEEP_ON_TOP))
    # currentDate = datetime(year, month, day)

    autumn1 = [datetime(2023, 9, 4), datetime(2023, 10, 21), '1st', 'autumn']
    autumn2 = [datetime(2023, 10, 30), datetime(2023, 12, 23), '2nd', 'autumn']

    spring1 = [datetime(2024, 1, 8), datetime(2024, 2, 10), '1st', 'spring']
    spring2  = [datetime(2024, 2, 19), datetime(2024, 3, 30), '2nd', 'spring']

    summer1 = [datetime(2024, 4, 15), datetime(2024, 5, 25), '1st', 'summer']
    summer2 = [datetime(2024, 5, 3), datetime(2024, 6, 24), '2nd', 'summer']

    termList = [autumn1, autumn2, spring1, spring2, summer1, summer2]

    weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November','December']

    # Custom Fonts
    textFont = ('Lucida Console', 13)
    
    # Colours
    READONLY_BACKGROUND_COLOUR = '#FF6961'

    # File Paths
    TEMPLATES_PATH = 'templates.json'
    SETTINGS_PATH = 'settings.ini'

    # Config Sections
    PREFERENCES_SECTION = 'Preferences'
    STATE_SECTION = 'State'
    
    # Config Values
    THEME = 'Theme'
    EMAIL_MODE = 'email-mode'
    CURRENT_TEMPLATE = 'current-template'
    LAST_WINDOW_X = 'last-win-x'
    LAST_WINDOW_Y = 'last-win-y'

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
    
    templateEditButtonsList = [DELETE_BUTTON, DRAFT_BUTTON, DRAFT_ALL_BUTTON, BODY_BUTTON, SUBJECT_BUTTON, EDIT_BUTTON]

    # Email Modes
    CLIPBOARD = 'Clipboard'
    AUTO_DRAFT = 'Auto Draft'

    # Combo Values
    NAMES_COMBOBOX = 'Names'
    THEME_COMBOBOX = '<Theme>'

    # Input Values
    RECIPIENT_INPUT = 'Recipient'
    COST_INPUT = 'Cost'
    INSTRUMENT_INPUT = 'Instrument'
    DAY_INPUT = 'Day'
    STUDENT_INPUT = 'Students'
    INPUT_SIZE = 15
    THEME_INPUT_SIZE = 21

    instrumentsList = ['piano', 'drum', 'guitar', 'vocal', 'music', 'singing', 'bass guitar', 'classical guitar']

    def __init__(self):
        # Create resources directory if it does not exist
        if not os.path.exists(self.RESOURCE_DIR):
            os.makedirs(self.RESOURCE_DIR)
        
        # Create templates file if it does not exist
        if (not os.path.isfile(self.TEMPLATES_PATH)):
            with open(self.TEMPLATES_PATH, 'w') as f:
                f.write('{}')
                f.close()
        
        # Create settings file if it does not exist
        if (not os.path.isfile(self.SETTINGS_PATH)):
            with open(self.SETTINGS_PATH, 'w') as f:
                f.write(f'[{self.PREFERENCES_SECTION}]\n')
                f.write(f'{self.THEME} = ' + self.DEFAULT_THEME + '\n')
                f.write(f'{self.EMAIL_MODE} = {self.CLIPBOARD}\n')
            
                f.write(f'\n[{self.STATE_SECTION}]\n')
                f.write(f'{self.CURRENT_TEMPLATE} = \n')
                f.write(f'{self.LAST_WINDOW_X} = {self.STARTING_WINDOW_X}\n')
                f.write(f'{self.LAST_WINDOW_Y} = {self.STARTING_WINDOW_Y}\n')
                f.close()
        
        # Create config parser
        self.config = ConfigParser()
        self.config.read(self.SETTINGS_PATH)
            
        sg.theme(self.getTheme())
    
    def run(self):
        self.mainWindow()
        
    def getTheme(self) -> str:
        return self.config.get(self.PREFERENCES_SECTION, self.THEME)
        
    def getEmailMode(self) -> str:
        return self.config.get(self.PREFERENCES_SECTION, self.EMAIL_MODE)
    
    def getCurrentTemplate(self) -> str:
        return self.config.get(self.STATE_SECTION, self.CURRENT_TEMPLATE)  
        
    def getPhrases(self, startDate: datetime, endDate: datetime, half: str, term: str) -> list:
            
        bodyPhrase = str(self.numToWeekday(startDate.isoweekday())) + ' ' + str(startDate.day) + self.getDaySuffix(startDate.day) + str(self.numToMonth(startDate.month)) + ' to and including ' + str(self.numToWeekday(endDate.isoweekday())) + ' ' + str(endDate.day) + self.getDaySuffix(endDate.day) + str(self.numToMonth(endDate.month))
        
        currentTerm = [
            half + ' half ' + term + ' term ' + str(startDate.year), 
            half + ' Half ' + term.capitalize() + ' Term ' + str(endDate.year), bodyPhrase
        ]
        
        return currentTerm

    def getDaySuffix(self, day: int) -> str:
        match(day):
            case 1:
                return 'st '
            case 2:
                return 'nd '
            case 3:
                return 'rd '
            case _:
                return 'th '

    def getTermLengthInWeeks(self, startDate, endDate) -> timedelta:
        monday1 = (startDate - timedelta(days=startDate.weekday()))
        monday2 = (endDate - timedelta(days=endDate.weekday()))   
        
        return timedelta(weeks=(((monday2 - monday1).days / 7) + 1))

    def whichTerm(self, date: datetime, day: str) -> tuple:
        currentTerm = [self.OUTSIDE_TERM_TIME_MSG, self.OUTSIDE_TERM_TIME_MSG, self.OUTSIDE_TERM_TIME_MSG] 

        for term in self.termList:
            if (date >= term[0] and date <= term[1]):
                startDate = self.nextDayInWeek(term[0], day)
                dateGap = self.getTermLengthInWeeks(term[0], term[1])
                endDate = startDate + dateGap
        
                currentTerm = self.getPhrases(startDate, endDate - timedelta(weeks=1), term[2], term[3])
                
                return currentTerm, int(dateGap.days / 7)
            
    def checkSelectFieldsAreNotEmpty(self, values: dict) -> bool:
        return len(str(values[self.RECIPIENT_INPUT])) == 0 or len(str(values[self.COST_INPUT])) == 0 or len(str(values[self.INSTRUMENT_INPUT])) == 0 or len(str(values[self.STUDENT_INPUT])) == 0

    def numToWeekday(self, num: int) -> str:
        return self.weekdays[num-1]

    def numToMonth(self, num: int) -> str:
        return self.months[num-1]

    def nextDayInWeek(self, date: datetime, targetDay: str) -> datetime:    
        dayFound = False
        searchDay = date
        
        while dayFound == False:
            if (self.numToWeekday(searchDay.isoweekday()) == targetDay):
                dayFound = True
            else:
                searchDay = searchDay + timedelta(days=1)
            
        return searchDay

    def toggleButtonsDisabled(self, window: sg.Window, buttonList: str, disabled: bool):
        for button in buttonList:
            window[button].update(disabled=disabled)

    def getNamesList(self):
        with open(self.TEMPLATES_PATH, 'r') as f:
            jsonData = json.load(f)
            namesList = list(jsonData.keys())
            namesList = sorted(namesList, key=str.lower)
            f.close()
            
        return namesList

    def getSubject(self, name: str) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
            jsonData = json.load(f)
         
        day = str(jsonData[name][self.DAY_INPUT])
        instrument = str(jsonData[name][self.INSTRUMENT_INPUT])
        phrases, _ = self.whichTerm(self.currentDate, day)
        
        return f"Invoice for {instrument.title()} Lessons {phrases[PhraseType.SUBJECT]}"
    
    def getBody(self, name: str) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
            jsonData = json.load(f)
            f.close()
        
        costOfLessons = jsonData[name][self.COST_INPUT]
        instrument = jsonData[name][self.INSTRUMENT_INPUT]
        day = jsonData[name][self.DAY_INPUT]
        students = jsonData[name][self.STUDENT_INPUT]
        
        phrases, numOfLessons = self.whichTerm(self.currentDate, day)
        
        totalCost = '%.2f' % (round(float(int(numOfLessons) * float(costOfLessons)), 2))
    
        numberOfLessonsPhrase = f"There are {numOfLessons} sessions"
        
        if int(numOfLessons) == 1:
            numberOfLessonsPhrase = f"There is {numOfLessons} session"
    
        return f"Hi {name},\n\nHere is my invoice for {students}'s {instrument} lessons {phrases[PhraseType.INVOICE]}.\n--------\n{numberOfLessonsPhrase} this {phrases[PhraseType.INVOICE]} from {phrases[PhraseType.DATES]}.\n\n{numOfLessons} x £{costOfLessons} = £{totalCost}\n\nThank you\n--------\n\nKind regards\nRobert"

    def displayPopUpMessage(self, title: str, yesorno: bool) -> None|str:
        if yesorno:
             choice = sg.popup_yes_no(title, title='', font=self.textFont, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
             
             return choice
        else:
            sg.popup_quick_message(title, font=self.textFont, title='', icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, background_color="Black", text_color="White")
            
            return None
            
    def settingsWindow(self):        
        layout = [
                    [
                        sg.Text(self.THEME_COMBOBOX, font=self.textFont, pad=self.SETTINGS_PADDING), 
                        sg.Combo(self.themeList, font=self.textFont, size=self.THEME_INPUT_SIZE, readonly=True, key=self.THEME_COMBOBOX, default_value=self.getTheme()),
                        sg.Button('Randomise', font=self.textFont, pad=self.SETTINGS_PADDING)
                    ],
                    [
                        sg.Text('<Email Mode>', font=self.textFont, pad=self.SETTINGS_PADDING), 
                        sg.Combo([self.CLIPBOARD, self.AUTO_DRAFT], font=self.textFont, size=self.THEME_INPUT_SIZE, readonly=True, key='Email Mode', default_value=self.getEmailMode()),
                    ],
                    [
                        sg.VPush()
                    ],
                    [
                        sg.Button(self.SAVE_BUTTON, font=self.textFont),
                        sg.Push(), 
                        sg.Button(self.EXIT_BUTTON, font=self.textFont)
                    ]
            ]
        
        window = sg.Window(self.SETTINGS_BUTTON, layout, element_justification='l', size=(self.SETTINGS_WIDTH, self.SETTINGS_HEIGHT), modal=True, icon=self.MAIN_ICO, keep_on_top=self.KEEP_ON_TOP)
        
        # Event Loop
        while True:
            event, values = window.read()
            if event == self.EXIT_BUTTON or event == sg.WIN_CLOSED:
                break
            if event == 'Randomise':
                window[self.THEME_COMBOBOX].update(value=random.choice(self.themeList))
                
            if event == self.SAVE_BUTTON:
                self.config.set(self.PREFERENCES_SECTION, self.THEME, values[self.THEME_COMBOBOX])
                self.config.set(self.PREFERENCES_SECTION, self.EMAIL_MODE, values['Email Mode'])
                
                with open(self.SETTINGS_PATH, 'w') as configfile:
                    self.config.write(configfile)
                    configfile.close()
                break
            
        window.close()

    def selectedTemplateWindow(self, isNewTemplate: bool, name: str) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
            jsonData = json.load(f)
            f.close()
            
        recipientDefault = ''
        recipientDisabled = False
        costDefault = ''
        instrumentDefault = ''
        dayDefault = ''
        studentDefault = ''
        
        if not isNewTemplate:
            recipientDefault = name
            recipientDisabled = True
            costDefault = jsonData[name][self.COST_INPUT]   
            instrumentDefault = jsonData[name][self.INSTRUMENT_INPUT]
            dayDefault = jsonData[name][self.DAY_INPUT]
            studentDefault = jsonData[name][self.STUDENT_INPUT] 
        
        layout = [
                    [
                        sg.Text('Recipient', font=self.textFont, pad=self.SELECT_PADDING), 
                        sg.Input(size=self.INPUT_SIZE*2, font=self.textFont, key=self.RECIPIENT_INPUT, default_text=recipientDefault, disabled=recipientDisabled, disabled_readonly_background_color=self.READONLY_BACKGROUND_COLOUR)
                    ],
                    [
                        sg.Text('Cost of lesson  £', font=self.textFont, pad=self.SELECT_PADDING), 
                        sg.Input(size=self.INPUT_SIZE, font=self.textFont, key=self.COST_INPUT, default_text=costDefault)
                    ],
                    [   sg.Text('Instrument', font=self.textFont, pad=self.SELECT_PADDING), 
                        sg.Combo(values=sorted(self.instrumentsList), size=self.INPUT_SIZE*2,
                        font=self.textFont, key=self.INSTRUMENT_INPUT, default_value=instrumentDefault, readonly=True)
                    ],
                    [
                        sg.Text('Start Day', font=self.textFont, pad=self.SELECT_PADDING),
                        sg.Combo(values=self.weekdays, size=self.INPUT_SIZE, font=self.textFont, key=self.DAY_INPUT, default_value=dayDefault, readonly=True)
                    ],
                    [   
                        sg.Text('Student(s)', font=self.textFont)
                    ], 
                    [
                        sg.Multiline(size=(self.INPUT_SIZE*2, 2), font=self.textFont, key=self.STUDENT_INPUT, default_text=studentDefault)
                    ],
                    [
                        sg.VPush()
                    ],
                    [
                        sg.Button(self.SAVE_BUTTON, font=self.textFont), 
                        sg.Push(), 
                        sg.Button(self.EXIT_BUTTON, font=self.textFont)
                    ]
                ]
        
        if isNewTemplate:
            templateTitle = self.NEW_TEMPLATE_BUTTON
        else:
            templateTitle = self.EDIT_TEMPLATE_TITLE
        
        window = sg.Window(templateTitle, layout, element_justification='l', size=(self.SELECT_WIDTH, self.SELECT_HEIGHT), modal=True, icon=self.MAIN_ICO, keep_on_top=self.KEEP_ON_TOP)
        
        # Event Loop
        while True:
            event, values = window.read()
            if event == self.EXIT_BUTTON or event == sg.WIN_CLOSED:
                break
            if event == self.SAVE_BUTTON: 
                # Input error checking
                if re.search('\d', values[self.RECIPIENT_INPUT]):
                    sg.popup('Recipient name cannot contain numbers!', title='', font=self.textFont, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
                elif re.search('\D', str(values[self.COST_INPUT]).replace('.', '')):
                    sg.popup('Cost of lesson cannot contain characters!', title='', font=self.textFont, icon=self.BLANK_ICO,
                    keep_on_top=self.KEEP_ON_TOP)
                elif re.search('\d', values[self.STUDENT_INPUT]):
                    sg.popup('Students names cannot contain numbers!', title='', font=self.textFont, icon=self.BLANK_ICO,keep_on_top=self.KEEP_ON_TOP)
                elif self.checkSelectFieldsAreNotEmpty(values):
                    sg.popup('All fields must be completed!', title='', font=self.textFont, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
                else:
                    if (name in jsonData) and (isNewTemplate):
                        sg.popup('Template with that name already exists!', title='', font=self.textFont, icon=self.BLANK_ICO,keep_on_top=self.KEEP_ON_TOP)
                    else:
                        name = values[self.RECIPIENT_INPUT]
                        
                        info = {
                            self.COST_INPUT : '%.2f' % (round(float(values[self.COST_INPUT]), 2)),
                            self.INSTRUMENT_INPUT : values[self.INSTRUMENT_INPUT],
                            self.DAY_INPUT: values[self.DAY_INPUT],
                            self.STUDENT_INPUT : values[self.STUDENT_INPUT]
                        }
                        
                        jsonData[name] = info

                        with open(self.TEMPLATES_PATH, 'w') as f:
                            f.write(json.dumps(jsonData))
                            f.close()
                            
                        break
                
        window.close()
        
        return name

    def mainWindow(self):
        namesList = self.getNamesList()
        
        if (self.getEmailMode() == self.CLIPBOARD):
            subBodyEnabled = True
            draftEnabled = False
        else:
            draftEnabled = True
            subBodyEnabled = False
            
        currentTemplate = self.getCurrentTemplate()
        
        supportButtonsDisabled = False
        
        if currentTemplate == "":
            supportButtonsDisabled = True
        
        supportButtons = [
                            [
                                sg.Button(self.DELETE_BUTTON, font=self.textFont, disabled=supportButtonsDisabled),
                                sg.Push(),
                                sg.Button(self.NEW_TEMPLATE_BUTTON, font=self.textFont), 
                                sg.Button(self.SETTINGS_BUTTON, font=self.textFont),
                                sg.Button(self.EXIT_BUTTON, font=self.textFont)
                            ]       
                        ]

        layout = [  
                    [
                        [
                            sg.Text('<Templates>', font=self.textFont),
                            sg.Combo(namesList, enable_events=True, default_value=currentTemplate, pad=self.MAIN_PADDING, key=self.NAMES_COMBOBOX, size=15, font=self.textFont, readonly=True),
                            sg.Button(self.EDIT_BUTTON, font=self.textFont, disabled=supportButtonsDisabled),
                            sg.Button(self.DRAFT_BUTTON, font=self.textFont, disabled=supportButtonsDisabled, visible=draftEnabled),
                            sg.Button(self.DRAFT_ALL_BUTTON, font=self.textFont, disabled=supportButtonsDisabled, visible=draftEnabled),
                            sg.Button(self.SUBJECT_BUTTON, font=self.textFont, disabled=supportButtonsDisabled, visible=subBodyEnabled),
                            sg.Button(self.BODY_BUTTON, font=self.textFont, disabled=supportButtonsDisabled, visible=subBodyEnabled)
                        ]
                    ],
                    [
                        sg.VPush()
                    ],
                    [
                        sg.Column(supportButtons, element_justification='right',expand_x=True)
                    ]
                ]

        window = sg.Window('Invoice Templates', layout, element_justification='l', size=(self.MAIN_WIDTH, self.MAIN_HEIGHT), icon=self.MAIN_ICO, keep_on_top=self.KEEP_ON_TOP)
        
        # Event Loop
        while True:
            event, values = window.read()
            
            if event == sg.WIN_CLOSED or event == self.EXIT_BUTTON:                
                self.config.set(self.PREFERENCES_SECTION, self.CURRENT_TEMPLATE, window[self.NAMES_COMBOBOX].get())
                
                with open(self.SETTINGS_PATH, 'w') as configfile:
                    self.config.write(configfile)
                    configfile.close()
                break
                   
            if event == self.NAMES_COMBOBOX:
                self.toggleButtonsDisabled(window, self.templateEditButtonsList, False)        
                    
            if event == self.DELETE_BUTTON:                
                if self.displayPopUpMessage('Are you sure you want to delete this template?', True) == 'Yes':
                    with open(self.TEMPLATES_PATH, 'r') as f:
                        jsonData = json.load(f)
                        f.close()
                        
                    del jsonData[values[self.NAMES_COMBOBOX]]   
                    
                    namesList = list(jsonData.keys())
                    namesList = sorted(namesList, key=str.lower)
                    window[self.NAMES_COMBOBOX].update(values=namesList) 
                    
                    with open(self.TEMPLATES_PATH, 'w') as f:
                        f.write(json.dumps(jsonData))   
                        f.close()        
                        
                    self.toggleButtonsDisabled(window, self.templateEditButtonsList, True)
                    
                    self.displayPopUpMessage('Template Deleted!', False)
                
            if event == self.EDIT_BUTTON:
                self.selectedTemplateWindow(False, values[self.NAMES_COMBOBOX])
                
            if event == self.NEW_TEMPLATE_BUTTON:
                name = self.selectedTemplateWindow(True, '')
                
                if name != '':
                    with open(self.TEMPLATES_PATH, 'r') as f:
                        jsonData = json.load(f)
                        namesList = list(jsonData.keys())
                        namesList = sorted(namesList, key=str.lower)
                        f.close()
                    
                    window[self.NAMES_COMBOBOX].update(value=name, values=namesList)
                    
                    self.toggleButtonsDisabled(window, self.templateEditButtonsList, False)
            
            if event == self.DRAFT_BUTTON:
                self.displayPopUpMessage('Sending...', False)
                
                gmail_create_draft(self.getSubject(values[self.NAMES_COMBOBOX]), self.getBody(values[self.NAMES_COMBOBOX]))
                
                self.displayPopUpMessage('Draft Sent!', False)

                
            if event == self.DRAFT_ALL_BUTTON:                
                if self.displayPopUpMessage('Are you sure you want to draft all templates?', True) == 'Yes':
                    self.displayPopUpMessage('Sending...', False)
                    
                    for name in self.getNamesList():
                        gmail_create_draft(self.getSubject(name), self.getBody(name))
                
                    self.displayPopUpMessage('All Drafts Sent!', False)
                
            if event == self.SUBJECT_BUTTON:
                pyperclip.copy(self.getSubject(values[self.NAMES_COMBOBOX]))
                
                self.displayPopUpMessage('Copied Subject!', False)
                
            if event == self.BODY_BUTTON:
                pyperclip.copy(self.getBody(values[self.NAMES_COMBOBOX]))
                
                self.displayPopUpMessage('Copied Body!', False)

            if event == self.SETTINGS_BUTTON:
                self.settingsWindow()
                
                sg.theme(self.getTheme())
                
                x, y = window.CurrentLocation()
                
                self.config.set(self.STATE_SECTION, self.LAST_WINDOW_X, x)
                self.config.set(self.STATE_SECTION, self.LAST_WINDOW_Y, y)
                
                with open(self.SETTINGS_PATH, 'w') as configfile:
                    self.config.write(configfile)
                    configfile.close()
                
                window.close()
                self.run()
                
                window.move(self.config.get(self.STATE_SECTION, self.LAST_WINDOW_X), self.config.get(self.STATE_SECTION, self.LAST_WINDOW_Y))
                
                if (self.config.get(self.PREFERENCES_SECTION, self.EMAIL_MODE) == self.CLIPBOARD):
                    window[self.DRAFT_BUTTON].update(visible=False)
                    window[self.DRAFT_ALL_BUTTON].update(visible=False)
                    window[self.SUBJECT_BUTTON].update(visible=True)
                    window[self.BODY_BUTTON].update(visible=True)
                else:
                    window[self.DRAFT_BUTTON].update(visible=True)
                    window[self.DRAFT_ALL_BUTTON].update(visible=True)
                    window[self.SUBJECT_BUTTON].update(visible=False)
                    window[self.BODY_BUTTON].update(visible=False)

        window.close()