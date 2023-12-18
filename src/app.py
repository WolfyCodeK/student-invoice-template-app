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

    # Term Dates
    today = datetime.now()
    currentDate = today

    # Debug code for when out of term time
    year = int(sg.popup_get_text('YEAR', size= 10, keep_on_top=KEEP_ON_TOP))
    month = int(sg.popup_get_text('MONTH', size= 10, keep_on_top=KEEP_ON_TOP))
    day = int(sg.popup_get_text('DAY', size= 10, keep_on_top=KEEP_ON_TOP))
    currentDate = datetime(year, month, day)

    autumn1 = [datetime(2023, 9, 4), datetime(2023, 10, 21)]
    autumn2 = [datetime(2023, 10, 30), datetime(2023, 12, 23)]

    spring1 = [datetime(2024, 1, 8), datetime(2024, 2, 10)]
    spring2  = [datetime(2024, 2, 19), datetime(2024, 3, 30)]

    summer1 = [datetime(2024, 4, 15), datetime(2024, 5, 25)]
    summer2 = [datetime(2024, 5, 3), datetime(2024, 6, 24)]

    weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November','December']

    # Custom Fonts
    textFont = ('Lucida Console', 13)

    # File Paths
    TEMPLATES_PATH = 'templates.json'
    SETTINGS_PATH = 'settings.ini'

    # Resource Paths
    RESOURCE_DIR = 'res/'
    BLANK_ICO = 'res/Blank.ico'
    MAIN_ICO = 'res/Email.ico'

    # Window Sizes
    MAIN_WIDTH = 600
    MAIN_HEIGHT = 225
    SELECT_WIDTH = 350
    SELECT_HEIGHT = 350
    SETTINGS_WIDTH = 475
    SETTINGS_HEIGHT = 150

    # Element Sizes
    MAIN_PADDING = 15
    SELECT_PADDING = 10
    SETTINGS_PADDING = 10

    # Button Values
    EDIT_BUTTON = 'Edit'
    EXIT_BUTTON = 'Exit'
    NEW_BUTTON = 'New Template'
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
    NUMBER_INPUT = 'Number'
    COST_INPUT = 'Cost'
    INSTRUMENT_INPUT = 'Instrument'
    DAY_INPUT = 'Day'
    STUDENT_INPUT = 'Student'
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
        
        # Create settings file if it does not exist
        if (not os.path.isfile(self.SETTINGS_PATH)):
            with open(self.SETTINGS_PATH, 'w') as f:
                f.write('[Preferences]\n')
                f.write('Theme = ' + self.DEFAULT_THEME + '\n')
                f.write(f'Email-Mode = {self.CLIPBOARD}\n')
                f.write('Email-Recipient = example@gmail.com')
        
        # Create config parser
        self.config = ConfigParser()
        self.config.read(self.SETTINGS_PATH)
            
        sg.theme(self.getTheme())
    
    def run(self):
        self.mainWindow()
        
    def getTheme(self) -> str:
        return self.config.get('Preferences', 'Theme')
        
    def getEmailMode(self) -> str:
        return self.config.get('Preferences', 'Email-Mode')
    
    def getEmailRecipient(self) -> str:
        return self.config.get('Preferences', 'Email-Recipient')
    
    def getCurrentTemplate(self) -> str:
        return self.config.get('Preferences', 'Current-Template')  
        
    def getPhrases(self, startDate: datetime, endDate: datetime, half: str, term: str) -> list:
            
        bodyPhrase = str(self.numToWeekday(startDate.isoweekday())) + ' ' + str(startDate.day) + self.getDaySuffix(startDate.day) + str(self.numToMonth(startDate.month)) + ' to and including ' + str(self.numToWeekday(endDate.isoweekday())) + ' ' + str(endDate.day) + self.getDaySuffix(endDate.day) + str(self.numToMonth(endDate.month))
        currentTerm = [half + ' half ' + term + ' term ' + str(startDate.year), half + ' Half ' + term.capitalize() + ' Term ' + str(endDate.year), bodyPhrase]
        
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

    def whichTerm(self, date: datetime, numberOfLessons: int, day: str) -> list:
        currentTerm = ['<DATE>', '<DATE>', '<DATE>']
        dateGap = timedelta(weeks=int(numberOfLessons))
            
        if (date >= self.autumn1[0] and date <= self.autumn1[1]):
            startDate = self.nextDayInWeek(self.autumn1[0], day)
            endDate = startDate + dateGap
            
            currentTerm = self.getPhrases(startDate, endDate, '1st', 'autumn')
            
        elif (date >= self.autumn2[0] and date <= self.autumn2[1]):
            startDate = self.nextDayInWeek(self.autumn2[0], day)
            endDate = startDate + dateGap
            
            currentTerm = self.getPhrases(startDate, endDate, '2nd', 'autumn')
            
        elif (date >= self.spring1[0] and date <= self.spring1[1]):
            startDate = self.nextDayInWeek(self.spring1[0], day)
            endDate = startDate + dateGap
            
            currentTerm = self.getPhrases(startDate, endDate, '1st', 'spring')
            
        elif (date >= self.spring2[0] and date <= self.spring2[1]):
            startDate = self.nextDayInWeek(self.spring2[0], day)
            endDate = startDate + dateGap
            
            currentTerm = self.getPhrases(startDate, endDate, '2nd', 'spring')
            
        elif (date >= self.summer1[0] and date <= self.summer1[1]):
            startDate = self.nextDayInWeek(self.summer1[0], day)
            endDate = startDate + dateGap
            
            currentTerm = self.getPhrases(startDate, endDate, '1st', 'summer')
            
        elif (date >= self.summer2[0] and date <= self.summer2[1]):
            startDate = self.nextDayInWeek(self.summer2[0], day)
            endDate = startDate + dateGap
            
            currentTerm = self.getPhrases(startDate, endDate, '2nd', 'summer') 
            
        return currentTerm
            
    def checkSelectFieldsAreNotEmpty(self, values: dict) -> bool:
        return len(str(values[self.RECIPIENT_INPUT])) == 0 or len(str(values[self.NUMBER_INPUT])) == 0 or len(str(values[self.COST_INPUT])) == 0 or len(str(values[self.INSTRUMENT_INPUT])) == 0 or len(str(values[self.STUDENT_INPUT])) == 0

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

    def getSubject(self, values: dict) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
                    jsonData = json.load(f)
         
        day = str(jsonData[values[self.NAMES_COMBOBOX]][self.DAY_INPUT])
        numberOfLessons = str(jsonData[values[self.NAMES_COMBOBOX]][self.NUMBER_INPUT])
        instrument = str(jsonData[values[self.NAMES_COMBOBOX]][self.INSTRUMENT_INPUT])
        phrases = self.whichTerm(self.currentDate, numberOfLessons, day)
        
        return f"Invoice for {instrument.title()} Lessons {phrases[PhraseType.SUBJECT]}"
    
    def getBody(self, values: dict) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
            jsonData = json.load(f)
        
        name = values[self.NAMES_COMBOBOX]
        numberOfLessons = jsonData[values[self.NAMES_COMBOBOX]][self.NUMBER_INPUT]
        costOfLessons = jsonData[values[self.NAMES_COMBOBOX]][self.COST_INPUT]
        totalCost = int(numberOfLessons) * float(costOfLessons)
        totalCost = '%.2f' % (round(float(totalCost), 2))
        instrument = jsonData[values[self.NAMES_COMBOBOX]][self.INSTRUMENT_INPUT]
        day = jsonData[values[self.NAMES_COMBOBOX]][self.DAY_INPUT]
        students = jsonData[values[self.NAMES_COMBOBOX]][self.STUDENT_INPUT]
        
        phrases = self.whichTerm(self.currentDate, numberOfLessons, day)
    
        numberOfLessonsPhrase = f"There are {numberOfLessons} sessions"
        
        if int(numberOfLessons) == 1:
            numberOfLessonsPhrase = f"There is {numberOfLessons} session"
    
        return f"Hi {name},\n\nHere is my invoice for {students}'s {instrument} lessons {phrases[PhraseType.INVOICE]}.\n--------\n{numberOfLessonsPhrase} this {phrases[PhraseType.INVOICE]} from {phrases[PhraseType.DATES]}.\n\n{numberOfLessons} x £{costOfLessons} = £{totalCost}\n\nThank you\n--------\n\nKind regards\nRobert"

    def settingsWindow(self):        
        layout = [
                    [
                        sg.Text(self.THEME_COMBOBOX, font=self.textFont, pad=self.SETTINGS_PADDING), 
                        sg.Combo(self.themeList, font=self.textFont, size=self.THEME_INPUT_SIZE, readonly=True, key=self.THEME_COMBOBOX, default_value=self.getTheme()),
                        sg.Button('Randomise', font=self.textFont)
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
        
        window = sg.Window('', layout, element_justification='l', size=(self.SETTINGS_WIDTH, self.SETTINGS_HEIGHT), modal=True, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
        
        # Event Loop
        while True:
            event, values = window.read()
            if event == self.EXIT_BUTTON or event == sg.WIN_CLOSED:
                break
            if event == 'Randomise':
                window[self.THEME_COMBOBOX].update(value='')
                theme = ''
            if event == self.SAVE_BUTTON:
                theme = values[self.THEME_COMBOBOX]
                emailMode = values['Email Mode']

                self.config.set('Preferences', 'Theme', theme)
                self.config.set('Preferences', 'Email-Mode', emailMode)
                
                with open(self.SETTINGS_PATH, 'w') as configfile:
                    self.config.write(configfile)
                break
            
        window.close()

    def selectedTemplateWindow(self, isNewTemplate: bool, name: str) -> str:
        with open(self.TEMPLATES_PATH, 'r') as f:
            jsonData = json.load(f)
            f.close()
            
        recipientDefault = ''
        recipientDisabled = False
        numberDefault = ''
        costDefault = ''
        instrumentDefault = ''
        dayDefault = ''
        studentDefault = ''
        
        if not isNewTemplate:
            recipientDefault = name
            recipientDisabled = True
            numberDefault = jsonData[name][self.NUMBER_INPUT]
            costDefault = jsonData[name][self.COST_INPUT]   
            instrumentDefault = jsonData[name][self.INSTRUMENT_INPUT]
            dayDefault = jsonData[name][self.DAY_INPUT]
            studentDefault = jsonData[name][self.STUDENT_INPUT] 
        
        layout = [
                    [
                        sg.Text('Recipient', font=self.textFont, pad=self.SELECT_PADDING), 
                        sg.Input(size=self.INPUT_SIZE*2, font=self.textFont, key=self.RECIPIENT_INPUT, default_text=recipientDefault, disabled=recipientDisabled, disabled_readonly_background_color='#FF6961')
                    ],
                    [
                        sg.Text('Number of lessons', font=self.textFont, pad=self.SELECT_PADDING), 
                        sg.Input(size=self.INPUT_SIZE, font=self.textFont, key=self.NUMBER_INPUT, default_text=numberDefault)
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
        
        window = sg.Window('', layout, element_justification='l', size=(self.SELECT_WIDTH, self.SELECT_HEIGHT), modal=True, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
        
        # Event Loop
        while True:
            event, values = window.read()
            if event == self.EXIT_BUTTON or event == sg.WIN_CLOSED:
                break
            if event == self.SAVE_BUTTON: 
                # Input error checking
                if re.search('\d', values[self.RECIPIENT_INPUT]):
                    sg.popup('Recipient name cannot contain numbers!', title='', font=self.textFont, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
                elif re.search('\D', values[self.NUMBER_INPUT]):
                    sg.popup('Number of lessons cannot contain characters!', title='', font=self.textFont, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
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
                            self.NUMBER_INPUT : values[self.NUMBER_INPUT],
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
                                sg.Button(self.NEW_BUTTON, font=self.textFont), 
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
                        sg.Text('<Email Recipient>', font=self.textFont, visible=draftEnabled, key='EmailRecipient'),
                        sg.Input(size=35, font=self.textFont, key='EmailInput', default_text=self.getEmailRecipient(), visible=draftEnabled)
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
                self.config.set('Preferences', 'Email-Recipient', window['EmailInput'].get())
                
                self.config.set('Preferences', 'Current-Template', window[self.NAMES_COMBOBOX].get())
                
                with open(self.SETTINGS_PATH, 'w') as configfile:
                    self.config.write(configfile)
                break
            if event == self.NAMES_COMBOBOX:
                if (not values[self.NAMES_COMBOBOX] == ''):
                    
                    self.toggleButtonsDisabled(window, self.templateEditButtonsList, False)
            if event == self.DELETE_BUTTON:
                choice = sg.popup_yes_no('Are you sure you want to delete this template?', title='', font=self.textFont, icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP)
                if choice == "Yes":
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
                
            if event == self.EDIT_BUTTON:
                self.selectedTemplateWindow(False, values[self.NAMES_COMBOBOX])
                
            if event == self.NEW_BUTTON:
                name = self.selectedTemplateWindow(True, '')
                
                with open(self.TEMPLATES_PATH, 'r') as f:
                    jsonData = json.load(f)
                    namesList = list(jsonData.keys())
                    namesList = sorted(namesList, key=str.lower)
                    
                window[self.NAMES_COMBOBOX].update(value=name, values=namesList)
                
            if event == self.DRAFT_BUTTON:
                gmail_create_draft(self.getEmailRecipient(), self.getSubject(values), self.getBody(values))
                
                sg.popup_quick_message('Draft Sent!', font=self.textFont, title='', icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, background_color="Black", text_color="White")
                
            if event == self.DRAFT_ALL_BUTTON:
                for name in self.getNamesList():
                    gmail_create_draft(self.getEmailRecipient(), self.getSubject(values), self.getBody(values))
                
                sg.popup_quick_message('All Drafts Sent!', font=self.textFont, title='', icon=self.BLANK_ICO, keep_on_top=self.KEEP_ON_TOP, background_color="Black", text_color="White")
                
            if event == self.SUBJECT_BUTTON:
                pyperclip.copy(self.getSubject(values))
                
            if event == self.BODY_BUTTON:
                pyperclip.copy(self.getBody(values))

            if event == self.SETTINGS_BUTTON:
                self.settingsWindow()
                
                sg.theme(self.getTheme())
                window.close()
                self.run()
                
                if (self.config.get('Preferences', 'Email-Mode') == self.CLIPBOARD):
                    window[self.DRAFT_BUTTON].update(visible=False)
                    window[self.DRAFT_ALL_BUTTON].update(visible=False)
                    window[self.SUBJECT_BUTTON].update(visible=True)
                    window[self.BODY_BUTTON].update(visible=True)
                    window['EmailRecipient'].update(visible=False)
                    window['EmailInput'].update(visible=False)
                else:
                    window[self.DRAFT_BUTTON].update(visible=True)
                    window[self.DRAFT_ALL_BUTTON].update(visible=True)
                    window[self.SUBJECT_BUTTON].update(visible=False)
                    window[self.BODY_BUTTON].update(visible=False)
                    window['EmailRecipient'].update(visible=True)
                    window['EmailInput'].update(visible=True)

        window.close()