import PySimpleGUI as sg
import os.path
import json
import re
import pyperclip
from datetime import date
from configparser import ConfigParser

# Theme List
themeList = ['Black', 'BlueMono', 'BluePurple', 'BrightColors', 'BrownBlue', 'Dark', 'Dark2', 'DarkAmber', 'DarkBlack', 'DarkBlack1', 'DarkBlue', 'DarkBlue1', 'DarkBlue10', 'DarkBlue11', 'DarkBlue12', 'DarkBlue13', 'DarkBlue14', 'DarkBlue15', 'DarkBlue16', 'DarkBlue17', 'DarkBlue2', 'DarkBlue3', 'DarkBlue4', 'DarkBlue5', 'DarkBlue6', 'DarkBlue7', 'DarkBlue8', 'DarkBlue9', 'DarkBrown', 'DarkBrown1', 'DarkBrown2', 'DarkBrown3', 'DarkBrown4', 'DarkBrown5', 'DarkBrown6', 'DarkGreen', 'DarkGreen1', 'DarkGreen2', 'DarkGreen3', 'DarkGreen4', 'DarkGreen5', 'DarkGreen6', 'DarkGrey', 'DarkGrey1', 'DarkGrey2', 'DarkGrey3', 'DarkGrey4', 'DarkGrey5', 'DarkGrey6', 'DarkGrey7', 'DarkPurple', 'DarkPurple1', 'DarkPurple2', 'DarkPurple3', 'DarkPurple4', 'DarkPurple5', 'DarkPurple6', 'DarkRed', 'DarkRed1', 'DarkRed2', 'DarkTanBlue', 'DarkTeal', 'DarkTeal1', 'DarkTeal10', 'DarkTeal11', 'DarkTeal12', 'DarkTeal2', 'DarkTeal3', 'DarkTeal4', 'DarkTeal5', 'DarkTeal6', 'DarkTeal7', 'DarkTeal8', 'DarkTeal9', 'Default', 'Default1', 'DefaultNoMoreNagging', 'Green', 'GreenMono', 'GreenTan', 'HotDogStand', 'Kayak', 'LightBlue', 'LightBlue1', 'LightBlue2', 'LightBlue3', 'LightBlue4', 'LightBlue5', 'LightBlue6', 'LightBlue7', 'LightBrown', 'LightBrown1', 'LightBrown10', 'LightBrown11', 'LightBrown12', 'LightBrown13', 'LightBrown2', 'LightBrown3', 'LightBrown4', 'LightBrown5', 'LightBrown6', 'LightBrown7', 'LightBrown8', 'LightBrown9', 'LightGray1', 'LightGreen', 'LightGreen1', 'LightGreen10', 'LightGreen2', 'LightGreen3', 'LightGreen4', 'LightGreen5', 'LightGreen6', 'LightGreen7', 'LightGreen8', 'LightGreen9', 'LightGrey', 'LightGrey1', 'LightGrey2', 'LightGrey3', 'LightGrey4', 'LightGrey5', 'LightGrey6', 'LightPurple', 'LightTeal', 'LightYellow', 'Material1', 'Material2', 'NeutralBlue', 'Purple', 'Reddit', 'Reds', 'SandyBeach', 'SystemDefault', 'SystemDefault1', 'SystemDefaultForReal', 'Tan', 'TanBlue', 'TealMono', 'Topanga']

# Term Dates


# Custom Fonts
titleFont = ("Courier New", 15)
textFont = ("Courier New", 12)

# File Paths
TEMPLATES_PATH = 'templates.json'
SETTINGS_PATH = 'settings.ini'

# Resource Paths
BLANK_ICO = 'res\Blank.ico'
MAIN_ICO = 'res\Moneybox.ico'

# Window Sizes
MAIN_WIDTH = 600
MAIN_HEIGHT = 225
SELECT_WIDTH = 300
SELECT_HEIGHT = 275
SETTINGS_WIDTH = 300
SETTINGS_HEIGHT = 100

# Element Sizes
PADDING = 15

# Button Values
EDIT_BUTTON = 'Edit'
EXIT_BUTTON = 'Exit'
NEW_BUTTON = 'New Template'
SAVE_BUTTON = 'Save & Close'

# Combo Values
NAMES_COMBOBOX = 'Names'
THEME_COMBOBOX = 'Theme'

# Input Values
RECIPIENT_INPUT = 'Recipient'
NUMBER_INPUT = 'Number'
COST_INPUT = 'Cost'
INSTRUMENT_INPUT = 'INSTRUMENT'
STUDENT_INPUT = 'Student'
INPUT_SIZE = 15

instrumentsList = ['piano', 'drum', 'guitar', 'vocal']

def checkSelectFieldsAreNotEmpty(values):
    return len(str(values[RECIPIENT_INPUT])) == 0 or len(str(values[NUMBER_INPUT])) == 0 or len(str(values[COST_INPUT])) == 0 or len(str(values[INSTRUMENT_INPUT])) == 0 or len(str(values[STUDENT_INPUT])) == 0

def settingsWindow():
    layout = [
                [
                    sg.Text(THEME_COMBOBOX, font=titleFont), 
                    sg.Combo(themeList, font=textFont, size=INPUT_SIZE, readonly=True, key=THEME_COMBOBOX)
                ],
                [
                    sg.VPush()
                ],
                [
                    sg.Button(SAVE_BUTTON, font=textFont),
                    sg.Push(), 
                    sg.Button(EXIT_BUTTON, font=textFont)
                ]
        ]
    
    window = sg.Window('', layout, element_justification='c', size=(SELECT_WIDTH, SETTINGS_HEIGHT), modal=True, icon=BLANK_ICO)
    
    # Event Loop
    while True:
        event, values = window.read()
        if event == EXIT_BUTTON or event == sg.WIN_CLOSED:
            break
        if event == SAVE_BUTTON:
            config = ConfigParser()
            config.read(SETTINGS_PATH)
            config.set('Preferences', 'theme', values[THEME_COMBOBOX])
            
            with open(SETTINGS_PATH, 'w') as configfile:
                config.write(configfile)
        break
        
    window.close()

def selectedTemplateWindow(isNewTemplate, name):
    with open(TEMPLATES_PATH, 'r') as f:
        jsonData = json.load(f)
        f.close()
        
    recipientDefault = ''
    recipientDisabled = False
    numberDefault = ''
    costDefault = ''
    instrumentDefault = ''
    studentDefault = ''
    
    if (isNewTemplate == False):
        recipientDefault = name
        recipientDisabled = True
        numberDefault = jsonData[name][NUMBER_INPUT]
        costDefault = jsonData[name][COST_INPUT]   
        instrumentDefault = jsonData[name][INSTRUMENT_INPUT]
        studentDefault = jsonData[name][STUDENT_INPUT] 
    
    layout = [
                [
                    sg.Text('Recipient', font=textFont), 
                    sg.Input(size=INPUT_SIZE*2, font=textFont, key=RECIPIENT_INPUT, default_text=recipientDefault, disabled=recipientDisabled, disabled_readonly_background_color='DarkRed')
                ],
                [
                    sg.Text('Number of lessons', font=textFont), 
                    sg.Input(size=INPUT_SIZE, font=textFont, key=NUMBER_INPUT, default_text=numberDefault)
                ],
                [
                    sg.Text('Cost of lesson  £', font=textFont), 
                    sg.Input(size=INPUT_SIZE, font=textFont, key=COST_INPUT, default_text=costDefault)
                ],
                [   sg.Text('Instrument', font=textFont), 
                    sg.Combo(values=sorted(instrumentsList), size=INPUT_SIZE,
                    font=textFont, key=INSTRUMENT_INPUT, default_value=instrumentDefault, readonly=True)
                ],
                [   
                    sg.Text('Student(s)', font=textFont)
                ], 
                [
                    sg.Multiline(size=(INPUT_SIZE*2, 3), font=textFont, key=STUDENT_INPUT, default_text=studentDefault)
                ],
                [
                    sg.VPush()
                ],
                [
                    sg.Button(SAVE_BUTTON, font=textFont), 
                    sg.Push(), 
                    sg.Button(EXIT_BUTTON, font=textFont)
                ]
            ]
    
    window = sg.Window('', layout, element_justification='c', size=(SELECT_WIDTH, SELECT_HEIGHT), modal=True, icon=BLANK_ICO)
    
    # Event Loop
    while True:
        event, values = window.read()
        if event == EXIT_BUTTON or event == sg.WIN_CLOSED:
            break
        if event == SAVE_BUTTON: 
            # Input error checking
            if re.search('\d', values[RECIPIENT_INPUT]):
                sg.popup('Recipient name cannot contain numbers!', title='', font=textFont, icon=BLANK_ICO)
            elif re.search('\D', values[NUMBER_INPUT]):
                sg.popup('Number of lessons cannot contain characters!', title='', font=textFont, icon=BLANK_ICO)
            elif re.search('\D', str(values[COST_INPUT]).replace('.', '')):
                sg.popup('Cost of lesson cannot contain characters!', title='', font=textFont, icon=BLANK_ICO)
            elif re.search('\d', values[STUDENT_INPUT]):
                sg.popup('Students names cannot contain numbers!', title='', font=textFont, icon=BLANK_ICO)
            elif checkSelectFieldsAreNotEmpty(values):
                sg.popup('All fields must be completed!', title='', font=textFont, icon=BLANK_ICO)
            else:
                if (name in jsonData) and (isNewTemplate):
                    sg.popup('Template with that name already exists!', title='', font=textFont, icon=BLANK_ICO)
                else:
                    name = values[RECIPIENT_INPUT]
                    
                    info = {
                        NUMBER_INPUT : values[NUMBER_INPUT],
                        COST_INPUT : '%.2f' % (round(float(values[COST_INPUT]), 2)),
                        INSTRUMENT_INPUT : values[INSTRUMENT_INPUT],
                        STUDENT_INPUT : values[STUDENT_INPUT]
                    }
                    
                    jsonData[name] = info

                    with open(TEMPLATES_PATH, 'w') as f:
                        f.write(json.dumps(jsonData))
                        f.close()
                        
                    break
            
    window.close()

def mainWindow():
    today = date.today()
    
    with open(TEMPLATES_PATH, 'r') as f:
        jsonData = json.load(f)
        namesList = list(jsonData.keys())
        namesList = sorted(namesList, key=str.lower)
        f.close()
    
    supportButtons = [
                        [
                            sg.Button('DELETE', font=textFont, disabled=True),
                            sg.Push(),
                            sg.Button(NEW_BUTTON, font=textFont), 
                            sg.Button('Settings', font=textFont),
                            sg.Button(EXIT_BUTTON, font=textFont)
                        ]       
                    ]

    layout = [  
                [
                    [
                        sg.Text('Templates', font=titleFont, pad=PADDING),
                        sg.Combo(namesList, enable_events=True, default_value="", pad=PADDING, key=NAMES_COMBOBOX, size=15, font=textFont, readonly=True),
                        sg.Button(EDIT_BUTTON, font=textFont, disabled=True),
                        sg.Button('Subject', font=textFont, disabled=True),
                        sg.Button('Body', font=textFont, disabled=True)
                    ]
                ],
                [
                    sg.VPush()
                ],
                [
                    sg.Column(supportButtons, element_justification='right',expand_x=True)
                ]
            ]

    window = sg.Window('Invoice Templates', layout, element_justification='l', size=(MAIN_WIDTH, MAIN_HEIGHT), icon=MAIN_ICO)
    
    # Event Loop
    while True:
        print("hello")
        
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == EXIT_BUTTON:
            break
        if event == NAMES_COMBOBOX:
            if (not values[NAMES_COMBOBOX] == ''):
                
                window[EDIT_BUTTON].update(disabled=False)
                window['Subject'].update(disabled=False)
                window['Body'].update(disabled=False)
                window['DELETE'].update(disabled=False)
        if event == 'DELETE':
            choice = sg.popup_yes_no('Are you sure you want to delete this template?', font=textFont)
            if choice == "Yes":
                with open(TEMPLATES_PATH, 'r') as f:
                    jsonData = json.load(f)
                    f.close()
                    
                del jsonData[values[NAMES_COMBOBOX]]   
                
                namesList = list(jsonData.keys())
                namesList = sorted(namesList, key=str.lower)
                window[NAMES_COMBOBOX].update(values=namesList) 
                
                with open(TEMPLATES_PATH, 'w') as f:
                    f.write(json.dumps(jsonData))   
                    f.close()          
                    
                window[EDIT_BUTTON].update(disabled=True)
                window['Subject'].update(disabled=True)
                window['Body'].update(disabled=True)
                window['DELETE'].update(disabled=True)
             
        if event == EDIT_BUTTON:
            selectedTemplateWindow(False, values[NAMES_COMBOBOX])
        if event == NEW_BUTTON:
            selectedTemplateWindow(True, '')
            
            window[EDIT_BUTTON].update(disabled=True)
            window['Subject'].update(disabled=True)
            window['Body'].update(disabled=True)
            window['DELETE'].update(disabled=True)
            
            with open(TEMPLATES_PATH, 'r') as f:
                jsonData = json.load(f)
                namesList = list(jsonData.keys())
                namesList = sorted(namesList, key=str.lower)
                
            window[NAMES_COMBOBOX].update(values=namesList)
        if event == 'Subject':
            insturment = str(jsonData[values[NAMES_COMBOBOX]][INSTRUMENT_INPUT])
            pyperclip.copy("""Invoice for """ + insturment + """ Lessons 2nd Half Summer Term """ + str(today.year))
            
        if event == 'Body':
            name = str(values[NAMES_COMBOBOX])
            numberOfLessons = str(jsonData[values[NAMES_COMBOBOX]][NUMBER_INPUT])
            costOfLessons = str(jsonData[values[NAMES_COMBOBOX]][COST_INPUT])
            totalCost = str(int(numberOfLessons) * float(costOfLessons))
            totalCost = '%.2f' % (round(float(totalCost), 2))
            insturment = str(jsonData[values[NAMES_COMBOBOX]][INSTRUMENT_INPUT])
            students = str(jsonData[values[NAMES_COMBOBOX]][STUDENT_INPUT])
            pyperclip.copy("""Hi """ + name +  """,

Here is my invoice for """ + students + """'s """ + insturment + """ lessons 2nd half summer term """ + str(today.year) + """.
--------
There are """ + numberOfLessons + """ sessions this 2nd half summer term from Tuesday 6th June to and including Tuesday 18th July.
 
""" + numberOfLessons + """ x £""" + costOfLessons + """ = £""" + totalCost + """

Thank you
--------

Kind regards

Robert""")
        if event == 'Settings':
            settingsWindow()

    window.close()

if __name__ == "__main__":    
    if (not os.path.isfile(TEMPLATES_PATH)):
        f = open(TEMPLATES_PATH, 'w')
        f.write('{}')
        f.close()
        
    if (not os.path.isfile(SETTINGS_PATH)):
        f = open(SETTINGS_PATH, 'w')
        f.write('[Preferences]\n')
        f.write('theme = DarkAmber')
        f.close()

    config = ConfigParser()
    config.read(SETTINGS_PATH)
    theme = config.get('Preferences', 'theme')
        
    sg.theme(theme)

    mainWindow()