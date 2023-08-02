import PySimpleGUI as sg
import os.path
import json
import re
import pyperclip

# Custom Fonts
titleFont = ("Courier New", 15)
textFont = ("Courier New", 12)

# File Paths
TEMPLATES_PATH = 'templates.json'

# Resource Paths
BLANK_ICO = 'res\Blank.ico'
MAIN_ICO = 'res\Moneybox.ico'

# Window Sizes
MAIN_WIDTH = 600
MAIN_HEIGHT = 225
SELECT_WIDTH = 300
SELECT_HEIGHT = 275

# Element Sizes
PADDING = 15

# Button Values
EDIT_BUTTON = 'Edit'
EXIT_BUTTON = 'Exit'
NEW_BUTTON = 'New Template'
SAVE_BUTTON = 'Save & Close'

# Combo Values
NAMES_COMBOBOX = 'Names'

# Input Values
RECIPIENT_INPUT = 'Recipient'
NUMBER_INPUT = 'Number'
COST_INPUT = 'Cost'
INSTRUMENT_INPUT = 'INSTRUMENT'
STUDENT_INPUT = 'Student'
INPUT_SIZE = 15

instrumentsList = ['Piano', 'Drum', 'Guitar', 'Vocal']

def checkSelectFieldsAreNotEmpty(values):
    return len(str(values[RECIPIENT_INPUT])) == 0 or len(str(values[NUMBER_INPUT])) == 0 or len(str(values[COST_INPUT])) == 0 or len(str(values[INSTRUMENT_INPUT])) == 0 or len(str(values[STUDENT_INPUT])) == 0

def selectedTemplateWindow(isNewTemplate, name):
    with open(TEMPLATES_PATH, 'r') as f:
        jsonData = json.load(f)
        
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
    
    layout = [[sg.Text('Recipient', font=textFont), sg.Input(size=INPUT_SIZE*2, font=textFont, key=RECIPIENT_INPUT,
            default_text=recipientDefault, disabled=recipientDisabled, disabled_readonly_background_color='DarkRed')],
            [sg.Text('Number of lessons', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=NUMBER_INPUT, default_text=numberDefault)],
            [sg.Text('Cost of lesson  £', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=COST_INPUT, default_text=costDefault)],
            [sg.Text('Instrument', font=textFont), sg.InputCombo(values=sorted(instrumentsList), size=INPUT_SIZE, font=textFont, key=INSTRUMENT_INPUT, default_value=instrumentDefault)],
            [sg.Text('Student(s)', font=textFont)], 
            [sg.Multiline(size=(INPUT_SIZE*2, 3), font=textFont, key=STUDENT_INPUT, default_text=studentDefault)],
            [sg.VPush()],
            [sg.Button(SAVE_BUTTON, font=textFont), sg.Push(), sg.Button(EXIT_BUTTON, font=textFont)]
            ]
    
    window = sg.Window('', layout, element_justification='c', size=(SELECT_WIDTH, SELECT_HEIGHT), modal=True, disable_close=True, icon=BLANK_ICO)
    
    while True:
        event, values = window.read()
        if event == EXIT_BUTTON:
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
                    
                    break
    f.close();  
    window.close()

def mainWindow():
    with open(TEMPLATES_PATH, 'r') as f:
        jsonData = json.load(f)
        namesList = list(jsonData.keys())
        namesList = sorted(namesList, key=str.lower)
    
    supportButtons = [
                        [sg.Button(NEW_BUTTON, font=textFont), 
                        sg.Button('Settings', font=textFont),
                        sg.Button(EXIT_BUTTON, font=textFont)]       
                    ]

    layout = [  
            [
                [
                sg.Text('Templates', font=titleFont, pad=PADDING),
                sg.Combo(namesList, enable_events=True, default_value="", pad=PADDING, key=NAMES_COMBOBOX, size=15, font=textFont, readonly=True),
                sg.Button(EDIT_BUTTON, font=textFont, disabled=True),
                sg.Button('Subject', font=textFont, disabled=True),
                sg.Button('Body', font=textFont, disabled=True)
                ]],
                [sg.VPush()],
                [sg.Column(supportButtons, element_justification='right',expand_x=True)]
            ]

    window = sg.Window('Invoice Templates', layout, element_justification='l', size=(MAIN_WIDTH, MAIN_HEIGHT), icon=MAIN_ICO)
    
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == EXIT_BUTTON: # if user closes window or clicks cancel
            break
        if event == NAMES_COMBOBOX:
            if (not values[NAMES_COMBOBOX] == ''):
                
                window[EDIT_BUTTON].update(disabled=False)
                window['Subject'].update(disabled=False)
                window['Body'].update(disabled=False)
                
        if event == EDIT_BUTTON:
            selectedTemplateWindow(False, values[NAMES_COMBOBOX])
        if event == NEW_BUTTON:
            selectedTemplateWindow(True, '')
            
            window[EDIT_BUTTON].update(disabled=True)
            window['Subject'].update(disabled=True)
            window['Body'].update(disabled=True)
            
            with open(TEMPLATES_PATH, 'r') as f:
                jsonData = json.load(f)
                namesList = list(jsonData.keys())
                namesList = sorted(namesList, key=str.lower)
                
            window[NAMES_COMBOBOX].update(values=namesList)
        if event == 'Subject':
            print(jsonData[values[NAMES_COMBOBOX]][NUMBER_INPUT])
        if event == 'Body':
            numberOfLessons = str(jsonData[values[NAMES_COMBOBOX]][NUMBER_INPUT])
            costOfLessons = str(jsonData[values[NAMES_COMBOBOX]][COST_INPUT])
            totalCost = str(int(numberOfLessons) * float(costOfLessons))
            totalCost = '%.2f' % (round(float(totalCost), 2))
            insturment = str(jsonData[values[NAMES_COMBOBOX]][INSTRUMENT_INPUT])
            students = str(jsonData[values[NAMES_COMBOBOX]][STUDENT_INPUT])
            pyperclip.copy("""Hi """ + str(values[NAMES_COMBOBOX]) +  """,

Here is my invoice for """ + students + """'s """ + insturment + """ lessons 2nd half summer term 2023.
--------
There are """ + numberOfLessons + """ sessions this 2nd half summer term from Tuesday 6th June to and including Tuesday 18th July.
 
""" + numberOfLessons + """ x £""" + costOfLessons + """ = £""" + totalCost + """

Thank you
--------

Kind regards

Robert""")
        if event == 'Settings':
            sg.popup('Work in progress')
            

    window.close()

if __name__ == "__main__":    
    if (not os.path.isfile(TEMPLATES_PATH)):
        f = open(TEMPLATES_PATH, 'w')

    sg.theme('DarkAmber')

    mainWindow()