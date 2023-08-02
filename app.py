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
SELECT_HEIGHT = 160

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
INPUT_SIZE = 15

def selectedTemplateWindow(isNewTemplate, name):
    with open(TEMPLATES_PATH, 'r') as f:
        jsonData = json.load(f)
    
    if (isNewTemplate):
        layout = [[sg.Text('Recipient', font=textFont), sg.Input(size=INPUT_SIZE*2, font=textFont, key=(RECIPIENT_INPUT))],
              [sg.Text('Number of lessons', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=(NUMBER_INPUT))],
              [sg.Text('Cost of lesson  £', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=(COST_INPUT))], 
              [sg.VPush()],
              [sg.Button(SAVE_BUTTON, font=textFont), sg.Push(), sg.Button(EXIT_BUTTON, font=textFont)]
              ]
    else:
        layout = [[sg.Text('Recipient', font=textFont), sg.Input(size=INPUT_SIZE*2, font=textFont, key=(RECIPIENT_INPUT), disabled=True, default_text=name, disabled_readonly_background_color='DarkRed')],
              [sg.Text('Number of lessons', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=(NUMBER_INPUT), default_text=jsonData[name][NUMBER_INPUT])],
              [sg.Text('Cost of lesson  £', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=(COST_INPUT), default_text=jsonData[name][COST_INPUT])], 
              [sg.VPush()],
              [sg.Button(SAVE_BUTTON, font=textFont), sg.Push(), sg.Button(EXIT_BUTTON, font=textFont)]
              ] 
    
    sg.Titlebar()
    
    window = sg.Window('', layout, element_justification='c', size=(SELECT_WIDTH, SELECT_HEIGHT), modal=True, disable_close=True, icon=BLANK_ICO)
    
    while True:
        event, values = window.read()
        if event == EXIT_BUTTON:
            break
        if event == SAVE_BUTTON: 
            if re.search('\d', values[RECIPIENT_INPUT]):
                sg.popup('Recipient name cannot contain numbers!', title='', font=textFont, icon=BLANK_ICO)
            elif re.search('\D', values[NUMBER_INPUT]):
                sg.popup('Number of lessons cannot contain characters!', title='', font=textFont, icon=BLANK_ICO)
            elif re.search('\D', str(values[COST_INPUT]).replace('.', '')):
                sg.popup('Cost of lesson cannot contain characters!', title='', font=textFont, icon=BLANK_ICO)
            elif len(str(values[RECIPIENT_INPUT])) == 0 or len(str(values[NUMBER_INPUT])) == 0 or len(str(values[COST_INPUT])) == 0:
                sg.popup('All fields must be completed!', title='', font=textFont, icon=BLANK_ICO)
            else:
                name = values['Recipient']
                
                if (name in jsonData) and (isNewTemplate == True):
                    sg.popup('Template with that name already exists!', title='', font=textFont, icon=BLANK_ICO)
                else:
                    info = {
                        'Number' : values['Number'],
                        'Cost' : '%.2f' % (round(float(values['Cost']), 2))
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
                
            window[NAMES_COMBOBOX].update(values=namesList)
        if event == 'Subject':
            print(jsonData[values[NAMES_COMBOBOX]][NUMBER_INPUT])
        if event == 'Body':
            numberOfLessons = str(jsonData[values[NAMES_COMBOBOX]][NUMBER_INPUT])
            costOfLessons = str(jsonData[values[NAMES_COMBOBOX]][COST_INPUT])
            totalCost = str(int(numberOfLessons) * float(costOfLessons))
            totalCost = '%.2f' % (round(float(totalCost), 2))
            pyperclip.copy("""Hi """ + str(values[NAMES_COMBOBOX]) +  """,

Here is my invoice for Melody's drum lessons 2nd half summer term 2023.
--------
There are """ + numberOfLessons + """ sessions this 2nd half summer term from Tuesday 6th June to and including Tuesday 18th July.
 
""" + numberOfLessons + """ x £""" + costOfLessons + """ = £""" + totalCost + """

Thank you
--------

Kind regards""")
        if event == 'Settings':
            sg.popup('Work in progress')
            

    window.close()

if __name__ == "__main__":    
    if (not os.path.isfile(TEMPLATES_PATH)):
        f = open(TEMPLATES_PATH, 'w')

    sg.theme('DarkAmber')

    mainWindow()