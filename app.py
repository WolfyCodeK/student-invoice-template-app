import PySimpleGUI as sg
import os.path
import json
import re

"""
    [sg.Text('Enter student name'), sg.InputText()],
            [sg.Button('Ok'), sg.Button('Cancel')],
            
    sg.Button('', button_color=(sg.theme_background_color(), sg.theme_background_color()), image_filename='res\cog.png', size=(512, 512), image_subsample=15, border_width=0)
    
     sg.set_options(suppress_raise_key_errors=False, suppress_error_popups=True, suppress_key_guessing=True)
"""

# All the stuff inside your window.
titleFont = ("Courier New", 15)
textFont = ("Courier New", 12)

PADDING = 15

# Elements IDs
EDIT_BUTTON = 'Edit'
EXIT_BUTTON = 'Exit'
NAMES_COMBOBOX = 'Names'
INPUT_SIZE = 15
NEW_BUTTON = 'New Template'
SAVE_BUTTON = 'Save & Close'

def selectedTemplateWindow():
    layout = [[sg.Text('Recipient', font=textFont), sg.Input(size=INPUT_SIZE*2, font=textFont, key=('Recipient'))],
              [sg.Text('Number of lessons', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=('Number'))],
              [sg.Text('Cost of lesson  £', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont, key=('Cost'))], 
              [sg.VPush()],
              [sg.Button(SAVE_BUTTON, font=textFont), sg.Push(), sg.Button(EXIT_BUTTON, font=textFont)]
              ]
    sg.Titlebar()
    
    window = sg.Window('', layout, element_justification='c', size=(300, 160), modal=True, disable_close=True, icon='res\Blank.ico')
    
    
    while True:
        event, values = window.read()
        if event == EXIT_BUTTON:
            break
        if event == SAVE_BUTTON: 
            if re.search('\d', values['Recipient']):
                sg.popup('Recipient name cannot contain numbers!', title='', font=textFont, icon='res\Blank.ico')
            elif re.search('\D', values['Number']):
                sg.popup('Number of lessons cannot contain characters!', title='', font=textFont, icon='res\Blank.ico')
            elif re.search('\D', str(values['Cost']).replace('.', '')):
                sg.popup('Cost of lesson cannot contain characters!', title='', font=textFont, icon='res\Blank.ico')
            elif (len(str(values['Recipient'])) or len(str(values['Number'])) or len(str(values['Cost']))) == 0:
                sg.popup('All fields must be completed!', title='', font=textFont, icon='res\Blank.ico')
            else:
                with open(TEMPLATES_PATH, 'r') as f:
                    jsonData = json.load(f)
                    print(jsonData[0])
                    data = ''
                    data = data + values['Recipient']
                    data = data + values['Number']  
                    data = data + '%.2f' % int(values['Cost'])
                    print(data)
                break
        
    
    window.close()

def mainWindow():
    supportButtons = [
                        [sg.Button(NEW_BUTTON, font=textFont), 
                        sg.Button('Settings', font=textFont),
                        sg.Button(EXIT_BUTTON, font=textFont)]       
                    ]

    layout = [  
            [
                [
                sg.Text('Templates', font=titleFont, pad=PADDING),
                sg.Combo('', enable_events=True, default_value="", pad=PADDING, key=NAMES_COMBOBOX, size=15, font=textFont, readonly=True),
                sg.Button(EDIT_BUTTON, font=textFont, disabled=True),
                sg.Button('Subject', font=textFont),
                sg.Button('Body', font=textFont)
                ]],
                [sg.VPush()],
                [sg.Column(supportButtons, element_justification='right',expand_x=True)]
                
                
            ]

    window = sg.Window('Invoice Templates', layout, element_justification='l', size=(600, 225), icon='res\Moneybox.ico')
    
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == EXIT_BUTTON: # if user closes window or clicks cancel
            break
        if event == NAMES_COMBOBOX:
            if (not values[NAMES_COMBOBOX] == ''):
                window[EDIT_BUTTON].update(disabled=False)
        if event == EDIT_BUTTON:
            selectedTemplateWindow()
        if event == NEW_BUTTON:
            selectedTemplateWindow()

    window.close()

if __name__ == "__main__":
    
    TEMPLATES_PATH = 'templates.json'
    data = ''
    
    if (os.path.isfile(TEMPLATES_PATH)):
        with open(TEMPLATES_PATH, 'r') as f:
            data = f.read()
            f.close()
    else:
        f = open(TEMPLATES_PATH, 'w')

    sg.theme('DarkAmber')

    mainWindow()
    
    