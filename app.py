import PySimpleGUI as sg

sg.theme('DarkGrey6')   # Add a touch of color

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


def selectedTemplateWindow():
    layout = [[sg.Text('Recipient', font=textFont), sg.Input(size=INPUT_SIZE*2, font=textFont)],
              [sg.Text('Number of lessons', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont)],
              [sg.Text('Cost of lesson  £', font=textFont), sg.Input(size=INPUT_SIZE, font=textFont)], 
              [sg.VPush()],
              [sg.Button('Save & Close', font=textFont), sg.Push(), sg.Button(EXIT_BUTTON, font=textFont)]
              ]
    sg.Titlebar()
    
    window = sg.Window('', layout, element_justification='c', size=(400, 225), modal=True,
            use_custom_titlebar=True, disable_close=True)
    
    while True:
        event, values = window.read()
        if event == EXIT_BUTTON:
            break
        if event == 'Save & Close':
            print('test')
            break
        
    
    window.close()

def mainWindow():
    supportButtons = [
                        [sg.Button('New Template', font=textFont), 
                        sg.Button('Settings', font=textFont),
                        sg.Button(EXIT_BUTTON, font=textFont)]       
                    ]

    layout = [  
            [
                [
                sg.Text('Templates', font=titleFont, pad=PADDING),
                sg.Combo('Jane', enable_events=True, default_value="", key=NAMES_COMBOBOX, size=15, font=textFont, readonly=True),
                sg.Button(EDIT_BUTTON, font=textFont, pad=PADDING, disabled=True)],
                sg.VPush()
                ],
                    
                    [sg.Column(supportButtons, element_justification='right',expand_x=True)]
                
                
            ]

    # Create the Window
    window = sg.Window('Invoice Templates', layout, element_justification='l', size=(600, 225))
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

    window.close()

if __name__ == "__main__":
    mainWindow()