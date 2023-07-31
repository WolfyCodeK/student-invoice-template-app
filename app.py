import PySimpleGUI as sg

#sg.theme('DarkAmber')   # Add a touch of color

"""
    [sg.Text('Enter student name'), sg.InputText()],
            [sg.Button('Ok'), sg.Button('Cancel')],
            
    sg.Button('', button_color=(sg.theme_background_color(), sg.theme_background_color()), image_filename='res\cog.png', size=(512, 512), image_subsample=15, border_width=0)
"""

# All the stuff inside your window.
titleFont = ("Courier New", 14)
textFont = ("Courier New", 11)

PADDING = 15

supportButtons = [
                    [sg.Button('New Template', font=textFont), 
                     sg.Button('Settings', font=textFont)]       
                ]

layout = [  
          [
            [
            sg.Text('Students', font=titleFont, pad=PADDING),
            sg.Combo('', default_value="", size=15, font=textFont, readonly=True),
            sg.Button('Select', font=textFont, pad=PADDING)],
            sg.VPush()
            ],
                
                [sg.Column(supportButtons, element_justification='right',expand_x=True)]
            
            
        ]

# Create the Window
window = sg.Window('Student Invoice Templates', layout, element_justification='l', size=(600, 225))
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    print('You entered ', values[0])

window.close()