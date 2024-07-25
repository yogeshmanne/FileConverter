import PySimpleGUI as sg

def create_window():
    sg.theme('LightBrown8')  # Set a visually appealing theme

    layout = [
        [sg.Text('User Data Form', size=(30, 1), justification='center', font=("Times New Roman", 25), text_color='lightgreen')],
        [sg.Text('', size=(1, 1))],
        [sg.Text('Name', size=(15, 1), font=("Times New Roman", 20), justification='center'),
         sg.InputText(key='name', font=("Times New Roman", 15))],
        [sg.Text('Email ID', size=(15, 1), font=("Times New Roman", 20), justification='center'),
         sg.InputText(key='email', font=("Times New Roman", 15))],
        [sg.Text('Phone Number', size=(15, 1), font=("Times New Roman", 20), justification='center'),
         sg.InputText(key='phone', font=("Times New Roman", 15))],
        [sg.Text('Branch', size=(15, 1), font=("Times New Roman", 20), justification='center'),
         sg.Checkbox('Engineering', key='Engineering', font=("Times New Roman", 15)),
         sg.Checkbox('Medical', key='Medical', font=("Times New Roman", 15)),
         sg.Checkbox('Degree', key='Degree', font=("Times New Roman", 15))],

        [sg.Text('', size=(1, 1))],

        # Designing Buttons
        [sg.Submit(button_color=('white', 'green'), font=("Times New Roman", 15)),
         sg.Button('Clear', button_color=('white', 'blue'), font=("Times New Roman", 15)),
         sg.Exit(button_color=('white', 'red'), font=("Times New Roman", 15))],

        [sg.Text('', size=(1, 1))],

        # File format dropdown and download button
        [sg.Text('Download As:', size=(15, 1), font=("Times New Roman", 15), justification='center'),
         sg.Combo(['json', 'txt', 'xml', 'xlsx'], key='file_format', default_value='xlsx', font=("Times New Roman", 15)),
         sg.Button('Download', button_color=('white', 'orange'), font=("Times New Roman", 15))],
    ]

    return sg.Window('User Data Form', layout, resizable=True, finalize=True)

def clear_input(window, values):
    for key in values:
        if key in ['name', 'email', 'phone']:
            window[key]('')
    for branch in ['Engineering', 'Medical', 'Degree']:
        window[branch](False)  # Set the value of the checkbox to False (unchecked)