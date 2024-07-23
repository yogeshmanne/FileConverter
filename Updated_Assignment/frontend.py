import PySimpleGUI as sg  

def create_window():
    layout = [
        #create the layout for the window
        [sg.Text('User Data Form ', size=(30, 1), justification='center', font=("TimesNewRoman", 25), text_color='lightgreen')],
        [sg.Text('', size=(15, 1))],  
        [sg.Text('Name', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.InputText(key='name', font=("Helvetica", 15))],
        [sg.Text('Email ID', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.InputText(key='email', font=("Helvetica", 15))],
        [sg.Text('Phone Number', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.InputText(key='phone', font=("Helvetica", 15))],
        [sg.Text('Branch', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.Checkbox('Engineering', key='Engineering', font=("Helvetica", 15)),
         sg.Checkbox('Medical', key='Medical', font=("Helvetica", 15)),
         sg.Checkbox('Degree', key='Degree', font=("Helvetica", 15))],
        
        [sg.Text('', size=(15, 1))],  
        
        # Designing Buttons
        [sg.Submit(button_color=('white', 'green'), font=("Helvetica", 15)),
         sg.Button('Clear', button_color=('white', 'blue'), font=("Helvetica", 15)), 
         sg.Exit(button_color=('white', 'red'), font=("Helvetica", 15))],
        
        [sg.Text('', size=(15, 1))], 
        
        # File format dropdown and download button
        [sg.Text('Download As:', size=(15, 1), font=("Helvetica", 15), justification='center'), 
         sg.Combo(['json', 'txt', 'xml', 'xlsx'], key='file_format', default_value='xlsx', font=("Helvetica", 15)), 
         sg.Button('Download', button_color=('white', 'orange'), font=("Helvetica", 15))],
    ]

    return sg.Window('User Data Form', layout, resizable=True, finalize=True)

def clear_input(window, values):
    for key in values:
        if key in ['name', 'email', 'phone']:
            window[key]('') 
    for branch in ['Engineering', 'Medical', 'Degree']:
        window[branch](False)  # Set the value of the checkbox to False (unchecked)