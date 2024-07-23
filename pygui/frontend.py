import PySimpleGUI as sg  # Import the PySimpleGUI library for creating GUI applications

def create_window():
    """
    Create and return the GUI window with the specified layout.
    
    The layout includes:
    - A title text
    - Input fields for name, email, and phone
    - Checkboxes for branch selection
    - Buttons for submit, clear, exit, and download
    - A dropdown menu for file format selection
    """
    layout = [
        # Title row with centered text
        [sg.Text('User Data Form ', size=(30, 1), justification='center', font=("TimesNewRoman", 25), text_color='lightgreen')],
        [sg.Text('', size=(15, 1))],  # Spacer row

        # Name input field
        [sg.Text('Name', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.InputText(key='name', font=("Helvetica", 15))],
        
        # Email input field
        [sg.Text('Email ID', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.InputText(key='email', font=("Helvetica", 15))],
        
        # Phone number input field
        [sg.Text('Phone Number', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.InputText(key='phone', font=("Helvetica", 15))],
        
        # Branch checkboxes
        [sg.Text('Branch', size=(15, 1), font=("Arial", 20), justification='center'), 
         sg.Checkbox('Engineering', key='Engineering', font=("Helvetica", 15)),
         sg.Checkbox('Medical', key='Medical', font=("Helvetica", 15)),
         sg.Checkbox('Degree', key='Degree', font=("Helvetica", 15))],
        
        [sg.Text('', size=(15, 1))],  # Spacer row
        
        # Buttons for submitting, clearing, and exiting
        [sg.Submit(button_color=('white', 'green'), font=("Helvetica", 15)),
         sg.Button('Clear', button_color=('white', 'blue'), font=("Helvetica", 15)), 
         sg.Exit(button_color=('white', 'red'), font=("Helvetica", 15))],
        
        [sg.Text('', size=(15, 1))],  # Spacer row
        
        # File format dropdown and download button
        [sg.Text('Download As:', size=(15, 1), font=("Helvetica", 15), justification='center'), 
         sg.Combo(['json', 'txt', 'xml', 'xlsx'], key='file_format', default_value='xlsx', font=("Helvetica", 15)), 
         sg.Button('Download', button_color=('white', 'orange'), font=("Helvetica", 15))],
    ]

    # Create the window with the title 'User Data Form', and finalize it to make it immediately visible
    return sg.Window('User Data Form', layout, resizable=True, finalize=True)

def clear_input(window, values):
    """
    Clear the input fields and checkboxes in the GUI.

    Args:
    window (sg.Window): The PySimpleGUI window object.
    values (dict): Dictionary of current values in the input fields.
    """
    # Clear text inputs (name, email, phone)
    for key in values:
        if key in ['name', 'email', 'phone']:
            window[key]('')  # Set the value of the input field to an empty string
    
    # Uncheck checkboxes (branches)
    for branch in ['Engineering', 'Medical', 'Degree']:
        window[branch](False)  # Set the value of the checkbox to False (unchecked)