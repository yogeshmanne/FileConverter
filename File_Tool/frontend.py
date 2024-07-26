import PySimpleGUI as sg  

#this function is used to create the layout of the page.
def create_window():
    sg.theme('SandyBeach')  #adding backgroundtheme
    layout = [
        [sg.Text('User Data Form', size=(30, 1), justification='center', font=("Times New Roman", 30), text_color='brown')],
        [sg.Text('', size=(1, 1))],    #heading of the form 

        #form collects details such as Name , Email ID , Phone Number and Branch
        [sg.Text('Name', size=(15, 1), font=("Times New Roman", 20), justification='center'), 
         sg.InputText(key='name', font=("Times New Roman", 15))],  
        [sg.Text('Email ID', size=(15, 1), font=("Times New Roman", 20), justification='center'), 
         sg.InputText(key='email', font=("Times New Roman", 15))],
        [sg.Text('Phone Number', size=(15, 1), font=("Times New Roman", 20), justification='center'), 
         sg.InputText(key='phone', font=("Times New Roman", 15))],

         #Branch has checkbox containing options such as Engineering , Medical and Degree.
        [sg.Text('Branch', size=(15, 1), font=("Times New Roman", 20), justification='center'), 
         sg.Checkbox('Engineering', key='Engineering', font=("Times New Roman", 15)),
         sg.Checkbox('Medical', key='Medical', font=("Times New Roman", 15)),
         sg.Checkbox('Degree', key='Degree', font=("Times New Roman", 15))],
        
        [sg.Text('', size=(1, 1))],  
        
        #designing buttons - Submit , Clear and Exit Button
        [sg.Submit(button_color=('white', 'green'), font=("Times New Roman", 15)),   #Submit button with green and white color
         sg.Button('Clear', button_color=('white', 'blue'), font=("Times New Roman", 15)), #Clear button with white and blue color
         sg.Exit(button_color=('white', 'red'), font=("Times New Roman", 15))],  #exit button with white and red color
        
        [sg.Text('', size=(1, 1))], 
        
        #Drop down menu has been created for File Formats with options such as JSON , TXT , XML , XLSX
        [sg.Text('Download As:', size=(15, 1), font=("Times New Roman", 15), justification='center'), 
         sg.Combo(['json', 'txt', 'xml', 'xlsx'], key='file_format', default_value='xlsx', font=("Times New Roman", 15)), 
         sg.Button('Download', button_color=('white', 'orange'), font=("Times New Roman", 15))],  #Download button styling
    ]

    return sg.Window('User Data Form', layout, resizable=True, finalize=True)   #function return value


def clear_input(window, values):
    for key in values:
        if key in ['name', 'email', 'phone']:
            window[key]('') 
    for branch in ['Engineering', 'Medical', 'Degree']:
        window[branch](False)  # Set the value of the checkbox to be False (unchecked)