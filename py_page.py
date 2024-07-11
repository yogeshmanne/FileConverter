
import PySimpleGUI as sg
import pandas as pd

sg.theme('DarkTeal9')


EXCEL_FILE = "https://docs.google.com/spreadsheets/d/1Q6g4Va8kPOOMYSKf4FwEJ9zqvo2wzSivh2rZrZ9UkOw/edit?gid=0#gid=0"
df=pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Enter Student Details :')],
    [sg.Text('Name of the Student', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('Enrollment Number ', size=(15,1)), sg.InputText(key='Enrollment')],
    [sg.Text('Lnaguages Known', size=(15,1)),
                            sg.Checkbox('C', key='C'),
                            sg.Checkbox('Python', key='Python'),
                            sg.Checkbox('Java', key='Java')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False) 
        sg.popup('Data saved!')
        clear_input()
window.close()
