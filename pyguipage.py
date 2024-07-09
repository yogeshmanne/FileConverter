import PySimpleGUI as sg
import os


layout = [
    [sg.Text("Select a file from your Downloads folder:")],
    [sg.Input(key="-FILE-", enable_events=True), sg.FileBrowse(initial_folder=os.path.expanduser("~/Downloads"))],
    [sg.Button("Upload"), sg.Button("Exit")]
]


window = sg.Window("File Uploader", layout)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == "Exit":
        break

    if event == "-FILE-":
        selected_file = values["-FILE-"]
        print(f"Selected file: {selected_file}")

    if event == "Upload":
        selected_file = values["-FILE-"]
        if selected_file:
            print(f"Uploading file: {selected_file}")
        else:
            sg.popup("No file selected!", "Please select a file to upload.")

window.close()