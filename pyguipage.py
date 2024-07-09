import PySimpleGUI as sg
import os

# Define the layout of the window
layout = [
    [sg.Text("Select a file from your Downloads folder:")],
    [sg.Input(key="-FILE-", enable_events=True), sg.FileBrowse(initial_folder=os.path.expanduser("~/Downloads"))],
    [sg.Button("Upload"), sg.Button("Exit")]
]

# Create the window
window = sg.Window("File Uploader", layout)

# Event loop
while True:
    event, values = window.read()

    # If user closes window or clicks 'Exit'
    if event == sg.WINDOW_CLOSED or event == "Exit":
        break

    # If user selects a file
    if event == "-FILE-":
        selected_file = values["-FILE-"]
        print(f"Selected file: {selected_file}")

    # If user clicks 'Upload'
    if event == "Upload":
        selected_file = values["-FILE-"]
        if selected_file:
            # Here you can add your file handling code
            print(f"Uploading file: {selected_file}")
        else:
            sg.popup("No file selected!", "Please select a file to upload.")

# Close the window
window.close()