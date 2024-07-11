import PySimpleGUI as sg
import os
from file_handler import process_file, convert_to_xml

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
            try:
                df = process_file(selected_file)
                output_file = os.path.splitext(selected_file)[0] + ".xlsx"
                convert_to_xml(df, output_file)
                sg.popup("File Uploaded!", f"File has been processed and saved as {output_file}")
                print(f"Uploading file: {selected_file}")
            except Exception as e:
                sg.popup("Error", str(e))
                print(f"Error: {str(e)}")
        else:
            sg.popup("No file selected!", "Please select a file to upload.")

window.close()