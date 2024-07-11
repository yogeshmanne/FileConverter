import PySimpleGUI as sg
import pandas as pd
import os

def process_file(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.csv':
        df = pd.read_csv(file_path)
    elif file_extension in ['.xls', '.xlsx']:
        df = pd.read_excel(file_path)
    elif file_extension == '.json':
        df = pd.read_json(file_path)
    elif file_extension == '.txt': 
        df = pd.read_csv(file_path, delimiter='\t')
    else:
        raise ValueError("Unsupported file format: Only CSV, Excel, JSON, and TXT files are supported.")
    
    return df

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
                output_file = os.path.splitext(selected_file)[0] + "_output.xlsx"
                df.to_excel(output_file, index=False)
                sg.popup("File Uploaded!", f"File has been processed and saved as {output_file}")
                print(f"Uploading file: {selected_file}")
            except Exception as e:
                sg.popup("Error", str(e))
                print(f"Error: {str(e)}")
        else:
            sg.popup("No file selected!", "Please select a file to upload.")

window.close()