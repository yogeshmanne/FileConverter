import os
import json
import openpyxl
import xml.etree.ElementTree as ET
from openpyxl import Workbook
import PySimpleGUI as sg

# Excel File Path where data will append.
excel_file = r"C:\Users\T.Venkateshwar Reddy\OneDrive\Desktop\data.xlsx"

# Dictionary to keep track of the number of files created for each format
file_count = {
    'json': 0,
    'txt': 0,
    'xml': 0,
    'xlsx': 0,
}

def create_excel_file():
    if not os.path.exists(excel_file):  # Check if the file does not exist
        wb = Workbook()  
        ws = wb.active
        ws.title = "UserData" 
        ws.append(["Name", "Email ID", "Phone Number", "Branch"]) 
        wb.save(excel_file)  # Save the workbook to the specified path

def save_data_to_excel(name, email, phone, branch):
    wb = openpyxl.load_workbook(excel_file)  
    ws = wb.active  
    ws.append([name, email, phone, branch])  
    wb.save(excel_file)  #All the changes will be saved in the workbook ie; excel

def save_as_json(data, file_path):
    with open(file_path, 'w') as f:
        json.dump(data, f)  #JSON Format

def save_as_txt(data, file_path):
    with open(file_path, 'w') as f:
        for entry in data:
            f.write(f"{entry}\n")  #TXT Format

def save_as_xml(data, file_path):
    root = ET.Element("Users")  
    for entry in data:
        user = ET.SubElement(root, "User")  
        for key, value in entry.items():
            ET.SubElement(user, key).text = value  
    tree = ET.ElementTree(root)  
    tree.write(file_path)  #XML File Format

def download_data(file_format, data):
    global file_count
    file_path = f"document_{file_count[file_format]}.{file_format}"  # Generate file path with a count suffix
    if file_format == 'json':
        save_as_json(data, file_path)  # Save data as JSON
    elif file_format == 'txt':
        save_as_txt(data, file_path)  # Save data as text
    elif file_format == 'xml':
        save_as_xml(data, file_path)  # Save data as XML
    elif file_format == 'xlsx':
        wb = Workbook()
        ws = wb.active  
        ws.title = "UserData"  
        ws.append(["Name", "Email ID", "Phone Number", "Branch"])  
        for entry in data:
            ws.append([entry['Name'], entry['Email'], entry['Phone'], entry['Branch']])  
        wb.save(file_path) 

    sg.popup(f'Downloaded as {file_format}', title='Download')  # Show a popup indicating the file has been downloaded
    file_count[file_format] += 1  # Increment the count for the specified file format