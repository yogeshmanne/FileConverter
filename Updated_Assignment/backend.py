import os
import json
import openpyxl
import xml.etree.ElementTree as ET
from openpyxl import Workbook
import PySimpleGUI as sg

# Define the Excel file path
current_dir = os.getcwd()
excel_file = os.path.join(current_dir, "data.xlsx")

# Dictionary to keep track of the number of files created for each format
file_count = {fmt: 0 for fmt in ['json', 'txt', 'xml', 'xlsx']}


def save_data_to_excel(name, email, phone, branch):
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active
    ws.append([name, email, phone, branch])
    wb.save(excel_file)

def save_as_file(data, file_path, fmt):
    if fmt == 'json':
        with open(file_path, 'w') as f:
            json.dump(data, f)
    elif fmt == 'txt':
        with open(file_path, 'w') as f:
            f.write("\n".join(data))
    elif fmt == 'xml':
        root = ET.Element("Users")
        for entry in data:
            user = ET.SubElement(root, "User")
            for k, v in entry.items():
                ET.SubElement(user, k).text = v
        ET.ElementTree(root).write(file_path)
    elif fmt == 'xlsx':
        wb = Workbook()
        ws = wb.active
        ws.title = "UserData"
        ws.append(["Name", "Email ID", "Phone Number", "Branch"])
        for entry in data:
            ws.append([entry['Name'], entry['Email'], entry['Phone'], entry['Branch']])
        wb.save(file_path)

def download_data(file_format, data):
    global file_count
    file_path = os.path.join(current_dir, f"document_{file_count[file_format]}.{file_format}")
    save_as_file(data, file_path, file_format)
    sg.popup(f'Downloaded as {file_format}', title='Download')
    file_count[file_format] += 1
