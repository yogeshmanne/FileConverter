import pyodbc
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

driver = '{ODBC Driver 17 for SQL Server}'
server = 'VENKEY\\SQLEXPRESS'  # Note the double backslash to escape the backslash
database = 'file converter'
username = 'sa'
password = '223355'

connection_string = (
    f'DRIVER={driver};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password}'
)

def prettify_xml(elem):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")

def sanitize_column_name(name):
    """Sanitize column names to be valid XML tags."""
    return ''.join(c if c.isalnum() else '_' for c in name)

def dataframe_to_xml(df):
    """Convert a pandas DataFrame to an XML string."""
    root = ET.Element("root")
    for i, row in df.iterrows():
        item = ET.SubElement(root, "item")
        for field in df.columns:
            field_sanitized = sanitize_column_name(field)
            ET.SubElement(item, field_sanitized).text = str(row[field])
    return prettify_xml(root)

try:
    conn = pyodbc.connect(connection_string)

    print("Connection successful!")

    def read_file(file_path):
        file_name = file_path.split('\\')[-1]
        root = ET.Element("file_content")
        ET.SubElement(root, "file_name").text = file_name

        if file_path.endswith('.txt'):
            with open(file_path, 'r') as file:
                content = file.read()
                ET.SubElement(root, "content").text = content
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
            content = dataframe_to_xml(df)
            root.append(ET.fromstring(content))
        elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            df = pd.read_excel(file_path)
            content = dataframe_to_xml(df)
            root.append(ET.fromstring(content))
        elif file_path.endswith('.xml'):
            tree = ET.parse(file_path)
            root = tree.getroot()
        else:
            print("Unsupported file type")
            return

        xml_content = prettify_xml(root)
        print("XML file content:\n", xml_content)

        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO FileContents (FileName, Content) VALUES (?, ?)",
            (file_name, xml_content)
        )
        conn.commit()
        cursor.close()

        root = ET.fromstring(xml_content)
        print("Parsed XML content:")
        for elem in root.iter():
            print(f"{elem.tag}: {elem.text}")

    file_path = r"C:\Users\T.Venkateshwar Reddy\Downloads\access-code.csv"  
    read_file(file_path)

    conn.close()

except Exception as e:
    print("Error: ", e)