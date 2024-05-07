import os
import zipfile
import xml.etree.ElementTree as ET

testFile = "data/ISB_SSRS_Usage.xlsx"

def main():
    # Step 1: Extract the XML content from the .xlsx file
    with zipfile.ZipFile(testFile, 'r') as zip_ref:
        zip_ref.extractall('unzipped_content')

    # Step 2: Identify and modify the connections XML
    # For example, if the connections are stored in the 'xl/connections.xml' file:
    connections_file = 'unzipped_content/xl/connections.xml'
    tree = ET.parse(connections_file)
    root = tree.getroot()

    # Step 3: Edit the XML
    for connection in root:
        print(connection)
        for dbpr in connection.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dbPr'):
            print(dbpr)

def skip():
    # Step 4: Repackage the XML files into an .xlsx archive
    with zipfile.ZipFile('modified_workbook.xlsx', 'w') as zip_ref:
        for folder_name, subfolders, filenames in os.walk('unzipped_content'):
            for filename in filenames:
                file_path = os.path.join(folder_name, filename)
                archive_path = os.path.relpath(file_path, 'unzipped_content')
                zip_ref.write(file_path, archive_path)

    # Step 5: Test and verify the modified workbook
    # Ensure that the workbook opens correctly in Excel and that the connections behave as expected
    
def readFile(path):
    f = open(path, "rb")
    print(f.read())

def getFilesIn(directory="data"):
    files = []

    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        # checking if it is a file
        if os.path.isfile(f) and filename.split('.')[-1] == 'xlsx':
            files.append(f)

    return files
    
if __name__ == "__main__":
    main()