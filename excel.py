import os
import re
import zipfile
import xml.etree.ElementTree as ET

class Excel:
    connectionRegex = r'[a-zA-Z0-9]{8}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{12} Model'

    def __init__(self, directory):
        self.directory = directory
        self.files = self.getFiles(self.directory)

    def __getitem__(self, key):
        return getattr(self, key)

    def edit(self, sql):
        for file in self.files:
            fileName = file.split('/')[-1]
            connectionString = sql.getConnectionFromFileName(fileName)
            if connectionString is None:
                print(f'No match for "{fileName}" in the SQL')
                continue

            # Step 1: Extract the XML content from the .xlsx file
            with zipfile.ZipFile(file, 'r') as zip_ref:
                zip_ref.extractall('unzipped_content')

            # Step 2: Identify and modify the connections XML
            # For example, if the connections are stored in the 'xl/connections.xml' file:
            connections_file = 'unzipped_content/xl/connections.xml'
            tree = ET.parse(connections_file)
            root = tree.getroot()

            # Step 3: Edit the XML
            for connection in root:
                attrib = str(connection.attrib)
                attrib = re.sub(Excel.connectionRegex, connectionString+' Model', attrib)
                connection.attrib = attrib
                for dbpr in connection.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dbPr'):
                    attrib = str(dbpr.attrib)
                    attrib = re.sub(Excel.connectionRegex, connectionString+' Model', attrib)
                    dbpr.attrib = attrib

            # Step 4: Repackage the XML files into an .xlsx archive
            with zipfile.ZipFile('modified_workbook.xlsx', 'w') as zip_ref:
                for folder_name, subfolders, filenames in os.walk('unzipped_content'):
                    for filename in filenames:
                        file_path = os.path.join(folder_name, filename)
                        archive_path = os.path.relpath(file_path, 'unzipped_content')
                        zip_ref.write(file_path, archive_path)

            # Step 5: Test and verify the modified workbook
            # Ensure that the workbook opens correctly in Excel and that the connections behave as expected

    def getFiles(self, directory, files=None):
        if files is None:
            files = []

        for name in os.listdir(directory):
            f = '/'.join([directory, name])
            
            if os.path.isdir(f):
                files = self.getFiles(f, files)
            elif os.path.isfile(f) and name.split('.')[-1] == 'xlsx':
                files.append(f)
            
        return files

    def test():
        directory = 'data'
        #workspaceId = '86577be9-9e83-4688-8a66-9c10fb4f130d'

        return Excel(directory)

    