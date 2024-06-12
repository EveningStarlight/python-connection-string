import os
import re
import zipfile
import shutil
import xml.etree.ElementTree as ET

class Excel:
    connectionRegex = r'[a-zA-Z0-9]{8}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{4}-[a-zA-Z0-9]{12} Model'
    extractPath = 'unzipped_content'
    connectionPath = 'unzipped_content/xl/connections.xml'

    def __init__(self, directory):
        self.directory = directory
        self.files = self.getFiles(self.directory)

    def __getitem__(self, key):
        return getattr(self, key)

    def edit(self, sql):
        filesEdited=0
        for file in self.files:
            fileName = file.split('/')[-1]
            #connectionString = sql.getConnectionFromFileName(fileName)
            connectionString = 'f406979f-c560-4a00-b341-b89e07787eca'
            if connectionString is None:
                print(f'No match for "{fileName}" in the SQL')
                continue

            # Step 1: Extract the XML content from the .xlsx file
            with zipfile.ZipFile(file, 'r') as zip_ref:
                zip_ref.extractall(Excel.extractPath)

            # Read the contents of the file
            with open(Excel.connectionPath, 'r') as f:
                file_contents = f.read()

            # Perform the regex replacement
            modified_contents = re.sub(Excel.connectionRegex, connectionString+' Model', file_contents)

            
            if file_contents != modified_contents:
                print('edited string')
                filesEdited += 1
                # Save the modified contents to a new file
                with open(Excel.connectionPath, 'w') as f:
                    f.write(modified_contents)

                # Step 4: Repackage the XML files into an .xlsx archive
                with zipfile.ZipFile(file, 'w') as zip_ref:
                    for folder_name, subfolders, filenames in os.walk(Excel.extractPath):
                        for filename in filenames:
                            file_path = os.path.join(folder_name, filename)
                            archive_path = os.path.relpath(file_path, Excel.extractPath)
                            zip_ref.write(file_path, archive_path)

            # Step 5: Clean up
            shutil.rmtree(Excel.extractPath)
        return {'filesEdited': filesEdited, 'totalFiles': len(self.files)}
                

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

    