import os
import re
import zipfile
import shutil

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
        editedFiles = []
        filesNotFound=0
        for file in self.files:
            fileName = os.path.basename(file)
            connectionString = sql.getConnectionFromFileName(fileName)
            if connectionString is None:
                print(f'No match for "{fileName}" in the SQL')
                filesNotFound += 1
                continue

            with zipfile.ZipFile(file, 'r') as zip_ref:
                zip_ref.extractall(Excel.extractPath)

            with open(Excel.connectionPath, 'r') as f:
                file_contents = f.read()
                modified_contents = re.sub(Excel.connectionRegex, connectionString+' Model', file_contents)

            if file_contents != modified_contents:
                editedFiles.append(file)
                with open(Excel.connectionPath, 'w') as f:
                    f.write(modified_contents)

                with zipfile.ZipFile(file, 'w') as zip_ref:
                    for folder_name, subfolders, filenames in os.walk(Excel.extractPath):
                        for filename in filenames:
                            file_path = os.path.join(folder_name, filename)
                            archive_path = os.path.relpath(file_path, Excel.extractPath)
                            zip_ref.write(file_path, archive_path)

            # Clean up
            shutil.rmtree(Excel.extractPath)
        return {'editedFiles': editedFiles, 'files': self.files, 'filesNotFound': filesNotFound}
                

    def getFiles(self, directory, files=None):
        if files is None:
            files = []

        for name in os.listdir(directory):
            f = os.path.join(directory, name)
            
            if os.path.isdir(f):
                files = self.getFiles(f, files)
            elif os.path.isfile(f) and name.split('.')[-1] == 'xlsx':
                files.append(f)
            
        return files
