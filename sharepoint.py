import os
import tempfile
import shutil
from pathlib import Path
import json

from office365.sharepoint.client_context import ClientContext, ClientCredential
from office365.sharepoint.files.system_object_type import FileSystemObjectType

class Sharepoint:
    def __init__(self):
        with open('secrets.json', 'r') as file:
            secrets = json.load(file)

        self.url = 'https://054gc.sharepoint.com/sites/EnterpriseBusinessIntelligenceIntelligencedaffairesdelentreprise/'
        self.directory = 'Shared Documents/FRS Migration/Excel and PortalDoc/Program/'
        self.tenant = secrets.get('tenant')
        self.client = secrets.get('client')
        self.secret = secrets.get('secret')

        self.tempDirectory = None

        client_credentials = ClientCredential(self.client, self.secret)
        self.ctx = ClientContext(self.url).with_credentials(client_credentials)
        
        self.files = self.getAllFiles()

    def uploadFiles(self, updatedFiles):

        for file_path in updatedFiles:
            with open(file_path, 'rb') as content_file:
                file_content = content_file.read()
            
            file_name = os.path.basename(file_path)
            file_url = self.getServerRelativeUrl(file_name)
            folder_url = Path(file_url).parent

            target_folder = self.ctx.web.get_folder_by_server_relative_path(str(folder_url))
            target_folder.upload_file(file_name, file_content)
            self.ctx.execute_query()

    def downloadFiles(self):
        print(f'Fetching {len(self.files)} files from sharepoint')

        self.tempDirectory = tempfile.mkdtemp()
        directoryPath = Path(os.path.join(self.tempDirectory, self.directory))
        self.directoryPath = directoryPath
        directoryPath.mkdir(parents=True, exist_ok=True)

        for file in self.files:
            file_url = file.file.serverRelativeUrl

            path = Path(file_url.split(self.directory)[-1])
            path.parent.mkdir(parents=True, exist_ok=True)
            
            filePath = Path(os.path.join(directoryPath, path))

            with open(filePath, "wb") as local_file:
                file = (
                    self.ctx.web.get_file_by_server_relative_path(file_url)
                    .download(local_file)
                    .execute_query()
                )
        
        print(f'Fetched {len(self.files)} files from Sharepoint')
        return directoryPath

    def getAllFiles(self):
        doc_lib = self.ctx.web.default_document_library()
        items = (
            doc_lib.items.select(["FileSystemObjectType"])
            .expand(["File", "Folder"])
            .get_all()
            .execute_query()
        )
        
        files = []
        for item in items:
            if item.file_system_object_type != FileSystemObjectType.Folder:
                if self.directory in item.file.serverRelativeUrl:
                    files.append(item)

        return files
    
    def getServerRelativeUrl(self, fileName):

        for file in self.files:
            if fileName == os.path.basename(file.file.serverRelativeUrl):
                return file.file.serverRelativeUrl
        return None
    
    def cleanUp(self):
        if self.tempDirectory is not None:
            shutil.rmtree(self.tempDirectory)


def main():
    sharepoint = Sharepoint()
    sharepoint.downloadFiles()

    print(Path(sharepoint.getServerRelativeUrl('FRS Report_POC_agg_withPeriod.xlsx')).parent)

if __name__ == "__main__":
    main()
