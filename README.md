# Python Connection String

This is a tool to update excel connection strings by matching them with an SQL table. 
This allows for fast editing of multiple sheets

## installation

1. Download the repositor to your machine

2. Install the dependancies using pip install while in the root folder

```
pip install
```

Configure the secrets.json file. 
The tenant ID, client ID, and client secret are required to connect to the sharepoint.

## Running the program

From the root of the project, run 
```
python main.py
```

This will open a small application window. This window dispalys the directory where files were downloaded, and the total files downloaded. 

Press the `Change Connection Strings` button to connect to the SQL table, and update all the connection strings of downloaded files. This will bring up the Azure authentication dialouge box. 

Then a dialouge will be shown that tells you how many files were checked, how many matches were found in the SQL table, and how many connection strings were modified. Selecting `OK` will then begin the upload of edited files back to the sharepoint. 
