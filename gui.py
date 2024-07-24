import sys
from PyQt6.QtWidgets import QApplication, QLabel, QWidget, QLineEdit, QFrame, QPushButton, QFileDialog, QMessageBox
from PyQt6.QtWidgets import QVBoxLayout, QHBoxLayout

from excel import Excel
from sql import SQL
from sharepoint import Sharepoint

"""
Constructs the gui
"""
class GUI:
    def __init__(self):
        self.sharepoint = Sharepoint()
        path = self.sharepoint.downloadFiles()

        self.app = QApplication(sys.argv)
        self.title = "Connection String Updater"
        self.excel = Excel(path)
        self.sql = None

        self.window = MainWindow(self)

        self.inputs = {}
        self.inputs['text'] = {}

        self.initializeLayout()

        self.window.show()
        sys.exit(self.app.exec())

    # Builds the gui using specific builders
    def initializeLayout(self):
        layout = QVBoxLayout()

        layout.addWidget(ControlFrame(self))

        self.window.setLayout(layout)


# The main window of the project
class MainWindow(QWidget):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui
        
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.gui.title)
        self.setGeometry(800, 400, 400, 120)

# The Control Frame
class ControlFrame(QFrame):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui

        self.initUI()
        self.initContents()
        
    def initUI(self):
        self.setFrameStyle(QFrame.Shape.Panel | QFrame.Shadow.Sunken)
        self.setLineWidth(2)

    def initContents(self):
        layout = QVBoxLayout()

        self.initControlButtons(layout)
        layout.addLayout(InputBox(self.gui, "Directory", "directory", self.gui.excel["directory"]))
        layout.addLayout(InputBox(self.gui, "Files", "fileCount", len(self.gui.excel.files)))
        self.initRunButton(layout)

        self.setLayout(layout)

    # Adds the buttons to the Control Frame
    def initControlButtons(self, controlLayout):
        layout = QHBoxLayout()

        controlLayout.addLayout(layout)

    # Adds the run program button
    def initRunButton(self, controlLayout):
        layout = QHBoxLayout()

        directoryButton = QPushButton()
        directoryButton.setText("Change Connection Strings")
        directoryButton.clicked.connect(lambda : self.changeConnectionStrings())
        layout.addWidget(directoryButton)

        controlLayout.addLayout(layout)
    
    def changeConnectionStrings(self):
        if self.gui.sql is None:
            self.gui.sql = SQL()
        value = self.gui.excel.edit(self.gui.sql)

        messageText = f"Total files found: {len(value['files'])}\nTotal matches found in SQL table: {len(value['files'])-value['filesNotFound']}\nTotal files modified: {len(value['editedFiles'])}"

        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setText(messageText)
        msg_box.setWindowTitle("Success")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        
        # Show the message box and wait for user interaction
        retval = msg_box.exec()
        
        # Exit the application when OK is clicked
        if retval == QMessageBox.StandardButton.Ok:
            print(f"Uploading {len(value['editedFiles'])} files")
            try:
                self.gui.sharepoint.uploadFiles(value['editedFiles'])
            finally:
                self.gui.sharepoint.cleanUp()
            self.gui.app.quit()

    
# An input box for the user to enter the key values for the activity
class InputBox(QHBoxLayout):
    def __init__(self, gui, label, key, value):
        super().__init__()
        self.gui = gui

        labelWidget = QLabel(label)
        labelWidget.setMinimumWidth(60)
        self.addWidget(labelWidget)

        inputBox = QLineEdit()
        inputBox.setText(str(value))
        inputBox.setReadOnly(True)
        self.gui.inputs['text'][key] = inputBox
        self.addWidget(inputBox)
