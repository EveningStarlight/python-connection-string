import sys
from PyQt6 import QtCore
from PyQt6.QtWidgets import QApplication, QLabel, QWidget, QLineEdit, QFrame, QPushButton, QFileDialog
from PyQt6.QtWidgets import QVBoxLayout, QHBoxLayout, QStackedLayout

from excel import Excel

"""
Constructs the gui for the pipline healper.

The gui is made up of two main parts, the input window on top,
and an Instruction pannel below
"""
class GUI:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.title = "Connection String Updater"
        self.excel = Excel.test()

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
        self.setGeometry(100, 100, 280, 80)

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
        layout.addLayout(InputBox(self.gui, "File Name", "directory"))
        #layout.addLayout(InputBox(self.gui, "workspace ID", "workspaceID"))

        self.setLayout(layout)

    # Adds the buttons to the Control Frame
    def initControlButtons(self, controlLayout):
        layout = QHBoxLayout()

        directoryButton = QPushButton()
        directoryButton.setText("Select Folder")
        directoryButton.clicked.connect(lambda : self.selectDirectory())
        layout.addWidget(directoryButton)

        controlLayout.addLayout(layout)

    def selectDirectory(self):
        path = QFileDialog.getExistingDirectory(self, 'Select Folder', './data/')
        if path:
            self.gui.inputs['text']['directory'].setText(path)
            self.excel = Excel(path)
    
# An input box for the user to enter the key values for the activity
class InputBox(QHBoxLayout):
    def __init__(self, gui, label, key):
        super().__init__()
        self.gui = gui

        labelWidget = QLabel(label)
        labelWidget.setMinimumWidth(120)
        self.addWidget(labelWidget)

        inputBox = QLineEdit()
        inputBox.setText(self.gui.excel[key])
        self.gui.inputs['text'][key] = inputBox
        self.addWidget(inputBox)