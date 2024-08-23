
import os
import sys
import pandas as pd

from PyQt5.uic import loadUi
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog

from GlobalVariables import FileLoc

VERSION = "Alpha v0.1"

class Stream(QtCore.QObject):
    """Redirects console output to text widget"""
    newText = QtCore.pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))

    def flush(self):
        """Pass the flush, so we don't get an attribute error"""
        pass

class MainWindow(QDialog):
    """Generates the main window for our program"""
    def __init__(self):
        super(MainWindow, self).__init__()

        # <= INITIALIZE WINDOW STATE =>
        # Initialize state variables
        self.input_filepath = ""
        self.input_df = None
        self.updated_lookup_files = {}
        # Create custom output stream
        self.stream = Stream()
        self.stream.newText.connect(self.writeToConsole)
        sys.stdout = self.stream

        # <= CONNECT USER INTERFACE =>
        # Load external UI design w/ QtDesigner
        loadUi("H2 Commissions.ui", self)
        # Group elements for future ease of access
        self.all_elements = [self.btn_select_file, self.btn_deselect_file,
                             self.btn_assign_fse, self.btn_add_to_master, self.btn_clear_console]
        self.file_selected_elements = [self.btn_deselect_file, self.btn_assign_fse, self.btn_add_to_master]
        self.file_deselected_elements = [self.btn_select_file]
        self.file_either_elements = [self.btn_clear_console]
        self.lockButtons()
        self.unlockButtons()
        # Connect buttons to functions
        self.btn_clear_console.clicked.connect(self.clearConsole)
        self.btn_select_file.clicked.connect(self.selectFile)
        self.btn_deselect_file.clicked.connect(self.deselectFile)
        self.btn_assign_fse.clicked.connect(self.assignFSE)
        self.btn_add_to_master.clicked.connect(self.addToMaster)

        # Show welcome message
        self.clearConsole()

    def resetStateVariables(self):
        """Reset app state variables"""
        self.input_filepath = ""
        self.input_df = None
        self.updated_lookup_files = {}

    def writeToConsole(self, text):
        """Write console output to text widget"""
        cursor = self.txt_console.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.txt_console.setTextCursor(cursor)
        self.txt_console.ensureCursorVisible()
        QtWidgets.QApplication.processEvents()  # Process pending events

    # =======================
    #  GUI BUTTON OPERATIONS
    # -----------------------

    def clearConsole(self):
        """Clear console print statements"""
        self.txt_console.clear()
        print("> Welcome to the H&2 Commissions Program!")

    def selectFile(self):
        """Select input commissions file"""
        # Lock buttons until process concludes
        self.lockButtons()

        # <= DESELECT ANY SELECTED FILE =>
        if self.input_filepath:
            self.input_filepath = ""
            self.lbl_selected_file.setText("<No File Selected>")
            self.input_df = pd.DataFrame()
            print("..Selecting new file, old selection cleared..")

        # <= SHARE STATUS WITH USER =>
        print("..Loading file..")

        # <= PROMPT USER TO SELECT INPUT FILE =>
        self.input_filepath, _ = QFileDialog.getOpenFileName(self, directory=FileLoc.INPUT.value,
                                                             filter="Excel files (*.xls *.xlsx *.xlsm)")

        # <= MAKE SURE USER DID NOT CANCEL =>
        if not self.input_filepath:
            self.deselectFile()
            print("> Select file operation cancelled.")
        else:

            # <= LOAD FILE TO APP =>
            # Convert selected file to dataframe
            self.input_df = pd.read_excel(self.input_filepath, sheet_name=0).fillna("")

            # <= UPDATE USER WITH STATUS =>
            # Print out the selected filename
            filename = os.path.basename(self.input_filepath)
            print("> File load complete: " + filename)
            # Shorten filename if too long for selected files label
            if len(filename) > 43:
                filename = filename[:43] + "..."
            # Update current file label
            self.lbl_selected_file.setText("> " + filename)
            # Enable buttons and drop-downs now that file is selected (or even if not selected)
            self.unlockButtons()

    def deselectFile(self):
        """Deselect input file"""

        # <= DESELECT ANY SELECTED FILE =>
        if self.input_filepath:
            self.resetStateVariables()
            self.lbl_selected_file.setText("<No File Selected>")
            print("> File selection cleared.")
        # Disable buttons now that file is deselected
        self.lockButtons()
        [e.setEnabled(True) for e in self.file_deselected_elements]
        [e.setEnabled(True) for e in self.file_either_elements]

    # =======================
    #  GUI UTILITY FUNCTIONS
    # -----------------------

    def lockButtons(self):
        """Disable user interaction"""
        for element in self.all_elements:
            element.setEnabled(False)

    def unlockButtons(self):
        """Enable user interaction"""
        # Enable elements based on whether file is selected
        if not self.input_filepath:
            # Enable elements which are required regardless of fsr type
            [e.setEnabled(True) for e in self.file_deselected_elements]
        else:
            # Enable elements which are required regardless of time period type
            [e.setEnabled(True) for e in self.file_selected_elements]
        # Enable elements which are required regardless of whether file is selected or not
        [e.setEnabled(True) for e in self.file_either_elements]

    # ========================
    #  Main Program Functions
    # ------------------------

    def assignFSE(self):
        """Assign FSE to each line of input commissions file"""
        print("..Assigning FSE..")

    def addToMaster(self):
        """Add fse-assigned file to commissions master file"""
        print("..Adding to Commissions Master..")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Create widget container for QtDesigner UI
    widget = QtWidgets.QStackedWidget()
    widget.setWindowTitle("H&2 Commissions (" + VERSION + ")")
    main_window = MainWindow()
    widget.addWidget(main_window)
    widget.setFixedWidth(900)
    widget.setFixedHeight(600)
    widget.show()

try:
    sys.exit(app.exec_())
except Exception:
    print("..Exiting..")

