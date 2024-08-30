
import os
import sys
import pandas as pd

from PyQt5.uic import loadUi
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog

from ExcelHelper import ExcelHelper
from GlobalVariables import FileLoc
from LookupHelper import LookupHelper
from StandardizeHelper import StandardizeHelper

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
        self.input_filename = ""
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
        self.input_filename = ""
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
            self.input_filename = os.path.basename(self.input_filepath)
            print("> File load complete: " + self.input_filename)
            # Shorten filename if too long for selected files label
            if len(self.input_filename) > 43:
                filename = self.input_filename[:43] + "..."
            # Update current file label
            self.lbl_selected_file.setText("> " + self.input_filename)
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

        self.lockButtons()

        # <= MAKE SURE WE HAVE ALL LOOKUP FILES READY =>
        excel_helper = ExcelHelper()
        lookup_helper = LookupHelper()
        lookup_files_ready = True
        for file in lookup_helper.files.values():
            if not os.path.exists(file.path):
                lookup_files_ready = False
                print(f"> Cannot assign FSE. Lookup file {file.name} cannot be found."
                      f" Please make sure file is in the Lookup directory.")
            if file.updatable:
                if excel_helper.saveError(file.path):
                    lookup_files_ready = False
                    print(f"> Cannot assign FSE. Updatable lookup file {file.name} is open."
                          f" Please make sure file is not open in Excel.")
        if lookup_files_ready:

            # <= PULL LINE AND DATE FROM FILENAME =>
            try:
                filename, ext = os.path.splitext(self.input_filename)
                line, filedate = filename.split(sep="@")
                proper_filename = True
            except ValueError:
                filename = None
                line, filedate = None, None
                proper_filename = False
            # Make sure we have a proper filename
            if not proper_filename:
                print(f"> {self.input_filename} is an invalid input file name."
                      f' Please use "<LINE>@<YYYY-MM-DD>.xlsx"')
            else:

                # <= BACKUP ALL UPDATABLE LOOKUP FILES =>
                for file in lookup_helper.files.values():
                    if file.updatable:
                        excel_helper.backupFile(file.path)

                # <= STANDARDIZE COLUMNS =>
                print("..Standardizing Columns..")
                standardize_helper = StandardizeHelper(line, filedate)
                standard_df = standardize_helper.mapColumns(self.input_df)
                standard_df = standardize_helper.preprocessColumns(standard_df)
                standard_df = standardize_helper.generateColumns(standard_df)

                # <= PERFORM LOOKUP ON STANDARD FILE =>
                print("..Assigning FSE..")
                lookup_helper.setStandardizeHelper(standardize_helper)
                lookup_helper.setExcelHelper(excel_helper)
                fse_df = lookup_helper.performLookup(standard_df, 'FSE Code')

                # <= EXPORT FILE TO EXCEL =>
                # Create output filepath
                output_filepath = f"{FileLoc.OUTPUT.value}{filename}_(FSE)_{{" +\
                                  standardize_helper.upload_timestamp + "}.xlsx"
                excel_helper.createFile(output_filepath,
                                        dfs=[fse_df],
                                        sheets=['Data'],
                                        widths=[standardize_helper.column_widths])
                excel_helper.openFile(output_filepath)

        self.unlockButtons()

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

