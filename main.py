
import os
import sys
import pandas as pd

from PyQt5.uic import loadUi
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QDialog, QFileDialog, QMessageBox

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
        value_lookups = lookup_helper.files.values()
        general_lookups = [FileLoc.FIELD_MAPPINGS.value, FileLoc.LOOKUP_MATRIX.value, FileLoc.FORMAT_MATRIX.value]
        lookup_files_ready = True
        for filepath in general_lookups + [file.path for file in value_lookups]:
            filename = os.path.basename(filepath)
            if not os.path.exists(filepath):
                lookup_files_ready = False
                print(f"> Cannot assign FSE. Lookup file {filename} cannot be found."
                      f" Please make sure file is in the Lookup directory.")
        for file in value_lookups:
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
                # Sort file
                fse_df = fse_df.sort_values(by='Reported Customer',
                                            ascending=True,
                                            ignore_index=True)
                fse_df = fse_df.reset_index(drop=True)
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

        self.lockButtons()

        # <= MAKE SURE WE HAVE ALL FILES READY =>
        # Start with lookup files
        general_lookups = [FileLoc.FIELD_MAPPINGS.value, FileLoc.FORMAT_MATRIX.value]
        lookup_files_ready = True
        for filepath in general_lookups:
            filename = os.path.basename(filepath)
            if not os.path.exists(filepath):
                lookup_files_ready = False
                print(f"> Cannot add to master. Lookup file {filename} cannot be found."
                      f" Please make sure file is in the Lookup directory.")
        # Make sure we can edit and open the master file
        excel_helper = ExcelHelper()
        master_file_ready = True
        if excel_helper.saveError(FileLoc.MASTER.value):
            master_file_ready = False
            print(f"> Cannot add to master. Master file is open."
                  f" Please make sure file is not open in Excel.")
        if lookup_files_ready and master_file_ready:

            # <= CONFIRM USER WANTS TO ADD =>
            reply = QMessageBox.question(self, "Confirm Add To Master",
                                         f"Are you sure you would like to add"
                                         f" {self.input_filename} to the"
                                         f" commissions master file?",
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.No)
            if reply == QMessageBox.No:
                print(f"> Add to master not confirmed."
                      f" Commissions master not updated.")
            else:

                # <= BACKUP MASTER FILE =>
                excel_helper.backupFile(FileLoc.MASTER.value)

                # <= LOAD MASTER FILE =>
                master_df = pd.read_excel(FileLoc.MASTER.value, sheet_name=0).fillna("")

                # <= MAKE SURE ALL INPUT AND MASTER COLUMNS ARE STANDARD =>
                field_mappings = pd.read_excel(FileLoc.FIELD_MAPPINGS.value, sheet_name=0).fillna("")
                input_columns = set(self.input_df.columns)
                master_columns = set(master_df.columns)
                field_mappings_columns = set(field_mappings.columns)
                if input_columns != master_columns:
                    missing = list(master_columns - input_columns)
                    extra = list(input_columns - master_columns)
                    print(f"> Add to master cancelled."
                          f" Input file columns are not the"
                          f" same as master file columns.\n"
                          f" Missing columns: {missing}\n"
                          f" Extra columns: {extra}")
                elif master_columns != field_mappings_columns:
                    missing = list(field_mappings_columns - master_columns)
                    extra = list(master_columns - field_mappings_columns)
                    print(f"> Add to master cancelled."
                          f" Master file columns are not the"
                          f" same as field mappings columns.\n"
                          f" Missing columns: {missing}\n"
                          f" Extra columns: {extra}")
                else:

                    # <= CLEAR OUT ANY IDENTICAL LINE/FILEDATE DATA =>
                    unique_id_cols = ['Line', 'File Date']
                    file_unique_id = list(self.input_df[unique_id_cols].iloc[0])
                    # Create a condition that identifies rows to remove
                    condition = pd.Series([True] * len(master_df), dtype=bool)
                    for col, val in zip(unique_id_cols, file_unique_id):
                        condition = condition & (master_df[col] == val)
                    # Remove rows where condition is True
                    master_df = master_df[~condition]
                    print(f"> Removed previous {'@'.join(file_unique_id)}"
                          f" data from master file.")

                    # <= APPEND TO MASTER =>
                    master_df = pd.concat([self.input_df, master_df])
                    master_df = master_df.sort_values(by=['Upload Timestamp', 'Reported Customer'],
                                                      ascending=[False, True],
                                                      ignore_index=True)
                    master_df = master_df.reset_index(drop=True)

                    # <= EXPORT MASTER FILE =>
                    column_widths = list(field_mappings.iloc[0])
                    output_filepath = excel_helper.createFile(FileLoc.MASTER.value,
                                                              dfs=[master_df],
                                                              sheets=["Data"],
                                                              widths=[column_widths])
                    excel_helper.openFile(output_filepath)

        self.unlockButtons()


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

