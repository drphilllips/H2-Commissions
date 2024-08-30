
import os
import shutil
import pandas as pd
from datetime import datetime
import win32com.client as win32

from GlobalVariables import FileLoc
from FormatHelper import FormatHelper

class ExcelHelper:

    def __init__(self):
        pass

    def saveError(self, filepath):
        """ Checks for obstacles with saving the output file
        :param: filepath: Filepath to check
        :return: (boolean) Whether there is an error with saving
        """
        try:
            open(filepath, 'r+')
        except FileNotFoundError:
            pass
        except PermissionError:
            return True
        return False

    def openFile(self, filepath):
        """ Opens Excel file
        :param filepath: Path to the file
        :return: (void) Open file
        """
        # Open file using OS commands (pywin32)
        excel = win32.Dispatch("Excel.Application")
        excel.WindowState = -4137  # xlMaximized
        excel.Visible = True
        absolute_path = os.path.abspath(filepath)
        workbook = excel.Workbooks.Open(absolute_path)
        workbook.Activate()
        excel.Windows(workbook.Name).Activate()

        filename = os.path.basename(filepath)
        print(f"> Launched {filename}")

    def backupFile(self, filepath):
        """ Create copy of an Excel file to the backup directory
        :param filepath: path to original file
        :return: String path to backup file
        """
        # Extract filename for print statements
        filename = os.path.basename(filepath)
        # Make sure the file exists
        if not os.path.exists(filepath):
            print(f"> Unable to backup {filename}"
                  f" because it does not exist.")
        else:

            # <= DEFINE BACKUP PATH =>
            # Get backup directory
            backup_dir = FileLoc.BACKUP.value
            # Split filename to name and file-type extension
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime('%y-%m-%d')
            backup_name = f"{name}_({timestamp}){ext}"
            backup_path = os.path.join(backup_dir, backup_name)

            # <= COPY FILE TO NEW BACKUP PATH =>
            shutil.copy(filepath, backup_path)
            print(f"> {filename} successfully backed up!")

    def createFile(self, filepath, dfs, sheets, widths):
        """ Creates an Excel file from dataframes, where each
            dataframe-name-col_width arr defines each sheet
            :param filepath: Path to desired output location
            :param dfs: Array of dataframes (one per sheet)
            :param sheets: Array of names for each sheet
            :param widths: Array of column width arrays for each sheet
            :return: Create file and return New filepath
            """
        filepath = os.path.abspath(filepath)
        filename = os.path.basename(filepath)
        print(f"..Creating {filename}..")

        # <= ENSURE ALL INPUT ARRAYS ARE OF THE SAME LENGTH =>
        if not (len(dfs) == len(sheets) and len(dfs) == len(sheets)):
            print(f"> Cannot export {filepath}. Sheet-definition arrays are different sizes.")
            filepath = ""
        else:

            # <= VERIFY THAT FILEPATH IS GOOD TO EXPORT =>
            if self.saveError(filepath):
                print(f"> Could not save {filename}, the file is currently open in Excel!"
                      f" Please close the file and try again.")
                filepath = ""
            else:

                # <= WRITE THE OUTPUT FILE =>
                writer = pd.ExcelWriter(filepath,
                                        engine="xlsxwriter",
                                        date_format="yyyy-mm-dd", datetime_format="yyyy-mm-dd")
                # Get us a format helper
                format_helper = FormatHelper(writer)
                # Iterate through arrays of sheet definition
                for df, sheet, width in zip(dfs, sheets, widths):
                    # Export dataframe to Excel
                    df.to_excel(writer, sheet_name=sheet, index=False)
                    # Format the Excel file
                    format_helper.formatSheet(df, sheet, width)
                # Save the file
                writer.close()
                print(f"> New file saved at: {filepath}")

        return filepath


