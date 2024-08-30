
import pandas as pd
from datetime import datetime

from GlobalVariables import FileLoc

class StandardizeHelper:

    def __init__(self, line, filedate):
        self.field_mappings = pd.read_excel(FileLoc.FIELD_MAPPINGS.value, sheet_name=0).fillna("")
        self.column_widths = list(self.field_mappings.iloc[0])
        self.line = line
        self.filedate = filedate
        self.upload_timestamp = None

    def mapColumns(self, df):
        """ Maps dataframe columns to root columns as defined by
            the field mappings, where the top row is the desired end
            column and each value below it represents all column names
            which map to it
        :param df: Input dataframe
        :return: (dataframe) With columns from top row of field mappings
                    and all columns which can be mapped filled
        """
        # Establish mapped dataframe with same header as field mappings
        mapped_df = pd.DataFrame(columns=self.field_mappings.columns).fillna("")
        # For each of our report fields
        for field in self.field_mappings.columns:
            # Check if field has a direct match
            if field in df.columns:
                mapped_df[field] = df[field]
            else:
                # Check through all accepted alternate names
                alternate_names = self.field_mappings[field].values
                for alternate_name in alternate_names:
                    if alternate_name in df.columns:
                        mapped_df[field] = df[alternate_name]

        return mapped_df

    def preprocessColumns(self, mapped_df):
        """ Use file-specific knowledge to adjust
            or generate column values
        :param mapped_df: (dataframe) of mapped upload file
        :return: (dataframe) preprocessed columns
        """
        preprocessed_df = mapped_df

        # <= VISHAY =>
        # Convert Date column to YYYY-mm-dd
        if self.line == 'VISHAY':
            for i in preprocessed_df.index:
                try:
                    date = str(int(preprocessed_df.loc[i, 'Date']))
                    date = date[:4] + "-" + date[4:6] + "-" + date[6:]
                    preprocessed_df.loc[i, 'Date'] = date
                except ValueError or IndexError:
                    pass

        return preprocessed_df

    def generateColumns(self, preprocessed_df):
        """ Generate columns which currently are blank. This
            should be all columns highlighted "blue" in the
            field mappings file
        :param preprocessed_df: Mapped, preprocessed dataframe
        :return: (dataframe) with generated columns
        """
        standard_df = preprocessed_df

        # <= LINE AND FILEDATE =>
        standard_df.loc[:, 'Line'] = self.line
        standard_df.loc[:, 'File Date'] = self.filedate

        # <= UPLOAD TIMESTAMP =>
        # Get current datetime
        now = datetime.now()
        self.upload_timestamp = now.strftime("%Y-%m-%d %H.%M.%S")
        # Populate timestamp column
        standard_df.loc[:, 'Upload Timestamp'] = self.upload_timestamp

        return standard_df
