

import pandas as pd

from GlobalVariables import FileLoc

class FormatHelper:

    def __init__(self, writer):
        self.writer = writer
        columns, flags = [pd.read_excel(FileLoc.FORMAT_MATRIX.value, sheet_name=i).fillna("") for i in range(2)]

        # <= EXTRACT FORMATS FOR COLUMNS =>
        self.column_formats = {}
        self.format_columns = {}
        for i in columns.index:
            name = columns.loc[i, 'Name']
            # Create dict to store all of this font's attributes
            font_dict = {}
            for attr in ['font', 'font_size', 'num_format', 'align']:
                font_dict[attr] = columns.loc[i, attr]
            # Add font to column formats dictionary
            self.column_formats[name] = writer.book.add_format(font_dict)
            # Pull columns which have this format
            self.format_columns[name] = columns.loc[i, 'Columns'].split(sep="@")

        # <= EXTRACT FLAG FORMATS =>
        self.flag_formats = {}
        for i in flags.index:
            font_dict = {}
            for attr in ['font', 'font_size', 'bg_color']:
                font_dict[attr] = flags.loc[i, attr]
            self.flag_formats[flags.loc[i, 'Value']] = writer.book.add_format(font_dict)

    def formatSheet(self, df, sheet, width):
        """ Formats our output file to make it look nice
        :param df: Working dataframe for output
        :param sheet: Name of the sheet we are working on
        :param width: Widths of columns
        :return: (void) format Excel file
        """
        # <= ENSURE WE HAVE DATA TO FORMAT =>
        if not df.empty and len(width) > 0:
            print(f'..Formatting "{sheet}" sheet..')

            # <= STORE WORKING SHEET =>
            sheet = self.writer.sheets[sheet]

            # <= FORMAT THE HEADER =>
            header_left_cols = []
            # Write the DataFrame header to the worksheet with the defined format
            header_row = 0
            for col_num, value in enumerate(df.columns.values):
                if value in header_left_cols:
                    sheet.write(header_row, col_num, value, self.column_formats['left-aligned'])
                else:
                    sheet.write(header_row, col_num, value, self.column_formats['center-aligned'])
            # Freeze header, so it remains stationary when scrolling up or down
            sheet.freeze_panes(1, 0)

            # <= SET AUTO-FILTER =>
            sheet.autofilter(0, 0, df.shape[0], df.shape[1] - 1)

            # <= FORMAT THE BODY =>
            # Ignore number stored as text error
            sheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
            # Determine which data columns need which format
            # Style each column individually
            for column in df.columns:
                fmt = self.column_formats['default']
                # Setting style based on which format assigns to this column
                for name in self.column_formats:
                    if column in self.format_columns[name]:
                        fmt = self.column_formats[name]
                # Set column width and formatting
                col_idx = df.columns.get_loc(column)
                col_width = width[col_idx]
                sheet.set_column(col_idx, col_idx, col_width, fmt)
            # Set the row height for all rows
            row_height = 10.8
            for row_num in range(df.shape[0]):
                sheet.set_row(row_num, row_height)

            # <= FORMAT LOOKUP FLAGS =>
            if 'Lookup Flag' not in df.columns:
                # print('> No "Lookup Flag" column, unable to format with flags')
                pass
            else:
                # We are highlighting the end customer column
                customer_col_idx = list(df).index('Reported Customer')
                for i in df.index:
                    # Get the customer for us to rewrite in
                    customer = df.loc[i, 'Reported Customer']
                    # Read flag to determine color of highlight
                    lookup_flag = df.loc[i, 'Lookup Flag']
                    # Make sure our lookup flag is defined
                    try:
                        # Use predefined formatting based on lookup flag type
                        sheet.write(i + 1, customer_col_idx, customer, self.flag_formats[lookup_flag])
                    except KeyError:
                        pass
