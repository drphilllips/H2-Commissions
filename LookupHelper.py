
import pandas as pd

from GlobalVariables import FileLoc

class LookupHelper:

    def __init__(self):
        self.standardize_helper = None
        self.excel_helper = None
        files_, columns, values = [pd.read_excel(FileLoc.LOOKUP_MATRIX.value, sheet_name=i).fillna("") for i in range(3)]
        self.files = {}
        for n in files_['Number']:
            self.files[n] = File(files_, n)
        self.standard_name_dict = {}
        for i, lookup_name in enumerate(columns['Lookup Name']):
            self.standard_name_dict[lookup_name] = columns.loc[i, 'Standard Name']
        self.paths = {}
        for i, value in enumerate(values['Value']):
            try:
                path = values.loc[i, 'Path'].split(sep="@")
                path = [int(s) for s in path]
            except AttributeError:  # Given a single number
                path = [values.loc[i, 'Path']]
            if value in self.paths.keys():
                self.paths[value].append(path)
            else:
                self.paths[value] = [path]

    def setStandardizeHelper(self, standardize_helper):
        self.standardize_helper = standardize_helper

    def setExcelHelper(self, excel_helper):
        self.excel_helper = excel_helper

    def performLookup(self, standard_df, value):
        """ For each row of the dataframe, perform lookups
        which determine the value
        :param standard_df: Standardized columns
        :param value: Name of column which we perform lookup for
        :return: Standardized, preprocessed, generated DF
                 with lookup value populated
        """
        lookup_df = standard_df

        # Set flag to know which lookup files have entries need fixing
        files_enf = []

        for i in lookup_df.index:
            lookup_flag = ""

            # <= ATTEMPT ALL LOOKUP PATHS =>
            num_paths = len(self.paths[value])
            for path_index in range(num_paths):
                path = self.paths[value][path_index]
                lookup_output = ""
                num_path_steps = len(path)
                # Perform each file lookup (step) along the path
                for step_index in range(num_path_steps):
                    file = self.files[path[step_index]]
                    key_col, val_col = file.key_val_pair
                    key_list, val_list = file.key_val_list
                    standard_key, standard_val = self.standard_name_dict[key_col], self.standard_name_dict[val_col]
                    # Use previous step's lookup output as key (if it's there)
                    key = lookup_output or str(lookup_df.loc[i, standard_key]).upper()

                    # <= SEARCH VALUE COLUMN =>
                    try:
                        val_index = val_list.index(key)
                        lookup_output = key
                        if file.updatable or step_index + 1 < num_path_steps:
                            lookup_df.loc[i, standard_val] = lookup_output
                    except ValueError:

                        # <= SEARCH KEY COLUMN =>
                        try:
                            key_index = key_list.index(key)
                            lookup_output = val_list[key_index]
                            if file.updatable or step_index + 1 < num_path_steps:
                                lookup_df.loc[i, standard_val] = lookup_output
                        except ValueError:

                            # <= ONTO NEXT PATH =>
                            # Determine lookup flag
                            lookup_flag = file.lookup_flag
                            if file.updatable:
                                lookup_df.loc[i, standard_val] = "ENF"
                                if key not in file.new_keys:
                                    file.new_keys.append(key)
                                    if file.number not in files_enf:
                                        files_enf.append(file.number)
                            # If we had a bad key, set previous file's val invalid
                            if step_index > 0:
                                previous_file = self.files[path[step_index - 1]]
                                if previous_file.updatable:
                                    if key not in previous_file.invalid_vals:
                                        previous_file.invalid_vals.append(key)
                                        if previous_file.number not in files_enf:
                                            files_enf.append(previous_file.number)
                            if lookup_output and path_index + 1 < num_paths:
                                lookup_output = ""
                            break
                # Break once we have an output
                if lookup_output and lookup_output != "ENF":
                    lookup_df.loc[i, value] = lookup_output
                    break
            # Save whatever flags were raised
            if lookup_flag:
                lookup_df.loc[i, 'Lookup Flag'] = lookup_flag

        # <= UPDATE LOOKUP FILES AUTOMATICALLY FOR IMPROVEMENT =>
        for file_number in files_enf:
            self.updateLookupFile(lookup_df, file_number)

        return lookup_df

    def updateLookupFile(self, lookup_df, number):
        """ Add new keys and replace invalid values
        :param lookup_df: Standardized dataframe
        :param number: File number
        :return: (void) update and open file
        """
        file = self.files[number]
        key_col, val_col = file.key_val_pair
        columns = file.id_columns + file.key_val_pair
        id_vals = list(lookup_df[file.id_columns].iloc[0])

        # <= CHANGE INVALID VALS =>
        for invalid_val in file.invalid_vals:
            file.df[val_col] = file.df[val_col].replace(invalid_val, 'ENF')

        # <= APPEND NEW KEYS =>
        append_df = pd.DataFrame(columns=columns)
        for new_key in file.new_keys:
            new_row = id_vals + [new_key, 'ENF']
            append_df.loc[len(append_df)] = new_row
        append_df.drop_duplicates(subset=key_col,
                                  keep='last',
                                  ignore_index=True)
        append_df = append_df.reset_index(drop=True)
        # Append to file
        file.df = pd.concat([file.df, append_df])
        file.df = file.df.drop_duplicates(subset=key_col,
                                          keep='last',
                                          ignore_index=True)
        file.df = file.df.reset_index(drop=True)

        # <= SORT FILE ROWS =>
        # Custom sort function: Place 'ENF' on top, everything else is sorted regularly
        def sortEnfOnTop(x): return (0, x) if x == 'ENF' else (1, x)
        # Sort by the custom key
        sort_by = ['Upload Timestamp', val_col] if 'Upload Timestamp' in file.id_columns else val_col
        ascending = [False, True] if 'Upload Timestamp' in file.id_columns else True
        file.df = file.df.sort_values(by=sort_by,
                                      ascending=ascending,
                                      ignore_index=True,
                                      key=lambda col: col.map(sortEnfOnTop)
                                      if col.name == val_col else col)
        file.df = file.df.reset_index(drop=True)

        # <= EXPORT UPDATED FILE =>
        # Get column widths from field mappings
        fields = self.standardize_helper.field_mappings[columns]
        column_widths = list(fields.iloc[0])
        output_filepath = self.excel_helper.createFile(file.path,
                                                       dfs=[file.df],
                                                       sheets=['Lookup'],
                                                       widths=[column_widths])
        self.excel_helper.openFile(output_filepath)

class File:

    def __init__(self, files, number):
        self.number = number
        self.name = files.loc[number, 'Name']
        self.path = FileLoc.LOOKUP.value + self.name
        self.df = pd.read_excel(self.path, sheet_name=0).fillna("")
        self.updatable = files.loc[number, 'Updatable']
        self.key_val_pair = files.loc[number, 'Key-Value Pair'].split(sep="@")
        self.key_val_list = [self.df[self.key_val_pair[0]].astype(str).str.upper().tolist(),
                             self.df[self.key_val_pair[1]].astype(str).str.upper().tolist()]
        self.lookup_flag = files.loc[number, 'Lookup Flag']
        try:
            self.id_columns = files.loc[number, 'ID Columns'].split(sep="@")
        except AttributeError:
            self.id_columns = None
        self.new_keys = []
        self.invalid_vals = []



