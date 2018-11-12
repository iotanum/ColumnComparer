import os
import string
import time
import pandas as pd
import numpy as np
from pathlib import Path


class DirectoryTree:
    def __init__(self):
        self.main_excel_needed_column = pd.Series(dtype=np.dtype(object))
        self.main_excel_df = pd.Series(dtype=np.dtype(object))
        self.not_found_df = pd.Series(dtype=np.dtype(object))
        self.duplicates = set()
        self.not_found_list = []
        self.found_rows_list = []
        self.secondary_column_list = []
        self.result_file_name = ""
        self.input_column_name = ""
        self.found_status = False
        self.row_status = False
#        self.full_secondary = None  # if doing specific tasks

    def file_tree(self, directory):
        paths = list()
        for path, dirs, files in os.walk(directory):
            paths.append(path)
        return paths

    def excel_file_names(self, directories):
        excel_list = list()
        for directory in directories:
            for file in os.listdir(directory):
                if file.endswith(".xls") or file.endswith(".xlsx") or file.endswith(".xlt"):
                    excel_list.append(Path(directory + "/" + file))
        return excel_list

    def get_excel_list(self, parent_dir):
        file_list = self.excel_file_names(self.file_tree(parent_dir))
        if file_list:
            return file_list

    def find_needed_column(self, file_path, column_name, main=False):
        xl = pd.read_excel(file_path, header=None)
        file_name = os.path.basename(file_path)
        self.input_column_name = " ".join(column_name)
        for column_num in xl:
            for column_row in xl[column_num]:
                if all(x.lower() in str(column_row).lower() for x in column_name):
                    df = pd.DataFrame(xl[column_num].dropna())  # specific task - delete dropna()
                    df.columns = [self.input_column_name]
                    if not main:
                        series = df.assign(file=pd.Series([file_name for _ in
                                                           range(len(df[self.input_column_name]) + 50)]))
#                        maybe I'll need it later for specific tasks
#                        self.full_secondary = xl
                        self.secondary_column_list.append(series)
                        return f"\n     Stulpelis {self.column_alphabetic_value(column_num)} - '{column_row}'.\n"
                    else:
                        self.main_excel_df = xl
                        self.main_excel_needed_column = xl[column_num]
                        return f"\n     Stulpelis {self.column_alphabetic_value(column_num)} - '{column_row}'.\n"

    def column_alphabetic_value(self, column_num):
        return string.ascii_uppercase[column_num]

    def export_excel(self, directory, df, found=True):
        self.result_file_name = f'{"rezultatai" if found else "nerasti rezultatai"} ' + \
                                                                str(time.strftime("%Y%m%d %H%M%S")) + ".xlsx"
        df.to_excel(directory + "/" + self.result_file_name)

    def make_excel(self, directory):
        df = pd.concat(self.secondary_column_list, ignore_index=True)
        self.compare_main_with_secondaries(directory, df)

    def make_found_excel(self, directory):
        # for specific tasks remove two first lines and last line
        df2 = pd.concat(self.not_found_list, ignore_index=True)
        print(len(df2), 'not found')
        df = pd.concat(self.found_rows_list, ignore_index=True)
        print(len(df), 'found')
        self.export_excel(directory, df)
        self.export_excel(directory, df2, found=False)

    def compare_main_with_secondaries(self, directory, sec_df):
        print(len(self.main_excel_needed_column), 'full file')
        for main_idx, column_row in enumerate(self.main_excel_needed_column):
            column_row = str(column_row)
            self.row_status = False
            for sec_idx, secondary_column_row in enumerate(sec_df[self.input_column_name]):
                secondary_column_row = str(secondary_column_row)
                if self.row_check(column_row, secondary_column_row) and self.input_column_name_check(column_row)\
                        and str(column_row) not in self.duplicates:
                    which_file_row = sec_df.iloc[[sec_idx], [1]]
                    which_file_row = which_file_row['file'][sec_idx]
                    new_df = self.main_excel_df.iloc[[main_idx]].assign(Rasta_dokumente=which_file_row)
                    self.duplicates.add(str(column_row))
                    self.found_rows_list.append(new_df)
                    self.row_status = True
            if not self.row_status and self.input_column_name_check(column_row):
                not_found_row = self.main_excel_df.iloc[[main_idx]]
                self.not_found_list.append(not_found_row)
        if self.found_rows_list:
            self.found_status = True
            self.make_found_excel(directory)

# For future, specific task
#    def compare_main_with_secondaries(self, directory, sec_df):
#        try:
#            for main_idx, column_row in enumerate(self.main_excel_needed_column):
#                column_row = str(column_row)
#                for sec_idx, secondary_column_row in enumerate(sec_df[self.input_column_name]):
#                    secondary_column_row = str(secondary_column_row)
#                    if self.row_check(column_row, secondary_column_row) and self.input_column_name_check(column_row)\
#                            and str(secondary_column_row) != 'nan':
#                        # number according to what is needed
#                        which_file_row = self.full_secondary.iloc[[sec_idx], [1]]
#                        # number according to which_file_row
#                        new_df = self.main_excel_df.iloc[[main_idx]].assign(Vieta=which_file_row[1][sec_idx])
#                        self.found_rows_list.append(new_df)
#                        self.found_status = True
#            if self.found_rows_list:
#                self.make_found_excel(directory)
#        except:
#            print(traceback.format_exc())

    def row_check(self, main_row: str, secondary_row: str, ):
        return main_row.lower() in secondary_row.lower()  # either in or ==

    def input_column_name_check(self, main_row: str):
        return self.input_column_name not in main_row.lower()


Parsing = DirectoryTree()
