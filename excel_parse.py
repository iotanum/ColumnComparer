import os
import string
import time
import pandas as pd
import numpy as np
from pathlib import Path
import traceback


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
        self.full_secondaries_df = None
        self.main_file_header = None
        self.full_secondary = None  # if doing specific tasks
        self.full_sec_hashtable = {}

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
            for idx, column_row in enumerate(xl[column_num]):
                if all(x.lower() in str(column_row).lower() for x in column_name):
                    df = pd.DataFrame(xl[column_num])  # specific task - delete dropna()
                    df.columns = [self.input_column_name]
                    if not main:
                        series = df.assign(file=pd.Series([file_name for _ in
                                                           range(len(df[self.input_column_name]) + 50)]))
#                       maybe I'll need it later for specific tasks
                        self.full_secondary = xl
                        self.secondary_to_hashtable(xl)
                        self.secondary_column_list.append(series)
                        return self.column_alphabetic_value(column_num), column_row
                    else:
                        self.main_excel_df = xl
                        self.main_file_header = list(xl.iloc[[idx]].values[0])
                        self.main_file_header.append("Rasta dokumente/dokumentuose")
                        self.main_excel_needed_column = xl[column_num]
                        return self.column_alphabetic_value(column_num), column_row

    def secondary_to_hashtable(self, full_secondary):
        for value in full_secondary.values:
            try:
                value = str(value)
                # nums according to needs
                self.full_sec_hashtable[value[9]] = [value[3], value[4]]
            except:
                print(traceback.format_exc())

    def column_alphabetic_value(self, column_num):
        return string.ascii_uppercase[column_num]

    def export_excel(self, directory, df, found=True):
        self.result_file_name = f'{"rezultatai" if found else "nerasti rezultatai"} ' + \
                                                                str(time.strftime("%Y%m%d %H%M%S")) + ".xlsx"
        df.to_excel(directory + "/" + self.result_file_name)

    def make_excel(self, directory):
        self.full_secondaries_df = pd.concat(self.secondary_column_list, ignore_index=True)
        self.compare_main_with_secondaries(directory, self.full_secondary)  # specific task, full secondary

    def make_found_excel(self, directory):
        # for specific tasks remove 3rd and 4th, and last line
        df = pd.concat(self.found_rows_list, ignore_index=True)
        df.index += 1
        df.columns = self.main_file_header
        print("Found:", len(df))
        df2 = pd.concat(self.not_found_list, ignore_index=True)
        df2.columns = self.main_file_header[:-1]
        df2.index += 1
        print("Didn't find:", len(df2))
        self.export_excel(directory, df)
        self.export_excel(directory, df2, found=False)

#    def compare_main_with_secondaries(self, directory):
#        print("Full file", len(self.main_excel_needed_column) - 1)  # minus header
#        for main_idx, column_row in enumerate(self.main_excel_needed_column):
#            column_row = str(column_row)
#            self.row_status = False
#            for sec_idx, secondary_column_row in enumerate(self.full_secondaries_df[self.input_column_name]):
#                secondary_column_row = str(secondary_column_row)
#                if self.row_check(column_row, secondary_column_row) and self.input_column_name_check(column_row)\
#                        and column_row not in self.duplicates and column_row != "nan"\
#                        and column_row not in self.main_file_header:
#                    list_of_files = self.find_multiple_file(column_row)
#                    new_df = self.main_excel_df.iloc[[main_idx]].assign(Rasta_dokumente=list_of_files)
#                    self.duplicates.add(str(column_row))
#                    self.found_rows_list.append(new_df)
#                    self.row_status = True
#            if not self.row_status and self.input_column_name_check(column_row) and column_row != "nan"\
#                    and column_row not in self.main_file_header:
#                not_found_row = self.main_excel_df.iloc[[main_idx]]
#                self.not_found_list.append(not_found_row)
#        if self.found_rows_list:
#            self.found_status = True
#            self.make_found_excel(directory)

    def find_multiple_file(self, column_row):
        list_of_files = set()  # referred as a list of files that has the needed row
        for idx, row in enumerate(self.full_secondaries_df[self.input_column_name]):
            row = str(row)
            if self.row_check(column_row, row):
                which_file_row = self.full_secondaries_df.iloc[[idx], [1]]
                which_file_row = which_file_row['file'][idx]
                list_of_files.add(which_file_row)
            continue
        return ",\n".join([str(item) for item in list_of_files])

#    def compare_main_with_secondaries(self, directory, full_secondary):
#        try:
#            for main_idx, column_row in enumerate(self.main_excel_needed_column):
#                print(column_row)
#                column_row = str(column_row)
#                if column_row in self.full_sec_hashtable.keys():
#                    print(self.full_sec_hashtable[column_row])
#        except:
#            print(traceback.format_exc())
#

# For future, specific task
    def compare_main_with_secondaries(self, directory, sec_df):
        try:
            for main_idx, column_row in enumerate(self.main_excel_needed_column):
                column_row = str(column_row)
                for sec_idx, secondary_column_row in enumerate(sec_df[9]):  # kurio stulpelio ieskoma
                    secondary_column_row = str(secondary_column_row)
                    if self.row_check(column_row, secondary_column_row) and self.input_column_name_check(column_row)\
                            and str(secondary_column_row) != 'nan':
                        # number according to what is needed
                        which_file_row = self.full_secondary.iloc[[sec_idx], [1]]
                        which_file_row2 = self.full_secondary.iloc[[sec_idx], [3]]
                        which_file_row3 = self.full_secondary.iloc[[sec_idx], [4]]
                        # number according to which_file_row
                        new_df = self.main_excel_df.iloc[[main_idx]].assign(Vieta=which_file_row[1][sec_idx])
                        # apacioje numeris pagal which file row
                        new_df = self.main_excel_df.iloc[[main_idx]].assign(Dez=which_file_row2[3][sec_idx])
                        new_df = self.main_excel_df.iloc[[main_idx]].assign(Byl=which_file_row3[4][sec_idx])
                        self.found_rows_list.append(new_df)
                        self.found_status = True
            if self.found_rows_list:
                self.make_found_excel(directory)
        except:
            print(traceback.format_exc())

    def row_check(self, main_row: str, secondary_row: str, ):
        return main_row.lower() in secondary_row.lower()  # either in or ==

    def input_column_name_check(self, main_row: str):
        return self.input_column_name not in main_row.lower()


Parsing = DirectoryTree()
