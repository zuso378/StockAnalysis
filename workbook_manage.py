from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import control_variables as cv


class Workbook_Manage:
    __workbook = None
    __worksheet = None
    __workbook_path = ''

    def __init__(self, ws_n):
        self.__workbook_path = self.__get_workbook_path()
        self.__open_workbook()
        if ws_n in self.__workbook.sheetnames:
            self.__worksheet = self.__workbook.get_sheet_by_name(ws_n)
            self.write_array([])
            self.write_array([])
        else:
            self.__worksheet = self.__workbook.create_sheet(ws_n)

    def __del__(self):
        self.__close_workbook()

    def write_dataframe(self, df):
        for row in dataframe_to_rows(df):
            self.write_array(row)

    def write_string(self, str):
        self.write_array([str])

    def write_array(self, arr):
        self.__worksheet.append(arr)

    def __get_workbook_path(self):
        return cv.common_fname + '.xlsx'

    def __open_workbook(self):
        if os.path.exists(self.__workbook_path):
            self.__workbook = load_workbook(self.__workbook_path)
        else:
            self.__workbook = Workbook()

    def __close_workbook(self):
        self.__workbook.save(self.__workbook_path)

    
