from openpyxl import load_workbook
from openpyxl import Workbook
import os
import control_variables as cv

def get_workbook_path():
    return cv.stock_code + '_' + cv.stock_name + '.xlsx'

def open_workbook():
    wb_path = get_workbook_path()
    if os.path.exists(wb_path):
        return load_workbook(wb_path)
    else:
        return Workbook()

def close_workbook(wb):
    wb.save(get_workbook_path())
