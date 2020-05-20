import pyperclip
import xlrd

path = ("Enter path to excel file")

fp = xlrd.open_workbook(loc)
sheet = fp.sheet_by_index(0)        # Open the first sheet

