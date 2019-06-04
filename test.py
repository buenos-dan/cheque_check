import xlrd
import xlwt
from xlutils.copy import copy

f = xlrd.open_workbook("data.xlsx")
ff = copy(f)
ws = ff.get_sheet(0)
ws.write(390,0,"")
ff.save("data.xlsx")
