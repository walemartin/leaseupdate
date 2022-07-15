#!/usr/bin/env python
import pandas as pd
import openpyxl
#from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
import string
from openpyxl.utils import get_column_letter
import smtplib
import ssl
from email.message import EmailMessage


# Give the location of the file
path = "leasebook.xlsx"
 
# workbook object is created
wb_obj = openpyxl.load_workbook(path)


sheet_obj = wb_obj.active
# cell references (original spreadsheet) 
min_column = wb_obj.active.min_column
max_column = wb_obj.active.max_column
min_row = wb_obj.active.min_row
max_row = wb_obj.active.max_row

print(min_column,max_column,min_row,max_row)
sheet_obj.auto_filter.ref="A1:H1"

for rowNum in range(2,sheet_obj.max_row+1):
    n='=DATEDIF(D{},E{},"D")'.format(rowNum,rowNum)
    m ='=IF(TODAY()>E{},"Rent Expired","Active")'.format(rowNum,rowNum)
    sheet_obj.cell(row = rowNum, column = 6).value = n
    sheet_obj.cell(row=rowNum,column=7).value='=E{}-TODAY()'.format(rowNum,rowNum)
    sheet_obj.cell(row=rowNum,column=8).value=m
    #sheet_obj.cell(row = rowNum, column = 6).value ='=DATEDIF(D2,E2,"D")'


wb_obj.save(path)
wb_obj.close()
print("Excel sheet data updated successfully")




#df = pd.read_excel('C:\\Users\\OLAWALE MARTINS\\Documents\\leasebook.xlsx')
#print(df.head())





