import os
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo

start_number=2176
safe_overwrite=False
using_excel=True

folder = "Batches_To_Rename"
for count, filename in enumerate(os.listdir(folder)):
    batch_num=start_number+count-1
    if(using_excel):
        workbook=openpyxl.load_workbook(folder+"/"+filename)
        sheet = workbook.active
        cell_obj = sheet.cell(row = 3, column = 6)

        #Get batch number from cell E3. Cell row 3 column 5
        batch_num=int(cell_obj.value)
    
    src=folder+"/"+filename
    if(safe_overwrite):
        dst=folder+"/"+"Batch "+str(batch_num)+" 1.xlsx"
    else:
        dst=folder+"/"+"Batch "+str(batch_num)+".xlsx"
    os.rename(src, dst)