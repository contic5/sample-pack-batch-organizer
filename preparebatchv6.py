import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from pathlib import Path
import functools
import numbers

#Pandas made this much more efficient. Ignore V5.

#SETTINGS SECTION
#MODIFY THIS SECTION

#Set the name of your Excel file here. 
# Do not include the space or the number in the name.
# Do not include the .xlsx in the name.
name="Batch"
#If you do not set the name, then you will be unable to read the original excel file.

#This is the first batch you want to complete and the last batch you want to complete.
min_batch=2813
max_batch=2813
if min_batch>max_batch:
    min_batch,max_batch=max_batch,min_batch

#This is the text that will be added on to the batch file name.
extendname=""

#END OF MODIFIABLE SECTION

main_columns=["#","Reference","Ship To Company","Batch ID","Batch Name"]

#Using Pandas to make organization easier
def get_batch_name(row):
    if isinstance(row["Batch Name"],pd.Timestamp):
        return row["Batch Name"]
    batch_name=str(row["Batch Name"])
    batch_name=batch_name.split(" ")[0]
    batch_year=batch_name.split("/")[2]

    #Handle miswritten years. Years should be written as YYYY, instead of YY.
    if len(batch_year)<4:
        batch_year="20"+batch_year
        size=len(batch_name)
        batch_name=batch_name[0:size-2]+batch_year
    batch_name=pd.to_datetime(batch_name,format="%m/%d/%Y")
    batch_name= batch_name.strftime('%m/%d/%Y')
    return batch_name

def writedata(data_sheet,name,target_folder="Batches_Completed"):
    letters="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title="Sheet1"
    fontwidth=1

    maxlengths=[3 for _ in range(len(data_sheet.columns))]
    for i in range(len(maxlengths)):
        cell=letters[i]+"1"
        sheet[cell] =data_sheet.columns[i]
        maxlengths[i]=max(maxlengths[i],len(data_sheet.columns[i])+4)

    rowon=2
    for i in range(len(data_sheet)):
        for j in range(len(data_sheet.columns)):
            cell=letters[j]+str(rowon)
            val=data_sheet.iloc[i].iloc[j]
            if pd.isna(val):
                continue
            if isinstance(val,(int,float,numbers.Integral,pd.Timestamp)):
                maxlengths[j]=max(maxlengths[j],len(str(val)))
            else:
                maxlengths[j]=max(maxlengths[j],len(val))
            
            sheet[cell] =val
        rowon+=1
    rowon-=1

    tab = Table(displayName="Table1", ref="A1:"+letters[len(data_sheet.columns)-1]+str(rowon))

    style = TableStyleInfo(name="TableStyleLight8",showRowStripes=False)
    tab.tableStyleInfo = style
    sheet.add_table(tab)

    for i in range(len(maxlengths)):
        sheet.column_dimensions[letters[i]].width = fontwidth*(maxlengths[i])
    workbook.save(f"{target_folder}/"+name+extendname+".xlsx",)

def main():
    for i in range(min_batch,max_batch+1):
        file_name=f"Batches/{name} {i}.xlsx"
        path = Path(file_name)
        if(path.is_file()):
            print("Getting",path)
            sheet=pd.read_excel(file_name)
            sheet["Batch Name"]=sheet.apply(get_batch_name,axis=1)
            sheet=sheet[sheet["SKU/Quantity"]=="21354(1)"]
            sheet=sheet[main_columns]
            sheet=sheet.sort_values(by="Ship To Company")
            writedata(sheet,f"{name} {i}")
        else:
            print(path,"does not exist")
    print("Batches Completed")

if __name__=="__main__":
    main()