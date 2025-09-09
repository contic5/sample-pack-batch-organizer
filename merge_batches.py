import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from pathlib import Path
import functools
import numbers
import preparebatchv6

#SETTINGS SECTION
#MODIFY THIS SECTION

#Set the name of your Excel file here. 
# Do not include the space or the number in the name.
# Do not include the .xlsx in the name.
name="Batch"

#If you do not set the name, then you will be unable to read the original excel file.
min_batch=2791
max_batch=2812

#Set this to true to keep batches in original state. Set this to false to get rid of all but the main columns
original_form=False

#This is the text that will be added on to the batch file name.
extendname=""

target_folder="Batches_Completed"

#Tracking duplicates only or tracking all sheets
tracking_duplicates_only=False

#Remove duplicates
keeping_duplicates=False

#The column to sort by
#Sort column="Ship to Company"
sort_column="Batch ID"

#END OF MODIFIABLE SECTION

def main():
    merged_sheet=pd.read_excel(f"{target_folder}/{name} {max_batch}.xlsx")

    for i in range(max_batch-1,min_batch-1,-1):
        file_name=f"{target_folder}/{name} {i}.xlsx"
        path = Path(file_name)
        if(path.is_file()):
            #print("Getting",path)
            sheet=pd.read_excel(file_name)
            merged_sheet=pd.concat([merged_sheet,sheet])
        else:
            pass
            #print(path,"does not exist")
        print(f"Merged Sheet Length {len(merged_sheet)}")
    print()

    print("Preparing Batches...")
    merged_sheet["Batch Name"]=merged_sheet.apply(preparebatchv6.get_batch_name,axis=1)
    merged_sheet["Batch Name"]=pd.to_datetime(merged_sheet["Batch Name"], errors = 'coerce')

    #merged_sheet=merged_sheet[(merged_sheet["SKU/Quantity"]=="21354(1)")]
    merged_sheet=merged_sheet.sort_values(by=sort_column)
    if(tracking_duplicates_only and keeping_duplicates):
        #We want to output only the duplicates when we are tracking duplicates.
        merged_sheet=merged_sheet[merged_sheet.duplicated(subset="Ship To Company")]

    if not keeping_duplicates:
        merged_sheet=merged_sheet.drop_duplicates(subset="Ship To Company")

    print(f"Merged Sheet Length {len(merged_sheet)}")
    if original_form:
        preparebatchv6.writedata(merged_sheet,f"Merged Raw Batches {min_batch}-{max_batch}")
        print("Merged Raw Batches Finished")
    else:
        merged_sheet=merged_sheet[preparebatchv6.main_columns]
        preparebatchv6.writedata(merged_sheet,f"Merged Batches {min_batch}-{max_batch}")
        print("Merged Completed Batches Finished")

if __name__=="__main__":
    main()