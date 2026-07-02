import pandas as pd

#The first sample pack batch I completed
start=2129

end=9999

#The sample pack Batch organizer has organized over 170 batches and over 12500 sample packs.

def main():
    total_batches=0
    total_sample_packs=0

    for batch_number in range(start,end+1):
        try:
            df=pd.read_excel(f"Batches_Completed/Batch {batch_number}.xlsx")
            total_batches+=1
            total_sample_packs+=len(df)
        except FileNotFoundError:
            pass
    print(f"The Sample Pack Batch Organizer has organized {total_batches} Batches and {total_sample_packs} Sample Packs")

if __name__=="__main__":
    main()