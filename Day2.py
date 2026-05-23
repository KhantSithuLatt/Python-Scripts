from typing import List
import pandas as pd
from pathlib import Path
#import pathlib as Path

#finding files
FolderPath = Path('./excel_test_10')
AllFiles = sorted(FolderPath.glob('REL*.xlsx'))
DataHolder = []
ErrorHolder = []


#outer loop for readin each file
for file in AllFiles:
    print(f'Accessing: {file}')

    #using panda excel reader for current file for current outer loop adn store into a dictionary named CurrentFile
    CurrentFile = pd.read_excel(file, sheet_name=None)

    #Inner loop for the sheet
    for sheet, df in CurrentFile.items():
        #checking if sheet if empty
        if df.empty:
            print(f'{sheet} in {file} is empty')
        else:
            # 1. Create your clean list of strings
            clean_list: List[str] = [str(col).strip().lower() for col in df.columns]

            # 2. Explicitly cast that list back into a Pandas Index to satisfy Mypy
            # Equivalent to: df.setColumns((Index) cleanList);
            df.columns = pd.Index(clean_list)

            #Header Mapping
            df = df.rename(columns={
                'full name': 'name',       # <--- Add this (The standard)
                'customer name': 'name',   # <--- Keep this
                'customer': 'name',        # <--- For Sheet 7
                'email address': 'email',
                'e-mail address': 'email', # <--- For Sheet 7
                'cell': 'phone',           # <--- For Sheet 7
                'phone number': 'phone',
                'dob':'date of birth'
            })
    
            # 1. Clean the 100% empty rows first
            df = df.dropna(how='all')

            # 2. TRAP Missing Names (BUT CHECK IF COLUMN EXISTS FIRST)
            if 'name' in df.columns:
                BadRows = df[df['name'].isna()].copy()
                if not BadRows.empty:
                    BadRows['error type'] = 'Missing Name'
                    BadRows['doc id'] = file.name
                    BadRows['sheet name'] = sheet
                    ErrorHolder.append(BadRows)
                df = df[df['name'].notna()]
            else:
                # This sheet is so messy it doesn't even have a name column!
                print(f"   [❌] Skipping Name Validation for {sheet}: Column 'name' not found.")

            # 3. TRAP Invalid Emails (CHECK IF COLUMN EXISTS FIRST)
            if 'email' in df.columns:
                InvalidEmail = df[~df['email'].astype(str).str.contains('@', na=False)].copy()
                if not InvalidEmail.empty:
                    InvalidEmail['error type'] = 'Missing @ in Email'
                    InvalidEmail['doc id'] = file.name
                    InvalidEmail['sheet name'] = sheet
                    ErrorHolder.append(InvalidEmail)
                df = df[df['email'].astype(str).str.contains('@', na=False)]

            #Adding Doc ID and Sheet name
            df['doc id'] = file.name
            df['sheet name'] = sheet

            #df to DataHolder list
            DataHolder.append(df)

if DataHolder:
    MasterDF = pd.concat(DataHolder, ignore_index=True)
    MasterDF = MasterDF[['doc id', 'sheet name', 'name', 'date of birth', 'email', 'phone']]
    MasterDF.insert(0, 'breach id',[f'KSTL{n:08d}' for n in range(1, len(MasterDF) + 1)])
    MasterDF.to_csv('Final_Master_File.csv', index=False)
    print(f"{len(MasterDF)} rows of clean data saved to Final_Master_File.csv")
if ErrorHolder:
    ErrorDF = pd.concat(ErrorHolder, ignore_index=True)
    with pd.ExcelWriter('Manual Check Required.xlsx') as writer:
        ErrorDF.to_excel(writer, sheet_name='Missing Name', index=False)
    print(f'{len(ErrorDF)} errors sent to Manual Check Required.xlsx')



