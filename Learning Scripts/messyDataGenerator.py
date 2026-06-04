import pandas as pd
import numpy as np
from pathlib import Path

# Create directory
Path('./excel_test_10').mkdir(exist_ok=True)

names = ["John Doe", "Jane Smith", "Mike Ross", "Harvey Specter", "Donna Paulsen"]
emails = ["john@test.com", "jane@test.com", "invalid-email", "mike@pearson.veo", "donna@firm.com"]
phones = ["1234567890", "9876543210", "123", "None", "555-0199"]

for i in range(1, 11):
    file_path = f'./excel_test_10/REL{i:08d}.xlsx'
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        for s in range(1, 11):
            # Create a basic dataframe
            data = {
                'Full Name': names,
                'Email': emails,
                'Phone': phones,
                'DOB': ['01/01/1990'] * 5
            }
            df = pd.DataFrame(data)

            # --- INJECTING ERRORS ---
            if s == 3: # Make Sheet 3 completely empty
                df = pd.DataFrame()
            elif s == 5: # Add 2 completely blank rows at the top of Sheet 5
                blank = pd.DataFrame([[np.nan] * len(df.columns)], columns=df.columns)
                df = pd.concat([blank, blank, df], ignore_index=True)
            elif s == 7: # Change headers to be messy
                df.columns = ['customer', 'E-MAIL ADDRESS', 'cell', 'date_of_birth'] # type: ignore
            
            df.to_excel(writer, sheet_name=f'Sheet_{s}', index=False)

print("🧪 10 Poisoned Files generated in ./excel_test_10/")