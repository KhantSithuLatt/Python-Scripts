from pathlib import Path
import pandas as pd

# Setup path and find files
folder_path = Path('./excel_test_10')
all_files = sorted(folder_path.glob('REL*.xlsx'))

#array to add all data
all_row_list = []

# --- THE MASTER FILE LOOP ---
for file_path in all_files:
    print(f"📂 Opening File: {file_path.name}")
    
    # Read the current file
    current_sheets_dict = pd.read_excel(file_path, sheet_name=None)

# STEP 1: Process and "Normalize" each sheet individually
    for sheet_name, df in current_sheets_dict.items():
        # A. Standardize to lowercase and remove hidden spaces
        df.columns = [str(col).strip().lower() for col in df.columns] # type: ignore
        
        # B. The Unifier: Rename messy columns to match our "Final Order"
        df = df.rename(columns={
            'email address': 'email',
            'customer name': 'full name',
            'name': 'full name',
            'cell': 'phone',
            'phone number': 'phone',
            'dob': 'date of birth',
            'birthday': 'date of birth'
        })
        
        # C. Add Tracking (Stamp the data)
        df['doc_id'] = file_path.name
        df['sheet_name'] = sheet_name
        
        # D. Collect the clean sheet
        all_row_list.append(df)

    # STEP 2: Combine all clean sheets into one Master Table
    # This happens OUTSIDE the loop
    master_df = pd.concat(all_row_list, ignore_index=True)

# STEP 3: Reorder and Filter
# This throws away any column not in the list (like the old 'email address' column)
final_order = ['doc_id', 'sheet_name', 'full name', 'date of birth', 'email', 'phone']
master_df = master_df[final_order]

# STEP 4: Generate the KSTL Breach ID
total_rows = len(master_df)
formattedID = [f'KSTL{n:08d}' for n in range(1, total_rows + 1)]
master_df.insert(0, 'breach_id', formattedID)

# STEP 5: Export to CSV
master_df.to_csv('final_master_for_First_File_test_2.csv', index=False)

print(f"✅ Success! Processed {len(all_row_list)} sheets.")
print(f"📂 Final file contains {len(master_df)} rows and {len(master_df.columns)} columns.")