import pandas as pd
from pathlib import Path
from typing import List
# Importing our custom modules
from validators import clean_headers, validate_names, validate_emails

# 1. Setup
folder_path: Path = Path('./excel_test_10')
all_files: List[Path] = sorted(folder_path.glob('REL*.xlsx'))
data_holder: List[pd.DataFrame] = []
error_holder: List[pd.DataFrame] = []

# 2. Iteration (Outer Loop)
for file_path in all_files:
    print(f"Processing: {file_path.name}")
    current_workbook = pd.read_excel(file_path, sheet_name=None)

    # Inner Loop (Sheet processing)
    for sheet_name, df in current_workbook.items():
        if df.empty:
            print(f"   [!] Skipping empty sheet: {sheet_name}")
            continue

        # --- VALIDATION PIPELINE ---
        # Step A: Normalize Headers
        df = clean_headers(df)
        df = df.dropna(how='all')

        # Step B: Name Validation
        df, name_errs = validate_names(df, file_path.name, sheet_name)
        error_holder.extend(name_errs)

        # Step C: Email Validation
        df, email_errs = validate_emails(df, file_path.name, sheet_name)
        error_holder.extend(email_errs)

        # Step D: Final Stamp & Collect
        df['doc id'] = file_path.name
        df['sheet name'] = sheet_name
        data_holder.append(df)

# 3. Final Exports (The "Grand Finale")
if data_holder:
    master_df = pd.concat(data_holder, ignore_index=True)
    # Ensure column order matches our requirements
    columns_to_keep = ['doc id', 'sheet name', 'name', 'date of birth', 'email', 'phone']
    master_df = master_df[columns_to_keep]
    
    # Generate Breach ID
    master_df.insert(0, 'breach_id', [f'KSTL{n:08d}' for n in range(1, len(master_df) + 1)])
    
    master_df.to_csv('Final_Master_File.csv', index=False)
    print(f"\n✅ {len(master_df)} Clean rows exported.")

if error_holder:
    final_errors_df = pd.concat(error_holder, ignore_index=True)
    with pd.ExcelWriter('Manual_Check_Required.xlsx') as writer:
        final_errors_df.to_excel(writer, sheet_name='All_Errors', index=False)
    print(f"🚨 {len(final_errors_df)} Errors quarantined.")