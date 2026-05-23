import os
import random
import pandas as pd
from datetime import datetime, timedelta

# 1. Establish the validation backup testing environment directory
output_dir = "QC Sheet Backup"
os.makedirs(output_dir, exist_ok=True)

random.seed(42)  # Maintain stable distributions for testing runs

# Sample Data Pools for generating logical record overlaps
custodians = ["John Doe", "Jane Smith", "Alice Johnson", "Bob Brown", "Charlie Green"]
extensions = ["xlsx", "pdf", "docx", "csv", "msg"]
categories = ["Individual", "Employee"]

first_names = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda", "William", "Elizabeth"]
middle_names = ["Allen", "Lee", "Ann", "Marie", "Lynn", "James", "Robert", "John", "William", "None"]
last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez"]
suffixes = ["Jr.", "Sr.", "III", "None", "None", "None", "None"]

def generate_ssn():
    return f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(1000, 9999)}"

def generate_dob():
    start_date = datetime(1960, 1, 1)
    end_date = datetime(2005, 12, 31)
    random_days = random.randrange((end_date - start_date).days)
    return (start_date + timedelta(days=random_days)).strftime("%Y-%m-%d")

# Capture system timestamp for specific HHMM naming output requirement
current_time_str = datetime.now().strftime("%H%M")

# 2. Cycle out our 10 targeted test files
for file_idx in range(1, 11):
    data_rows = []
    
    # Track Document ID sequences dynamically across files to ensure unique document sets
    doc_id_counter = (file_idx - 1) * 3 + 1
    
    # Establish a bucket of 3 to 5 overlapping documents for this specific file
    file_docs = []
    for _ in range(random.randint(3, 5)):
        doc_id = f"REL{doc_id_counter:08d}"
        custodian = random.choice(custodians)
        ext = random.choice(extensions)
        f_name = f"Evidence_Leaked_Doc_{doc_id_counter}.{ext}"
        file_docs.append((doc_id, custodian, ext, f_name))
        doc_id_counter += 1
        
    # Generate exactly 100 rows of structural entity references
    for row_idx in range(1, 101):
        global_row_id = ((file_idx - 1) * 100) + row_idx
        breach_id = f"KSLT{global_row_id:08d}"
        
        # Pick from the pre-generated document bucket to guarantee document relationships exist
        doc_id, custodian, ext, f_name = random.choice(file_docs)
        
        cat = random.choice(categories)
        f_first = random.choice(first_names)
        
        f_mid = random.choice(middle_names)
        f_mid_val = "" if f_mid == "None" else f_mid
        
        f_last = random.choice(last_names)
        
        f_suf = random.choice(suffixes)
        f_suf_val = "" if f_suf == "None" else f_suf
        
        # Compute uniform full name string
        full_name = f"{f_first} {f_mid_val} {f_last}".replace("  ", " ").strip()
        ssn = generate_ssn()
        dob = generate_dob()
        
        data_rows.append([
            breach_id, doc_id, custodian, ext, f_name, cat, full_name, f_first, f_mid_val, f_last, f_suf_val, ssn, dob
        ])
        
    # Standardized header definitions matching your EXACT production format layout rule
    row1_banners = [
        "Breach Tracking", "", "", "","", "Category", 
        "Individual Information", "", "", "", "",
        "Identification information", ""
    ]
    row2_headers = [
        "Breach ID", "Document ID", "Custodian", "File extension", 
        "File Name", "Category", "Full Name", "First name", "Middle Name", "Last Name", "Suffix", 
        "SSN", "DOB"
    ]
    
    # Stitch structural frames together cleanly
    all_sheet_rows = [row1_banners, row2_headers] + data_rows
    df_matrix = pd.DataFrame(all_sheet_rows)
    
    # Apply precise requested naming structure: 20260523_HHMM_Project Test_KSTL_x.xlsx
    output_filename = os.path.join(output_dir, f"20260523_{current_time_str}_Project Test_KSTL_{file_idx}.xlsx")
    df_matrix.to_excel(output_filename, header=False, index=False)

print(f"✔ Done! Generated 10 test files with exact matrix matching schema headers inside '{output_dir}/'")