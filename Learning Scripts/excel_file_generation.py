import pandas as pd
import random
import os

# Configuration
file_count = 10
sheets_per_file = 10
rows_per_sheet = 100

# Variations for headers to test your mapping logic
header_variants = [
    ["Full Name", "Date of Birth", "Email Address", "Phone Number"],
    ["full name", "dob", "email", "phone"],
    ["NAME", "BIRTHDAY", "EMAIL", "PHONE NUMBER"],
    ["Customer Name", "DOB", "Email", "Phone"],
    ["Name", "Date of Birth", "Email Address", "Cell"]
]

first_names = ["John", "Jane", "Michael", "Emily", "Robert", "Linda", "William", "Barbara"]
last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis"]
suffixes = ["Jr.", "Sr.", "III", "IV", ""]
domains = ["gmail.com", "yahoo.com", "outlook.com", "company.net", "service.fr", "web.de", "info.org"]

def generate_fake_data(rows):
    data = []
    for _ in range(rows):
        fn = random.choice(first_names)
        ln = random.choice(last_names)
        mid = random.choice("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        suf = random.choice(suffixes)
        
        # Name format: Last Suffix, First Middle
        full_name = f"{ln} {suf}, {fn} {mid}.".replace(" ,", ",")
        
        dob = f"{random.randint(1,12):02d}/{random.randint(1,28):02d}/{random.randint(1950,2010)}"
        email = f"{fn.lower()}.{ln.lower()}{random.randint(1,99)}@{random.choice(domains)}"
        phone = f"{random.randint(200,999)}-{random.randint(200,999)}-{random.randint(1000,9999)}"
        
        data.append([full_name, dob, email, phone])
    return data

# Start Generating
print("🚀 Starting file generation...")

for i in range(1, file_count + 1):
    file_name = f"REL{i:08d}.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s in range(1, sheets_per_file + 1):
            sheet_name = f"Sheet_{s}"
            
            # Pick a random header variation
            headers = random.choice(header_variants)
            df = pd.DataFrame(generate_fake_data(rows_per_sheet), columns=headers)
            
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"✅ Created: {file_name}")

print("\nAll 10 files generated in your current directory!")