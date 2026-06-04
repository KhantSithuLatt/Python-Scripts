import os
import random
from datetime import datetime, timedelta

# --- SHARED DATA GENERATORS ---
first_names = ["John", "Jane", "Michael", "Emily", "David", "Sarah", "James", "Jessica", "Robert", "Karen"]
last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Miller", "Davis", "Garcia", "Rodriguez", "Wilson"]

def gen_ssn(): return f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(1000, 9999)}"
def gen_dob(): return (datetime(1970, 1, 1) + timedelta(days=random.randrange(13149))).strftime("%Y-%m-%d")
def gen_phone(): return f"({random.randint(100, 999)}) {random.randint(100, 999)}-{random.randint(1000, 9999)}"
def gen_email(n): return f"{n.lower().replace(' ', '.')}@{random.choice(['example.com', 'testmail.net', 'mockdata.org'])}"

def build_txt_page(file_id, page_num, current_date):
    sb = []
    sb.append(f"APEX GLOBAL SOLUTIONS INC. | AUDIT DIVISION".ljust(50) + f"RUN DATE: {current_date}".rjust(35))
    sb.append("_" * 85)
    sb.append(f"DATA DUMP REPORT - SOURCE ID: {file_id}")
    sb.append("=" * 85)
    sb.append(f"{'FULL NAME'.ljust(22)}{'SSN'.ljust(13)}{'DOB'.ljust(12)}{'PHONE'.ljust(16)}{'EMAIL'}")
    sb.append("-" * 85)
    return "\n".join(sb) + "\n"

def create_txt_file(filepath, file_id, target_pages):
    current_date = datetime.now().strftime("%Y-%m-%d")
    records_per_page = 42
    total_records = target_pages * records_per_page - random.randint(5, 15)
    
    with open(filepath, "w", encoding="utf-8") as f:
        page_num = 1
        f.write(build_txt_page(file_id, page_num, current_date))
        
        lines_on_page = 0
        for i in range(total_records):
            name = f"{random.choice(first_names)} {random.choice(last_names)}"
            ssn, dob, phone, email = gen_ssn(), gen_dob(), gen_phone(), gen_email(name)
            
            if random.random() < 0.15: # 15% Error Corruption
                cp = random.choice(['ssn', 'email', 'phone'])
                if cp == 'ssn': ssn = ssn.replace("-", "")
                elif cp == 'email': email = email.replace("@", "_at_")
                elif cp == 'phone': phone = phone[:5]
            
            row = f"{name.ljust(22)}{ssn.ljust(13)}{dob.ljust(12)}{phone.ljust(16)}{email}\n"
            f.write(row)
            lines_on_page += 1
            
            if lines_on_page >= records_per_page and i < total_records - 1:
                f.write(f"\n" + "_" * 85 + f"\nCONFIDENTIAL RECORD USE ONLY".ljust(60) + f"PAGE {page_num}\n")
                f.write("\x0c") # Form-feed page separation token
                page_num += 1
                f.write(build_txt_page(file_id, page_num, current_date))
                lines_on_page = 0
                
        f.write(f"\n" + "_" * 85 + f"\nCONFIDENTIAL RECORD USE ONLY".ljust(60) + f"PAGE {page_num}\n")

if __name__ == "__main__":
    out_dir = "TXT Extraction"
    os.makedirs(out_dir, exist_ok=True)
    page_targets = [3, 5, 2, 8, 4, 11, 3, 6, 12, 5]
    for i in range(1, 11):
        fid = f"REL{str(i).zfill(6)}"
        create_txt_file(os.path.join(out_dir, f"{fid}.txt"), fid, page_targets[i-1])
    print("Completed: 10 files saved to './TXT Extraction/'")