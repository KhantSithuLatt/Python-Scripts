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

def build_html_header(file_id, page_num, current_date, is_first=False):
    pb_css = "" if is_first else "style='page-break-before: always;'"
    html = f"""
    <div class="page" {pb_css}>
        <div class="header">
            <span>APEX GLOBAL SOLUTIONS INC. | AUDIT DIVISION</span>
            <span style="float: right;">RUN DATE: {current_date}</span>
        </div>
        <hr/>
        <h3>DATA DUMP REPORT - SOURCE ID: {file_id}</h3>
        <div class="divider-double">============</div>
        <pre class="table-headers">{'FULL NAME'.ljust(22)}{'SSN'.ljust(13)}{'DOB'.ljust(12)}{'PHONE'.ljust(16)}{'EMAIL'}</pre>
        <div class="divider-single">------------</div>
        <pre class="data-block">"""
    return html

def create_html_file(filepath, file_id, target_pages):
    current_date = datetime.now().strftime("%Y-%m-%d")
    records_per_page = 42
    total_records = target_pages * records_per_page - random.randint(5, 15)
    
    html_start = """<!DOCTYPE html>
    <html>
    <head>
        <style>
            body { font-family: 'Courier New', monospace; font-size: 12px; color: #333; margin: 40px; }
            .page { margin-bottom: 50px; position: relative; }
            .header { font-size: 11px; font-weight: bold; }
            .footer { font-size: 11px; margin-top: 15px; border-top: 1px solid #333; padding-top: 5px; }
            h3 { margin: 5px 0; }
            .divider-double { letter-spacing: -1px; margin-bottom: 5px; font-weight: bold; }
            .divider-single { letter-spacing: -1px; margin-top: 5px; margin-bottom: 5px; }
            .table-headers { margin: 0; font-weight: bold; color: #000; }
            .data-block { margin: 0; line-height: 1.4; }
        </style>
    </head>
    <body>"""
    
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html_start)
        
        page_num = 1
        f.write(build_html_header(file_id, page_num, current_date, is_first=True))
        
        lines_on_page = 0
        for i in range(total_records):
            name = f"{random.choice(first_names)} {random.choice(last_names)}"
            ssn, dob, phone, email = gen_ssn(), gen_dob(), gen_phone(), gen_email(name)
            
            if random.random() < 0.15:
                cp = random.choice(['ssn', 'email', 'phone'])
                if cp == 'ssn': ssn = ssn.replace("-", "")
                elif cp == 'email': email = email.replace("@", "_at_")
                elif cp == 'phone': phone = phone[:5]
            
            row_str = f"{name.ljust(22)}{ssn.ljust(13)}{dob.ljust(12)}{phone.ljust(16)}{email}\n"
            f.write(row_str)
            lines_on_page += 1
            
            if lines_on_page >= records_per_page and i < total_records - 1:
                # Close data pre-block and seal footer
                f.write(f"""</pre>
                <div class="footer">
                    <span>CONFIDENTIAL RECORD USE ONLY</span>
                    <span style="float: right;">PAGE {page_num}</span>
                </div>
                </div>""")
                
                page_num += 1
                f.write(build_html_header(file_id, page_num, current_date))
                lines_on_page = 0
                
        # Seal final remaining page tags
        f.write(f"""</pre>
        <div class="footer">
            <span>CONFIDENTIAL RECORD USE ONLY</span>
            <span style="float: right;">PAGE {page_num}</span>
        </div>
        </div>
        </body>
        </html>""")

if __name__ == "__main__":
    out_dir = "HTML Extraction"
    os.makedirs(out_dir, exist_ok=True)
    page_targets = [3, 5, 2, 8, 4, 11, 3, 6, 12, 5]
    for i in range(1, 11):
        fid = f"REL{str(i).zfill(6)}"
        create_html_file(os.path.join(out_dir, f"{fid}.html"), fid, page_targets[i-1])
    print("Completed: 10 files saved to './HTML Extraction/'")