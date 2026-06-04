import os
import random
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from xml.dom import minidom

# --- SHARED DATA GENERATORS ---
first_names = ["John", "Jane", "Michael", "Emily", "David", "Sarah", "James", "Jessica", "Robert", "Karen"]
last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Miller", "Davis", "Garcia", "Rodriguez", "Wilson"]

def gen_ssn(): return f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(1000, 9999)}"
def gen_dob(): return (datetime(1970, 1, 1) + timedelta(days=random.randrange(13149))).strftime("%Y-%m-%d")
def gen_phone(): return f"({random.randint(100, 999)}) {random.randint(100, 999)}-{random.randint(1000, 9999)}"
def gen_email(n): return f"{n.lower().replace(' ', '.')}@{random.choice(['example.com', 'testmail.net', 'mockdata.org'])}"

def create_xml_file(filepath, file_id, target_pages):
    current_date = datetime.now().strftime("%Y-%m-%d")
    # Simulate data size equivalent to target page footprints
    total_records = target_pages * 42 - random.randint(5, 15)
    
    root = ET.Element("AuditLog")
    
    # Metadata Header Node
    meta = ET.SubElement(root, "Metadata")
    ET.SubElement(meta, "Company").text = "APEX GLOBAL SOLUTIONS INC."
    ET.SubElement(meta, "Division").text = "AUDIT DIVISION"
    ET.SubElement(meta, "SourceID").text = file_id
    ET.SubElement(meta, "RunDate").text = current_date
    ET.SubElement(meta, "Classification").text = "CONFIDENTIAL"
    
    records_node = ET.SubElement(root, "Records")
    
    for _ in range(total_records):
        name = f"{random.choice(first_names)} {random.choice(last_names)}"
        ssn, dob, phone, email = gen_ssn(), gen_dob(), gen_phone(), gen_email(name)
        
        if random.random() < 0.15:
            cp = random.choice(['ssn', 'email', 'phone'])
            if cp == 'ssn': ssn = ssn.replace("-", "")
            elif cp == 'email': email = email.replace("@", "_at_")
            elif cp == 'phone': phone = phone[:5]
            
        rec = ET.SubElement(records_node, "Record")
        ET.SubElement(rec, "FullName").text = name
        ET.SubElement(rec, "SSN").text = ssn
        ET.SubElement(rec, "DOB").text = dob
        ET.SubElement(rec, "Phone").text = phone
        ET.SubElement(rec, "Email").text = email

    # Pretty-print XML structure using minidom so it stays human-readable
    raw_str = ET.tostring(root, 'utf-8')
    parsed_str = minidom.parseString(raw_str)
    pretty_xml = parsed_str.toprettyxml(indent="    ")
    
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(pretty_xml)

if __name__ == "__main__":
    out_dir = "XML Extraction"
    os.makedirs(out_dir, exist_ok=True)
    page_targets = [3, 5, 2, 8, 4, 11, 3, 6, 12, 5]
    for i in range(1, 11):
        fid = f"REL{str(i).zfill(6)}"
        create_xml_file(os.path.join(out_dir, f"{fid}.xml"), fid, page_targets[i-1])
    print("Completed: 10 files saved to './XML Extraction/'")