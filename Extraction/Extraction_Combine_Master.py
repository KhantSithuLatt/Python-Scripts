import os
import re
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber
from docx import Document
from bs4 import BeautifulSoup

# ==========================================
# 1. VALIDATION AND DIAGNOSTIC PATTERNS
# ==========================================
PII_STRICT_PATTERN = re.compile(
    r'^(?P<name>.+?)\s{2,}'          
    r'(?P<ssn>\d{3}-\d{2}-\d{4})\s{2,}' 
    r'(?P<dob>\d{4}-\d{2}-\d{2})\s{2,}' 
    r'(?P<phone>\(\d{3}\)\s\d{3}-\d{4})\s{2,}' 
    r'(?P<email>\S+@\S+)$'            
)

# Individual modular patterns to pinpoint specific validation failures
SSN_CHECK = re.compile(r'\d{3}-\d{2}-\d{4}')
DOB_CHECK = re.compile(r'\d{4}-\d{2}-\d{2}')
PHONE_CHECK = re.compile(r'\(\d{3}\)\s\d{3}-\d{4}')
EMAIL_CHECK = re.compile(r'\S+@\S+')

# ==========================================
# 2. GUI INTERFACE DIALOGS
# ==========================================

def get_folder_via_dialog():
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Select Folder Containing Source Audit Logs")
    root.destroy()
    return folder_selected

class FormatSelectorDialog(tk.Tk):
    def __init__(self, available_formats):
        super().__init__()
        self.title("Select Extraction Formats")
        self.geometry("400x320")
        self.resizable(False, False)
        
        self.selected_formats = []
        self.checkbox_vars = {}
        
        lbl = tk.Label(self, text="Select file types to extract from:", font=("Arial", 11, "bold"), pady=15)
        lbl.pack()
        
        cb_frame = tk.Frame(self)
        cb_frame.pack(pady=5)
        
        for fmt in sorted(available_formats):
            self.checkbox_vars[fmt] = tk.BooleanVar()
            cb = tk.Checkbutton(cb_frame, text=f" {fmt.upper()} Files (.{fmt})", variable=self.checkbox_vars[fmt], font=("Arial", 10))
            cb.pack(anchor="w", padx=20, pady=5)
            
        btn = tk.Button(self, text="Begin Extraction", command=self.on_submit, width=18, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        btn.pack(pady=25)
        
    def on_submit(self):
        self.selected_formats = [fmt for fmt, var in self.checkbox_vars.items() if var.get()]
        if not self.selected_formats:
            messagebox.showwarning("Warning", "Please select at least one format to proceed.")
        else:
            self.destroy()

# ==========================================
# 3. ADVANCED LINE AND DIAGNOSTIC ROUTINES
# ==========================================

def is_layout_garbage(line_text):
    clean = line_text.strip()
    if not clean or clean == "\x0c":
        return True
    if any(kwd in clean for kwd in ["APEX GLOBAL", "AUDIT DIVISION", "DATA DUMP", "FULL NAME", "---", "===", "RUN DATE:", "CONFIDENTIAL RECORD"]):
        return True
    if re.match(r'^[-_=+]+$', clean):
        return True
    return False

def slice_unmatched_line(line_text):
    clean_line = line_text.strip()
    tokens = re.split(r'\s{2,}', clean_line)
    while len(tokens) < 5:
        tokens.append("")
    return tokens[:5]

def diagnose_error_type(tokens, raw_line):
    """
    Analyzes the extracted tokens of a broken row to find out
    exactly which validation format it violated.
    """
    # If the line splits into completely unexpected column counts, it's structurally malformed
    non_empty_tokens = [t for t in tokens if t.strip()]
    if len(non_empty_tokens) < 3:
        return "Malformed Data Row"
        
    # Check tokens against our individual regex rules to find the broken element
    # Standard array indices based on our 5 columns: [Name, SSN, DOB, Phone, Email]
    ssn_token = tokens[1] if len(tokens) > 1 else ""
    dob_token = tokens[2] if len(tokens) > 2 else ""
    phone_token = tokens[3] if len(tokens) > 3 else ""
    email_token = tokens[4] if len(tokens) > 4 else ""

    if ssn_token and not SSN_CHECK.search(ssn_token):
        return "SSN Validation Error"
    if dob_token and not DOB_CHECK.search(dob_token):
        return "DOB Validation Error"
    if phone_token and not PHONE_CHECK.search(phone_token):
        return "Phone Validation Error"
    if email_token and not EMAIL_CHECK.search(email_token):
        return "Email Validation Error"
        
    return "Alignment/Spacing Error"

# ==========================================
# 4. COMPONENT PARSERS PER FORMAT
# ==========================================

def parse_pdf(filepath):
    passed, missed = [], []
    with pdfplumber.open(filepath) as pdf:
        for idx, page in enumerate(pdf.pages):
            page_num = idx + 1
            page_text = page.extract_text()
            if not page_text:
                continue
            lines = page_text.split("\n")
            for line in lines:
                if is_layout_garbage(line):
                    continue
                match = PII_STRICT_PATTERN.match(line.strip())
                if match:
                    gd = match.groupdict()
                    passed.append([page_num, gd['name'], gd['ssn'], gd['dob'], gd['phone'], gd['email']])
                else:
                    exploded_fields = slice_unmatched_line(line)
                    error_type = diagnose_error_type(exploded_fields, line)
                    # Insert Error Type directly into the array payload
                    missed.append([page_num, error_type] + exploded_fields)
    return passed, missed

def parse_txt(filepath):
    passed, missed = [], []
    with open(filepath, "r", encoding="utf-8") as f:
        raw_content = f.read()
        pages = raw_content.split("\x0c")
        for idx, page_content in enumerate(pages):
            page_num = idx + 1
            lines = page_content.split("\n")
            for line in lines:
                if is_layout_garbage(line):
                    continue
                match = PII_STRICT_PATTERN.match(line.strip())
                if match:
                    gd = match.groupdict()
                    passed.append([page_num, gd['name'], gd['ssn'], gd['dob'], gd['phone'], gd['email']])
                else:
                    exploded_fields = slice_unmatched_line(line)
                    error_type = diagnose_error_type(exploded_fields, line)
                    missed.append([page_num, error_type] + exploded_fields)
    return passed, missed

def parse_docx(filepath):
    passed, missed = [], []
    current_page = 1
    doc = Document(filepath)
    for p in doc.paragraphs:
        line = p.text
        if "CONFIDENTIAL RECORD" in line and "PAGE" in line:
            page_match = re.search(r'PAGE\s+(\d+)', line)
            if page_match:
                current_page = int(page_match.group(1)) + 1
            continue
        if is_layout_garbage(line):
            continue
        match = PII_STRICT_PATTERN.match(line.strip())
        if match:
            gd = match.groupdict()
            passed.append([current_page, gd['name'], gd['ssn'], gd['dob'], gd['phone'], gd['email']])
        else:
            exploded_fields = slice_unmatched_line(line)
            error_type = diagnose_error_type(exploded_fields, line)
            missed.append([current_page, error_type] + exploded_fields)
    return passed, missed

def parse_html(filepath):
    passed, missed = [], []
    with open(filepath, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")
    for page_block in soup.find_all("div", class_="page"):
        footer = page_block.find("div", class_="footer")
        page_num = 1
        if footer:
            page_match = re.search(r'PAGE\s+(\d+)', footer.text)
            if page_match:
                page_num = int(page_match.group(1))
        data_container = page_block.find("pre", class_="data-block")
        if data_container:
            lines = data_container.text.split("\n")
            for line in lines:
                if is_layout_garbage(line):
                    continue
                match = PII_STRICT_PATTERN.match(line.strip())
                if match:
                    gd = match.groupdict()
                    passed.append([page_num, gd['name'], gd['ssn'], gd['dob'], gd['phone'], gd['email']])
                else:
                    exploded_fields = slice_unmatched_line(line)
                    error_type = diagnose_error_type(exploded_fields, line)
                    missed.append([page_num, error_type] + exploded_fields)
    return passed, missed

def parse_xml(filepath):
    passed, missed = [], []
    tree = ET.parse(filepath)
    root = tree.getroot()
    for rec in root.findall(".//Record"):
        name  = rec.find("FullName").text.strip() if rec.find("FullName") is not None else ""
        ssn   = rec.find("SSN").text.strip() if rec.find("SSN") is not None else ""
        dob   = rec.find("DOB").text.strip() if rec.find("DOB") is not None else ""
        phone = rec.find("Phone").text.strip() if rec.find("Phone") is not None else ""
        email = rec.find("Email").text.strip() if rec.find("Email") is not None else ""
        
        test_line = f"{name.ljust(22)}{ssn.ljust(13)}{dob.ljust(12)}{phone.ljust(16)}{email}"
        if PII_STRICT_PATTERN.match(test_line):
            passed.append([1, name, ssn, dob, phone, email])
        else:
            exploded_fields = [name, ssn, dob, phone, email]
            error_type = diagnose_error_type(exploded_fields, test_line)
            missed.append([1, error_type, name, ssn, dob, phone, email])
    return passed, missed

# ==========================================
# 5. EXECUTION CONTROLLER PIPELINE
# ==========================================

def main():
    print("[Terminal Execution initiated... Running GUI Folder Dialog]")
    target_folder = get_folder_via_dialog()
    if not target_folder:
        print("[Execution Cancelled]: No folder selected.")
        return
        
    all_files = os.listdir(target_folder)
    discovered_extensions = set(os.path.splitext(f)[1].replace(".", "").lower() for f in all_files if os.path.splitext(f)[1])
    supported_extensions = discovered_extensions.intersection({"pdf", "txt", "docx", "xml", "html"})
    
    if not supported_extensions:
        print("[Error]: Zero supported file variations identified in target space.")
        return
        
    selector = FormatSelectorDialog(supported_extensions)
    selector.mainloop()
    
    chosen_formats = selector.selected_formats
    if not chosen_formats:
        print("[Execution Cancelled]: Formats skipped.")
        return
        
    clean_master_records = []
    quarantine_records = []
    
    for filename in sorted(all_files):
        name_part, ext_part = os.path.splitext(filename)
        ext = ext_part.replace(".", "").lower()
        if ext not in chosen_formats:
            continue
            
        filepath = os.path.join(target_folder, filename)
        doc_id = name_part
        print(f"Extracting: {filename}...")
        
        if ext == "pdf":    passed, missed = parse_pdf(filepath)
        elif ext == "txt":  passed, missed = parse_txt(filepath)
        elif ext == "docx": passed, missed = parse_docx(filepath)
        elif ext == "html": passed, missed = parse_html(filepath)
        elif ext == "xml":  passed, missed = parse_xml(filepath)
        
        # Format Clean Records
        for row in passed:
            clean_master_records.append([doc_id, ext, row[0], row[1], row[2], row[3], row[4], row[5]])
            
        # Format Misaligned Diagnostic Records
        for row in missed:
            # row format here is: [page_num, error_type, name, ssn, dob, phone, email]
            quarantine_records.append([doc_id, ext, row[0], row[1], row[2], row[3], row[4], row[5], row[6]])

    # ---- EXPORT CLEAN DATA RUN ----
    if clean_master_records:
        df_clean = pd.DataFrame(
            clean_master_records,
            columns=["Document ID", "File Extension", "Page Number", "Full Name", "SSN", "DOB", "Phone", "Email"]
        )
        breach_ids = [f"KSTL{str(1 + idx).zfill(8)}" for idx in range(len(df_clean))]
        df_clean.insert(0, "Breach ID", breach_ids)
        df_clean.to_excel("Consolidated_Clean_Extraction.xlsx", index=False)
        print(f" -> Output Generated: Consolidated_Clean_Extraction.xlsx ({len(df_clean)} Rows)")

    # ---- EXPORT TABULAR QUARANTINE RUN ----
    if quarantine_records:
        df_missed = pd.DataFrame(
            quarantine_records,
            columns=["Document ID", "File Extension", "Page Number", "Error Type", "Full Name", "SSN", "DOB", "Phone", "Email"]
        )
        breach_ids_missed = [f"ERR{str(1 + idx).zfill(8)}" for idx in range(len(df_missed))]
        df_missed.insert(0, "Error ID", breach_ids_missed)
        df_missed.to_excel("Manual_Check_Required.xlsx", index=False)
        print(f" -> Diagnostics Generated: Manual_Check_Required.xlsx ({len(df_missed)} Rows Captured)")
    else:
        print(" -> Success: 100% extraction accuracy achieved. Zero files required quarantine.")

if __name__ == "__main__":
    main()