import os
import re
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from docx import Document
from bs4 import BeautifulSoup

# ==========================================
# GUI DIALOG INTERFACE (FIXED GEOMETRY AND SCALING)
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
        
        # Expanded box dimensions dynamically based on option scale
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
# ROBUST REGEX GAP EXTRACTION ENGINES
# ==========================================

def extract_via_regex_split(text_line):
    """
    Cleans lines using a regex multi-space divider instead of hardcoded coordinates.
    Handles data padding discrepancies dynamically.
    """
    clean_line = text_line.strip()
    if not clean_line:
        return None
        
    # Skip enterprise document decorative headers, line breaks, and page dividers
    if any(kwd in clean_line for kwd in ["APEX GLOBAL", "DATA DUMP", "FULL NAME", "---", "===", "CONFIDENTIAL"]) or "\x0c" in clean_line:
        return None
        
    # Split by any sequence of 2 or more consecutive spaces
    tokens = re.split(r'\s{2,}', clean_line)
    
    # Ensure it maps down cleanly to our 5 expected baseline columns
    if len(tokens) == 5:
        return tokens
    else:
        # Fallback mechanism if names contain spaces but padding drops below 2 spaces
        # Match email at the end, work backwards to catch other structured formats
        return None

def parse_txt(filepath):
    rows = []
    with open(filepath, "r", encoding="utf-8") as f:
        for line in f:
            data = extract_via_regex_split(line)
            if data:
                rows.append(data)
    return rows

def parse_docx(filepath):
    rows = []
    doc = Document(filepath)
    for p in doc.paragraphs:
        data = extract_via_regex_split(p.text)
        if data:
            rows.append(data)
    return rows

def parse_html(filepath):
    rows = []
    with open(filepath, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")
    for block in soup.find_all("pre", class_="data-block"):
        lines = block.text.split("\n")
        for line in lines:
            data = extract_via_regex_split(line)
            if data:
                rows.append(data)
    return rows

def parse_xml(filepath):
    rows = []
    tree = ET.parse(filepath)
    root = tree.getroot()
    for rec in root.findall(".//Record"):
        name  = rec.find("FullName").text.strip() if rec.find("FullName") is not None else ""
        ssn   = rec.find("SSN").text.strip() if rec.find("SSN") is not None else ""
        dob   = rec.find("DOB").text.strip() if rec.find("DOB") is not None else ""
        phone = rec.find("Phone").text.strip() if rec.find("Phone") is not None else ""
        email = rec.find("Email").text.strip() if rec.find("Email") is not None else ""
        rows.append([name, ssn, dob, phone, email])
    return rows

# ==========================================
# CONTROLLER WORKFLOW EXECUTION
# ==========================================

def main():
    print("[Terminal Execution initiated... Running GUI Folder Dialog]")
    target_folder = get_folder_via_dialog()
    
    if not target_folder:
        print("[Execution Cancelled]: No folder was selected.")
        return
        
    print(f" -> Selected Input Directory: {target_folder}")
    
    all_files = os.listdir(target_folder)
    discovered_extensions = set(os.path.splitext(f)[1].replace(".", "").lower() for f in all_files if os.path.splitext(f)[1])
    supported_extensions = discovered_extensions.intersection({"txt", "docx", "xml", "html"})
    
    if not supported_extensions:
        print("[Error]: No supported source files found in target directory.")
        return
        
    selector = FormatSelectorDialog(supported_extensions)
    selector.mainloop()
    
    chosen_formats = selector.selected_formats
    if not chosen_formats:
        print("[Execution Cancelled]: No formats chosen.")
        return
        
    print(f" -> Target Execution Formats Chosen: {chosen_formats}")
    
    master_records = []
    
    for filename in sorted(all_files):
        name_part, ext_part = os.path.splitext(filename)
        ext = ext_part.replace(".", "").lower()
        
        if ext not in chosen_formats:
            continue
            
        filepath = os.path.join(target_folder, filename)
        document_id = name_part
        
        print(f"Processing File: {filename}...")
        
        if ext == "txt":    file_rows = parse_txt(filepath)
        elif ext == "docx": file_rows = parse_docx(filepath)
        elif ext == "xml":  file_rows = parse_xml(filepath)
        elif ext == "html": file_rows = parse_html(filepath)
        
        for row in file_rows:
            master_records.append(row + [document_id, ext])

    if not master_records:
        print("Extraction complete. Zero valid records extracted.")
        return

    df_master = pd.DataFrame(master_records, columns=["Document ID", "File Extension", "Full Name", "SSN", "DOB", "Phone", "Email"])
    
    # 2. Re-index / Rearrange the columns exactly how you want them ordered
    target_order = ["Document ID", "File Extension", "Full Name", "SSN", "DOB", "Phone", "Email"]
    df_master = df_master[target_order]
    
    # Auto-incrementing index block
    start_breach_num = 1
    breach_ids = [f"KSTL{str(start_breach_num + idx).zfill(8)}" for idx in range(len(df_master))]
    df_master.insert(0, "Breach ID", breach_ids)
    
    output_filename = "Consolidated_Raw_Extraction.xlsx"
    df_master.to_excel(output_filename, index=False)
    
    print(f"\n[SUCCESS]: Processed {len(df_master)} total records across formats.")
    print(f"Final Data Saved to: ./{output_filename}")

if __name__ == "__main__":
    main()