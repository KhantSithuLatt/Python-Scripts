import os
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class DataConsolidatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Pipeline Consolidation Engine")
        
        # --- DYNAMIC MONITOR CENTERING LOGIC ---
        window_width = 1100
        window_height = 650
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width / 2) - (window_width / 2))
        center_y = int((screen_height / 2) - (window_height / 2))
        
        self.root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        self.root.resizable(False, False)
        
        # Paths variables
        self.source_dir = ""
        self.output_dir = ""
        
        # Build UI layout
        self.build_ui()
        
    def build_ui(self):
        # 1. Top Header Banner
        header_frame = ttk.Frame(self.root, padding=15)
        header_frame.pack(fill="x", side="top")
        header_lbl = ttk.Label(
            header_frame, 
            text="⚙️ AUTOMATED DATA PIPELINE CONSOLIDATION MODULE", 
            font=("TkDefaultFont", 14, "bold")
        )
        header_lbl.pack(anchor="w")
        
        # 2. Workspace Control Panel
        workspace = ttk.Frame(self.root, padding=20)
        workspace.pack(fill="x", side="top")
        workspace.columnconfigure(1, weight=1)
        
        # Row 1: Source Folder Selection
        lbl_source = ttk.Label(workspace, text="Source Data Directory (Verified Sheets):")
        lbl_source.grid(row=0, column=0, sticky="w", pady=8, padx=5)
        
        self.lbl_src_path = ttk.Label(workspace, text=" No folder selected...", width=80, anchor="w", relief="sunken", padding=4)
        self.lbl_src_path.grid(row=0, column=1, sticky="ew", padx=10, pady=8)
        
        btn_browse_src = ttk.Button(workspace, text="Select Source Folder", command=self.select_source_folder)
        btn_browse_src.grid(row=0, column=2, padx=5, pady=8)
        
        # Row 2: Output Destination Selection
        lbl_dest = ttk.Label(workspace, text="Output Master Save Directory:")
        lbl_dest.grid(row=1, column=0, sticky="w", pady=8, padx=5)
        
        self.lbl_dest_path = ttk.Label(workspace, text=" No save path selected...", width=80, anchor="w", relief="sunken", padding=4)
        self.lbl_dest_path.grid(row=1, column=1, sticky="ew", padx=10, pady=8)
        
        btn_browse_dest = ttk.Button(workspace, text="Select Save Directory", command=self.select_output_directory)
        btn_browse_dest.grid(row=1, column=2, padx=5, pady=8)
        
        # Row 3: Output File Name Entry Row
        lbl_name = ttk.Label(workspace, text="Consolidated File Name:")
        lbl_name.grid(row=2, column=0, sticky="w", pady=8, padx=5)
        
        self.ent_filename = ttk.Entry(workspace, width=80)
        self.ent_filename.grid(row=2, column=1, sticky="ew", padx=10, pady=8)
        self.ent_filename.insert(0, "Consolidated_Master_Breach_Dataset") # Default name proposal
        
        lbl_ext = ttk.Label(workspace, text=".xlsx", font=("TkDefaultFont", 10, "bold"))
        lbl_ext.grid(row=2, column=2, sticky="w", padx=5, pady=8)
        
        # Row 4: Action Button Run Trigger
        self.btn_run = ttk.Button(workspace, text="⚡ RUN MASTER DATA CONSOLIDATION", command=self.execute_consolidation, state="disabled")
        self.btn_run.grid(row=3, column=0, columnspan=3, pady=20, ipady=6)
        
        # 3. Log Console Terminal
        console_frame = ttk.Frame(self.root, padding=15)
        console_frame.pack(fill="both", expand=True, side="top")
        
        console_lbl = ttk.Label(console_frame, text="Execution Log Stream Panel:")
        console_lbl.pack(anchor="w", pady=5)
        
        self.log_box = tk.Text(console_frame, background="#1e1e1e", foreground="#d4d4d4", font=("Monospace", 9))
        scrollbar = ttk.Scrollbar(console_frame, orient="vertical", command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.log_box.pack(side="left", fill="both", expand=True)

    def log(self, message):
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.root.update_idletasks()

    def select_source_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing .xlsx Data Sheets")
        if folder:
            self.source_dir = folder
            self.lbl_src_path.config(text=f"  {folder}")
            self.check_ready_state()

    def select_output_directory(self):
        folder = filedialog.askdirectory(title="Select Destination Directory to Save Output File")
        if folder:
            self.output_dir = folder
            self.lbl_dest_path.config(text=f"  {folder}")
            self.check_ready_state()

    def check_ready_state(self):
        if self.source_dir and self.output_dir:
            self.btn_run.state(["!disabled"])

    def execute_consolidation(self):
        # Clear log console screen
        self.log_box.delete("1.0", tk.END)
        
        output_name_raw = self.ent_filename.get().strip()
        if not output_name_raw:
            messagebox.showerror("Validation Error", "Please provide a valid output file name.")
            return
            
        # Enforce file extension standard safely
        if not output_name_raw.endswith(".xlsx"):
            output_name_raw += ".xlsx"
            
        final_output_path = os.path.join(self.output_dir, output_name_raw)
        
        self.log("=" * 110)
        self.log(" STARTING WORKSPACE RECORD CONSOLIDATION MERGE LOOP")
        self.log("=" * 110)
        self.log(f"[DIR INFO] Source Path : {self.source_dir}")
        self.log(f"[DIR INFO] Save Target : {final_output_path}")
        self.log("-" * 110)
        
        # Pull strict .xlsx format files only from the folder as specified
        excel_files = sorted([
            f for f in os.listdir(self.source_dir) 
            if f.endswith('.xlsx') and not f.startswith('~$')  # Filter out hidden windows lock files
        ])
        
        if not excel_files:
            self.log("[-] Critical Abort: No valid .xlsx spreadsheets located inside source path directory.")
            messagebox.showwarning("No Files Found", "No target .xlsx files located in the chosen folder.")
            return
            
        master_data_records = []
        
        # Isolate the exact template header array blocks from the very first matching workbook
        try:
            first_file_path = os.path.join(self.source_dir, excel_files[0])
            header_sample_df = pd.read_excel(first_file_path, header=None, nrows=2)
            row1_headers = list(header_sample_df.iloc[0])
            row2_headers = list(header_sample_df.iloc[1])
            
            # Map Column A header configurations dynamically over the baseline arrays
            final_row1 = ["File Metadata"] + row1_headers
            final_row2 = ["Source File Name"] + row2_headers
            
            self.log(f"[HEADER PARSE] Extracted baseline structural blueprint headers from: {excel_files[0]}")
        except Exception as e:
            self.log(f"[CRITICAL ERROR] Master header template mapping failed: {str(e)}")
            return
            
        self.log(f"[PROCESSING] Queued {len(excel_files)} target workbooks for deep extraction...")
        self.log("-" * 110)
        
        # Begin file injection loop
        for file in excel_files:
            file_full_path = os.path.join(self.source_dir, file)
            try:
                # Read structural cells skipping row 1 and 2 headers directly
                df_data = pd.read_excel(file_full_path, header=None, skiprows=2)
                
                # Strip file extensions off the path string completely for a clean Column A entry
                clean_file_identifier = os.path.splitext(file)[0]
                
                # Prepend current workbook name value into column index position 0 for all rows
                for _, file_row in df_data.iterrows():
                    master_data_records.append([clean_file_identifier] + list(file_row))
                    
                self.log(f" ✔ Consolidated {len(df_data):4d} records from -> {file}")
            except Exception as e:
                self.log(f" ❌ Failed to read data elements from {file}. Error message: {str(e)}")
                
        # Build the unified output dataframe matrix
        try:
            self.log("-" * 110)
            self.log("[WRITING] Packaging master arrays and rebuilding multi-row header structure...")
            
            combined_final_matrix = [final_row1, final_row2] + master_data_records
            df_final_master = pd.DataFrame(combined_final_matrix)
            
            # Save workbook without indices or auto pandas indexing labels
            df_final_master.to_excel(final_output_path, header=False, index=False)
            
            self.log("=" * 110)
            self.log(f"🎉 PIPELINE CONSOLIDATION COMPLETED SUCCESSFULLY!")
            self.log(f" Total Files Proccessed : {len(excel_files)}")
            self.log(f" Total Rows Consolidated : {len(master_data_records)}")
            self.log(f" Output Location Path   : {final_output_path}")
            self.log("=" * 110)
            
            messagebox.showinfo("Success", f"Data consolidation complete!\nMerged {len(excel_files)} files into master spreadsheet.")
        except Exception as e:
            self.log(f"[CRITICAL ERROR] Failed creating master compilation document: {str(e)}")
            messagebox.showerror("Write Error", f"Could not write master output file:\n{str(e)}")

if __name__ == "__main__":
    main_window = tk.Tk()
    app = DataConsolidatorApp(main_window)
    main_window.mainloop()