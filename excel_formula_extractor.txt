import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import re
from pathlib import Path
from python_calamine import load_workbook


class ExcelFormulaExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Formula Extractor")
        self.root.geometry("700x500")
        self.file_path = None
        self.sheet_names = []
        
        self.setup_ui()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="Excel File:").pack(side=tk.LEFT)
        self.file_entry = ttk.Entry(file_frame, textvariable=tk.StringVar())
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT)
        
        # Sheet selection
        sheet_frame = ttk.LabelFrame(main_frame, text="Sheet Selection", padding="5")
        sheet_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.sheet_var = tk.StringVar(value="all")
        ttk.Radiobutton(sheet_frame, text="All Sheets", variable=self.sheet_var, 
                       value="all", command=self.on_sheet_mode_change).pack(anchor=tk.W)
        ttk.Radiobutton(sheet_frame, text="Selected Sheets", variable=self.sheet_var, 
                       value="selected", command=self.on_sheet_mode_change).pack(anchor=tk.W)
        
        self.sheet_listbox = tk.Listbox(sheet_frame, selectmode=tk.EXTENDED, height=10)
        self.sheet_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Output file
        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Output JSON:").pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(output_frame, textvariable=tk.StringVar())
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_frame, text="Browse", command=self.browse_output).pack(side=tk.LEFT)
        
        # Extract button
        ttk.Button(main_frame, text="Extract Formulas", command=self.extract_formulas).pack(pady=10)
        
        # Progress
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack()
    
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file_path = path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, path)
            self.load_sheets()
    
    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", 
                                           filetypes=[("JSON files", "*.json")])
        if path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, path)
    
    def load_sheets(self):
        try:
            wb = load_workbook(self.file_path)
            self.sheet_names = wb.sheet_names
            wb.close()
            
            self.sheet_listbox.delete(0, tk.END)
            for name in self.sheet_names:
                self.sheet_listbox.insert(tk.END, name)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")
    
    def on_sheet_mode_change(self):
        pass
    
    def extract_formulas(self):
        if not self.file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first")
            return
        
        output_path = self.output_entry.get()
        if not output_path:
            messagebox.showwarning("Warning", "Please specify output JSON file")
            return
        
        try:
            self.status_label.config(text="Loading workbook...")
            self.root.update()
            
            wb = load_workbook(self.file_path, formulas=True)
            
            # Determine which sheets to process
            if self.sheet_var.get() == "all":
                sheets_to_process = self.sheet_names
            else:
                selected_indices = self.sheet_listbox.curselection()
                if not selected_indices:
                    messagebox.showwarning("Warning", "Please select at least one sheet")
                    return
                sheets_to_process = [self.sheet_names[i] for i in selected_indices]
            
            results = {}
            total_sheets = len(sheets_to_process)
            
            for idx, sheet_name in enumerate(sheets_to_process):
                self.status_label.config(text=f"Processing sheet: {sheet_name} ({idx+1}/{total_sheets})")
                self.progress['value'] = (idx / total_sheets) * 100
                self.root.update()
                
                sheet = wb[sheet_name]
                sheet_results = []
                
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                            # Extract external file references
                            external_refs = self.extract_external_refs(cell.value)
                            if external_refs:
                                sheet_results.append({
                                    "cell": cell.coordinator,
                                    "formula": cell.value,
                                    "external_files": external_refs
                                })
                
                if sheet_results:
                    results[sheet_name] = sheet_results
            
            wb.close()
            
            # Save results
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, ensure_ascii=False, indent=2)
            
            self.progress['value'] = 100
            self.status_label.config(text=f"Completed! Found {sum(len(v) for v in results.values())} cells with external references")
            messagebox.showinfo("Success", f"Results saved to:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract formulas: {e}")
    
    def extract_external_refs(self, formula):
        """Extract external file references from formula"""
        external_refs = []
        
        # Pattern for external references like [filename.xlsx]Sheet!Cell or 'path[file.xlsx]Sheet'!Cell
        patterns = [
            # [filename.xlsx]Sheet!Cell or [filename.xlsx]Sheet!A1:B2
            r'\[([^\]]+)\]',
            # 'path[file.xlsx]Sheet'!Cell
            r'\'([^\']+)\[[^\]]+\][^\']+\'!',
            # External workbook reference
            r'\[([^\]]+\.(?:xlsx|xls|xlsm|ods))\]',
        ]
        
        # Match patterns like [filename.xlsx] or 'C:\path[filename.xlsx]sheet'!
        matches = re.findall(r'[\'"]?(\[.+?\])[\'"]?!', formula)
        for match in matches:
            # Extract filename from brackets
            file_match = re.search(r'\[([^\]]+\.(?:xlsx|xls|xlsm|ods))\]', match)
            if file_match:
                external_refs.append(file_match.group(1))
            else:
                clean_match = match.strip('[]"\'')
                if clean_match:
                    external_refs.append(clean_match)
        
        # Also check for direct file path patterns
        direct_paths = re.findall(r'[\'"]([A-Za-z]:[^\'"]+\.(?:xlsx|xls|xlsm|ods))[\'"]', formula)
        external_refs.extend(direct_paths)
        
        return list(set(external_refs))


def main():
    root = tk.Tk()
    app = ExcelFormulaExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
