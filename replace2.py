import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class ReplacementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Replacement Tool")
        self.root.geometry("600x400")
        
        # File Path Selection
        self.frame_file = ttk.LabelFrame(root, text="Excel File Selection", padding="5")
        self.frame_file.pack(fill="x", padx=5, pady=5)
        
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(self.frame_file, textvariable=self.path_var, width=50)
        self.path_entry.pack(side="left", padx=5)
        
        self.browse_btn = ttk.Button(self.frame_file, text="Browse", command=self.browse_file)
        self.browse_btn.pack(side="left", padx=5)
        
        # Sheet Selection
        self.frame_sheet = ttk.LabelFrame(root, text="Sheet Selection", padding="5")
        self.frame_sheet.pack(fill="x", padx=5, pady=5)
        
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self.frame_sheet, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.pack(fill="x", padx=5)
        
        # Range Input
        self.frame_range = ttk.LabelFrame(root, text="Range Selection (e.g., E2:F26)", padding="5")
        self.frame_range.pack(fill="x", padx=5, pady=5)
        
        self.range_var = tk.StringVar()
        self.range_entry = ttk.Entry(self.frame_range, textvariable=self.range_var)
        self.range_entry.pack(fill="x", padx=5)
        
        # Replacement Rules
        self.frame_rules = ttk.LabelFrame(root, text="Replacement Rules (format: find1,replace1,find2,replace2,...)", padding="5")
        self.frame_rules.pack(fill="x", padx=5, pady=5)
        
        self.rules_text = tk.Text(self.frame_rules, height=5)
        self.rules_text.pack(fill="x", padx=5)
        
        # Process Button
        self.process_btn = ttk.Button(root, text="Process Replacements", command=self.process_replacements)
        self.process_btn.pack(pady=20)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.path_var.set(filename)
            self.update_sheet_list(filename)

    def update_sheet_list(self, filename):
        try:
            excel_file = pd.ExcelFile(filename)
            sheet_names = excel_file.sheet_names
            self.sheet_combo['values'] = sheet_names
            if sheet_names:
                self.sheet_combo.set(sheet_names[0])
        except Exception as e:
            messagebox.showerror("Error", f"Error reading Excel file: {str(e)}")

    def safe_replace(self, value, replace_dict):
        if pd.isna(value):
            return value
        try:
            str_value = str(value)
            for find_str, replace_str in replace_dict.items():
                str_value = str_value.replace(find_str, replace_str)
            return str_value
        except:
            return value

    def process_replacements(self):
        try:
            # Get input values
            file_path = self.path_var.get()
            sheet_name = self.sheet_var.get()
            cell_range = self.range_var.get()
            rules_text = self.rules_text.get("1.0", "end-1c")

            if not all([file_path, sheet_name, cell_range, rules_text]):
                messagebox.showerror("Error", "Please fill in all fields")
                return

            # Parse replacement rules
            rules = rules_text.strip().split(',')
            if len(rules) % 2 != 0:
                messagebox.showerror("Error", "Invalid replacement rules format")
                return

            # Create replacement dictionary
            replace_dict = {rules[i]: rules[i+1] for i in range(0, len(rules), 2)}

            # Read all sheets from the Excel file
            excel_file = pd.ExcelFile(file_path)
            all_sheets = {}
            for sheet in excel_file.sheet_names:
                all_sheets[sheet] = pd.read_excel(file_path, sheet_name=sheet)

            # Modify only the selected sheet
            df = all_sheets[sheet_name]
            
            # Apply replacements only to string columns
            for column in df.columns:
                df[column] = df[column].apply(lambda x: self.safe_replace(x, replace_dict))

            all_sheets[sheet_name] = df

            # Save all sheets back to the original file
            with pd.ExcelWriter(file_path, mode='w') as writer:
                for sheet, data in all_sheets.items():
                    data.to_excel(writer, sheet_name=sheet, index=False)

            messagebox.showinfo("Success", "Replacements completed and saved to the original file!")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReplacementApp(root)
    root.mainloop()