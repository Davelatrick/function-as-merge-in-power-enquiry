import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import pandas as pd
import os
import traceback
import sys
import warnings

class BatchMarksheetMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Batch Marksheet Merge Tool")
        self.root.geometry("900x700")
        
        # Suppress warnings to avoid cluttering the UI
        warnings.filterwarnings("ignore")
        
        # Create a list to hold all file pairs and their configurations
        self.file_pairs = []
        
        # Create the GUI
        self.create_widgets()
    
    def create_widgets(self):
        # Create main frames
        self.control_frame = ttk.Frame(self.root, padding=10)
        self.control_frame.pack(fill='x', padx=10, pady=5)
        
        self.pairs_frame = ttk.LabelFrame(self.root, text="File Pairs", padding=10)
        self.pairs_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create a canvas with scrollbar for the pairs
        self.canvas = tk.Canvas(self.pairs_frame)
        self.scrollbar = ttk.Scrollbar(self.pairs_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Control buttons
        ttk.Button(self.control_frame, text="Add File Pair", command=self.add_file_pair).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.control_frame, text="Process All Pairs", command=self.process_all_pairs).pack(side=tk.LEFT, padx=5)
        
        # Results area
        self.results_frame = ttk.LabelFrame(self.root, text="Results", padding=10)
        self.results_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.results_text = scrolledtext.ScrolledText(self.results_frame, height=10)
        self.results_text.pack(fill='both', expand=True)
        
        # Add the first pair by default
        self.add_file_pair()
    
    def add_file_pair(self):
        pair_id = len(self.file_pairs)
        pair_frame = ttk.LabelFrame(self.scrollable_frame, text=f"Pair #{pair_id+1}", padding=10)
        pair_frame.pack(fill='x', expand=True, padx=5, pady=5, anchor='nw')
        
        # Create a dictionary to store this pair's configuration
        pair_config = {
            'id': pair_id,
            'frame': pair_frame,
            'file1_path': tk.StringVar(),
            'file2_path': tk.StringVar(),
            'output_file_path': tk.StringVar(),
            'file1_key_column': tk.StringVar(),
            'file2_key_column': tk.StringVar(),
            'file1_df': None,
            'file2_df': None,
            'status': 'Not processed',
            'dropdown_widgets': {},
            'selected_columns': [],  # List to store selected output columns
            'column_vars': [],       # List to store checkbox variables
        }
        
        # File 1 selection
        file1_frame = ttk.Frame(pair_frame)
        file1_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(file1_frame, text="File 1:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(file1_frame, textvariable=pair_config['file1_path'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(file1_frame, text="Browse...", 
                   command=lambda: self.browse_file(pair_config['file1_path'])).pack(side=tk.LEFT, padx=5)
        
        # File 2 selection
        file2_frame = ttk.Frame(pair_frame)
        file2_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(file2_frame, text="File 2:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(file2_frame, textvariable=pair_config['file2_path'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(file2_frame, text="Browse...", 
                   command=lambda: self.browse_file(pair_config['file2_path'])).pack(side=tk.LEFT, padx=5)
        
        # Load files button
        ttk.Button(pair_frame, text="Load Files", 
                  command=lambda p=pair_config: self.load_files(p)).pack(anchor='center', pady=5)
        
        # Column selection (initially hidden, shown after loading files)
        columns_frame = ttk.LabelFrame(pair_frame, text="Key Column Selection", padding=5)
        pair_config['columns_frame'] = columns_frame
        
        # File 1 key column
        file1_cols_frame = ttk.Frame(columns_frame)
        file1_cols_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(file1_cols_frame, text="File 1 Key Column:").pack(side=tk.LEFT, padx=5)
        file1_key_combo = ttk.Combobox(file1_cols_frame, textvariable=pair_config['file1_key_column'], state='readonly', width=20)
        file1_key_combo.pack(side=tk.LEFT, padx=5)
        pair_config['dropdown_widgets']['file1_key'] = file1_key_combo
        
        # File 2 key column
        file2_cols_frame = ttk.Frame(columns_frame)
        file2_cols_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(file2_cols_frame, text="File 2 Key Column:").pack(side=tk.LEFT, padx=5)
        file2_key_combo = ttk.Combobox(file2_cols_frame, textvariable=pair_config['file2_key_column'], state='readonly', width=20)
        file2_key_combo.pack(side=tk.LEFT, padx=5)
        pair_config['dropdown_widgets']['file2_key'] = file2_key_combo
        
        # Output column selection (initially hidden, shown after loading files)
        output_columns_frame = ttk.LabelFrame(pair_frame, text="Select Output Columns (Max 10)", padding=5)
        pair_config['output_columns_frame'] = output_columns_frame
        
        # Create a frame with scrollbar for the output column checkboxes
        output_canvas = tk.Canvas(output_columns_frame, height=150)
        output_scrollbar = ttk.Scrollbar(output_columns_frame, orient="vertical", command=output_canvas.yview)
        output_scrollable_frame = ttk.Frame(output_canvas)
        
        output_scrollable_frame.bind(
            "<Configure>",
            lambda e: output_canvas.configure(scrollregion=output_canvas.bbox("all"))
        )
        
        output_canvas.create_window((0, 0), window=output_scrollable_frame, anchor="nw")
        output_canvas.configure(yscrollcommand=output_scrollbar.set)
        
        output_canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        output_scrollbar.pack(side="right", fill="y")
        
        pair_config['output_scrollable_frame'] = output_scrollable_frame
        pair_config['output_canvas'] = output_canvas
        
        # Output file selection
        output_frame = ttk.Frame(pair_frame)
        output_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(output_frame, text="Output File:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(output_frame, textvariable=pair_config['output_file_path'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(output_frame, text="Browse...", 
                   command=lambda: self.browse_save_file(pair_config['output_file_path'])).pack(side=tk.LEFT, padx=5)
        
        # Status indicator
        status_frame = ttk.Frame(pair_frame)
        status_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT, padx=5)
        status_label = ttk.Label(status_frame, text=pair_config['status'])
        status_label.pack(side=tk.LEFT, padx=5)
        pair_config['status_label'] = status_label
        
        # Add remove button
        ttk.Button(pair_frame, text="Remove Pair", 
                  command=lambda p=pair_config: self.remove_pair(p)).pack(anchor='e', pady=5)
        
        # Add this pair to our list
        self.file_pairs.append(pair_config)
        
        # Update the canvas scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def remove_pair(self, pair_config):
        if len(self.file_pairs) <= 1:
            messagebox.showinfo("Cannot Remove", "You must have at least one file pair")
            return
            
        pair_config['frame'].destroy()
        self.file_pairs.remove(pair_config)
        
        # Renumber the remaining pairs
        for i, pair in enumerate(self.file_pairs):
            pair['id'] = i
            pair['frame'].configure(text=f"Pair #{i+1}")
        
        # Update the canvas scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def browse_file(self, file_path_var):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv"), ("All files", "*.*")])
        if file_path:
            # Convert to normal string and normalize path
            file_path = os.path.normpath(str(file_path))
            file_path_var.set(file_path)
    
    def browse_save_file(self, file_path_var):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if file_path:
            # Add extension if not provided
            if not (file_path.endswith('.xlsx') or file_path.endswith('.csv')):
                file_path += '.xlsx'
            # Convert to normal string and normalize path
            file_path = os.path.normpath(str(file_path))
            file_path_var.set(file_path)
    
    def try_multiple_engines(self, file_path):
        """Try to load Excel file with multiple engines"""
        self.log_message(f"Attempting to load {file_path}")
        errors = []
        
        # Check file extension
        ext = os.path.splitext(file_path)[1].lower()
        
        # If CSV, try to load as CSV
        if ext == '.csv':
            try:
                self.log_message("Trying to load as CSV file...")
                return pd.read_csv(file_path)
            except Exception as e:
                errors.append(f"CSV engine error: {str(e)}")
        
        # Try with different engines
        engines = ['openpyxl', 'xlrd']
        for engine in engines:
            try:
                self.log_message(f"Trying Excel engine: {engine}...")
                # Try with explicit engine
                return pd.read_excel(file_path, engine=engine)
            except Exception as e:
                errors.append(f"{engine} engine error: {str(e)}")
        
        # If all of the above failed, try with default settings
        try:
            self.log_message("Trying default pandas Excel reader...")
            return pd.read_excel(file_path)
        except Exception as e:
            errors.append(f"Default engine error: {str(e)}")
            
        # If we got here, all attempts failed
        error_summary = "\n".join(errors)
        raise Exception(f"Failed to load file with any method. Errors:\n{error_summary}")
    
    def load_files(self, pair_config):
        file1_path = pair_config['file1_path'].get()
        file2_path = pair_config['file2_path'].get()
        
        if not file1_path or not file2_path:
            messagebox.showerror("Error", f"Please select both files for Pair #{pair_config['id']+1}")
            return
        
        try:
            # Check if files exist
            if not os.path.isfile(file1_path):
                raise FileNotFoundError(f"File 1 does not exist: {file1_path}")
            if not os.path.isfile(file2_path):
                raise FileNotFoundError(f"File 2 does not exist: {file2_path}")
                
            # Print file sizes for debugging
            file1_size = os.path.getsize(file1_path) / 1024 / 1024  # Size in MB
            file2_size = os.path.getsize(file2_path) / 1024 / 1024  # Size in MB
            self.log_message(f"File 1 size: {file1_size:.2f} MB")
            self.log_message(f"File 2 size: {file2_size:.2f} MB")
            
            # Use try/except for each file to provide more specific error messages
            try:
                self.log_message(f"Loading File 1: {file1_path}")
                pair_config['file1_df'] = self.try_multiple_engines(file1_path)
                self.log_message(f"File 1 loaded successfully with shape: {pair_config['file1_df'].shape}")
            except Exception as e:
                self.log_message(f"Error loading File 1: {str(e)}")
                raise Exception(f"Error loading File 1: {str(e)}")
                
            try:
                self.log_message(f"Loading File 2: {file2_path}")
                pair_config['file2_df'] = self.try_multiple_engines(file2_path)
                self.log_message(f"File 2 loaded successfully with shape: {pair_config['file2_df'].shape}")
            except Exception as e:
                self.log_message(f"Error loading File 2: {str(e)}")
                raise Exception(f"Error loading File 2: {str(e)}")
            
            # Show the column selection areas
            pair_config['columns_frame'].pack(fill='x', padx=5, pady=5)
            
            # Update dropdowns with column headers
            file1_columns = list(pair_config['file1_df'].columns)
            file2_columns = list(pair_config['file2_df'].columns)
            
            # Log column data types to identify potential issues
            self.log_message("File 1 column types:")
            for col in file1_columns:
                dtype = pair_config['file1_df'][col].dtype
                self.log_message(f"  - {col}: {dtype}")
                
            self.log_message("File 2 column types:")
            for col in file2_columns:
                dtype = pair_config['file2_df'][col].dtype
                self.log_message(f"  - {col}: {dtype}")
            
            pair_config['dropdown_widgets']['file1_key']['values'] = file1_columns
            pair_config['dropdown_widgets']['file2_key']['values'] = file2_columns
            
            # Try to auto-select columns with "name" or "id" in them
            for col in file1_columns:
                col_lower = str(col).lower()
                if any(key in col_lower for key in ['name', 'id', '姓名', '名字', '学号', '身份']):
                    pair_config['file1_key_column'].set(col)
                    break
            
            for col in file2_columns:
                col_lower = str(col).lower()
                if any(key in col_lower for key in ['name', 'id', '姓名', '名字', '学号', '身份']):
                    pair_config['file2_key_column'].set(col)
                    break
            
            # Clear any existing output column checkboxes
            for widget in pair_config['output_scrollable_frame'].winfo_children():
                widget.destroy()
            pair_config['column_vars'] = []
            pair_config['selected_columns'] = []
            
            # Setup output column selection
            # Create preview merged dataframe to show all potential columns
            file1_df_preview = pair_config['file1_df'].copy().add_suffix('_file1')
            file2_df_preview = pair_config['file2_df'].copy().add_suffix('_file2')
            
            # Get all possible column names
            all_columns = list(file1_df_preview.columns) + list(file2_df_preview.columns)
            
            # Add checkboxes for column selection
            ttk.Label(pair_config['output_scrollable_frame'], 
                    text="Select columns to include in output (max 10):").pack(anchor='w', padx=5, pady=5)
            
            def update_selection():
                # Update the selected_columns list based on checkboxes
                pair_config['selected_columns'] = [col for col, var in zip(all_columns, pair_config['column_vars']) if var.get()]
                
                # If more than 10 are selected, deselect the most recent one
                if len(pair_config['selected_columns']) > 10:
                    # Find index of last selected item
                    for i in range(len(pair_config['column_vars'])-1, -1, -1):
                        if pair_config['column_vars'][i].get():
                            pair_config['column_vars'][i].set(False)
                            break
                    # Update selected columns again
                    pair_config['selected_columns'] = [col for col, var in zip(all_columns, pair_config['column_vars']) if var.get()]
                
                # Update selection count
                selection_count_label.config(text=f"Selected: {len(pair_config['selected_columns'])}/10")
            
            for col in all_columns:
                var = tk.BooleanVar(value=False)
                pair_config['column_vars'].append(var)
                chk = ttk.Checkbutton(
                    pair_config['output_scrollable_frame'], 
                    text=col,
                    variable=var,
                    command=update_selection
                )
                chk.pack(anchor='w', padx=20, pady=2)
            
            # Add a label to show selection count
            selection_count_label = ttk.Label(pair_config['output_scrollable_frame'], text="Selected: 0/10")
            selection_count_label.pack(anchor='w', padx=5, pady=5)
            pair_config['selection_count_label'] = selection_count_label
            
            # Show the output column selection frame
            pair_config['output_columns_frame'].pack(fill='x', padx=5, pady=5)
            
            # Suggest output file name based on input files
            if not pair_config['output_file_path'].get():
                file1_basename = os.path.basename(file1_path)
                base, ext = os.path.splitext(file1_basename)
                output_name = f"merged_{base}.xlsx"
                output_dir = os.path.dirname(file1_path)
                pair_config['output_file_path'].set(os.path.join(output_dir, output_name))
            
            # Update status
            pair_config['status'] = 'Files loaded'
            pair_config['status_label'].configure(text=pair_config['status'], foreground='green')
            
            self.log_message(f"Pair #{pair_config['id']+1}: Files loaded successfully")
            
        except Exception as e:
            error_details = traceback.format_exc()
            self.log_message(f"Error loading files for Pair #{pair_config['id']+1}: {str(e)}")
            self.log_message(error_details)
            messagebox.showerror("Error", f"Error loading files for Pair #{pair_config['id']+1}: {str(e)}")
    
    def process_all_pairs(self):
        if not self.file_pairs:
            messagebox.showinfo("No Pairs", "No file pairs to process")
            return
            
        total_pairs = len(self.file_pairs)
        successful_pairs = 0
        failed_pairs = 0
        
        self.log_message(f"Starting to process {total_pairs} file pairs...")
        
        for pair in self.file_pairs:
            # Skip if files not loaded
            if pair['file1_df'] is None or pair['file2_df'] is None:
                self.log_message(f"Pair #{pair['id']+1}: Skipped - Files not loaded")
                pair['status'] = 'Skipped - Files not loaded'
                pair['status_label'].configure(text=pair['status'], foreground='orange')
                failed_pairs += 1
                continue
                
            # Get column selections
            file1_key_col = pair['file1_key_column'].get()
            file2_key_col = pair['file2_key_column'].get()
            output_path = pair['output_file_path'].get()
            selected_columns = pair['selected_columns']
            
            # Validate selections
            if not file1_key_col or not file2_key_col:
                self.log_message(f"Pair #{pair['id']+1}: Missing key column selections")
                pair['status'] = 'Failed - Missing key column selections'
                pair['status_label'].configure(text=pair['status'], foreground='red')
                failed_pairs += 1
                continue
                
            if not output_path:
                self.log_message(f"Pair #{pair['id']+1}: No output path specified")
                pair['status'] = 'Failed - No output path'
                pair['status_label'].configure(text=pair['status'], foreground='red')
                failed_pairs += 1
                continue
                
            try:
                self.log_message(f"Processing Pair #{pair['id']+1}...")
                
                # Create deep copies to avoid modifying original dataframes
                file1_df = pair['file1_df'].copy()
                file2_df = pair['file2_df'].copy()
                
                # Clean column names to avoid issues
                file1_df.columns = [str(col).strip() for col in file1_df.columns]
                file2_df.columns = [str(col).strip() for col in file2_df.columns]
                
                # Ensure key columns exist
                if file1_key_col not in file1_df.columns:
                    raise ValueError(f"Key column '{file1_key_col}' not found in File 1")
                if file2_key_col not in file2_df.columns:
                    raise ValueError(f"Key column '{file2_key_col}' not found in File 2")
                
                # Convert key columns to string to ensure proper merging
                file1_df[file1_key_col] = file1_df[file1_key_col].fillna('').astype(str)
                file2_df[file2_key_col] = file2_df[file2_key_col].fillna('').astype(str)
                
                # Add suffixes to columns to differentiate them after merge
                file1_df_renamed = file1_df.add_suffix('_file1')
                file2_df_renamed = file2_df.add_suffix('_file2')
                
                # Rename the key columns back to a common column for merging
                file1_df_renamed = file1_df_renamed.rename(columns={f"{file1_key_col}_file1": "Merge_Key"})
                file2_df_renamed = file2_df_renamed.rename(columns={f"{file2_key_col}_file2": "Merge_Key"})
                
                # Log data for debugging
                self.log_message(f"File 1 shape before merge: {file1_df_renamed.shape}")
                self.log_message(f"File 2 shape before merge: {file2_df_renamed.shape}")
                
                # Alternative merge approach using pandas merge function instead of join
                merged_df = pd.merge(
                    file1_df_renamed, 
                    file2_df_renamed,
                    on="Merge_Key",
                    how="outer"
                )
                
                self.log_message(f"Merged dataframe shape: {merged_df.shape}")
                
                # If user has selected specific columns for output
                if selected_columns:
                    # Make sure Merge_Key is always included
                    if "Merge_Key" not in selected_columns:
                        final_cols = ["Merge_Key"] + [col for col in selected_columns]
                    else:
                        final_cols = selected_columns
                        
                    # Only include columns that exist in the merged dataframe
                    final_cols = [col for col in final_cols if col in merged_df.columns]
                    
                    self.log_message(f"Using {len(final_cols)} selected columns for output")
                    final_df = merged_df[final_cols]
                else:
                    # Reorder columns to have file1 columns first, then file2
                    file1_cols = ["Merge_Key"] + [col for col in merged_df.columns if col.endswith('_file1') and col != "Merge_Key"]
                    file2_cols = [col for col in merged_df.columns if col.endswith('_file2')]
                    
                    # Final column order
                    final_cols = file1_cols + file2_cols
                    self.log_message(f"Using default column ordering with {len(final_cols)} columns")
                    final_df = merged_df[final_cols]
                
                # Create output directory if it doesn't exist
                output_dir = os.path.dirname(output_path)
                if output_dir and not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                # Normalize output path
                output_path = os.path.normpath(output_path)
                
                # Save based on file extension
                if output_path.lower().endswith('.csv'):
                    self.log_message(f"Saving as CSV: {output_path}")
                    final_df.to_csv(output_path, index=False)
                else:
                    self.log_message(f"Saving as Excel: {output_path}")
                    final_df.to_excel(output_path, index=False, engine='openpyxl')
                
                self.log_message(f"Pair #{pair['id']+1}: Successfully processed and saved to {output_path}")
                self.log_message(f"  - Records merged: {len(final_df)}")
                if selected_columns:
                    self.log_message(f"  - Selected columns: {len(final_cols)}")
                
                pair['status'] = 'Processed successfully'
                pair['status_label'].configure(text=pair['status'], foreground='green')
                successful_pairs += 1
                
            except Exception as e:
                error_details = traceback.format_exc()
                self.log_message(f"Pair #{pair['id']+1}: Failed - {str(e)}")
                self.log_message(error_details)
                
                pair['status'] = 'Failed - Processing error'
                pair['status_label'].configure(text=pair['status'], foreground='red')
                failed_pairs += 1
        
        # Show summary
        self.log_message("\nProcessing Summary:")
        self.log_message(f"Total pairs: {total_pairs}")
        self.log_message(f"Successfully processed: {successful_pairs}")
        self.log_message(f"Failed: {failed_pairs}")
        
        messagebox.showinfo("Processing Complete", 
                           f"Processing complete.\nSuccessful: {successful_pairs}\nFailed: {failed_pairs}")
    
    def log_message(self, message):
        self.results_text.configure(state='normal')
        self.results_text.insert(tk.END, str(message) + "\n")
        self.results_text.see(tk.END)
        self.results_text.configure(state='disabled')
        # Update UI
        self.root.update_idletasks()

def main():
    root = tk.Tk()
    app = BatchMarksheetMergeApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()