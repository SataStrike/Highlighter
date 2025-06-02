import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import re
import os

class SupplyChainValidator:
    def __init__(self, root):
        self.root = root
        self.root.title("Supply Chain Validation Tool")
        self.root.geometry("700x500")
        
        # File paths
        self.supply_chain_file = tk.StringVar()
        self.lines_referential_file = tk.StringVar()
        self.output_file = tk.StringVar()
        
        # Create UI components
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        # Supply Chain File (XLSX)
        ttk.Label(file_frame, text="Supply Chain File (XLSX):").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.supply_chain_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_supply_chain).grid(row=0, column=2, padx=5, pady=5)
        
        # Lines Referential File (CSV)
        ttk.Label(file_frame, text="Lines Referential File (CSV):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.lines_referential_file, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_lines_referential).grid(row=1, column=2, padx=5, pady=5)
        
        # Output File (CSV)
        ttk.Label(file_frame, text="Output File (CSV):").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_file, width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output).grid(row=2, column=2, padx=5, pady=5)
        
        # Processing section
        process_frame = ttk.Frame(main_frame, padding="10")
        process_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(process_frame, text="Process Files", command=self.process_files).pack(pady=10)
        
        # Status section
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.status_text = tk.Text(status_frame, height=10, wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)

    def browse_supply_chain(self):
        filename = filedialog.askopenfilename(
            title="Select Supply Chain Validation Report",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.supply_chain_file.set(filename)
            self.log_status(f"Supply Chain file selected: {filename}")

    def browse_lines_referential(self):
        filename = filedialog.askopenfilename(
            title="Select Lines Referential File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.lines_referential_file.set(filename)
            self.log_status(f"Lines Referential file selected: {filename}")

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save Output File",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
            self.log_status(f"Output file set: {filename}")

    def log_status(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def parse_missing_lines(self, text):
        """Parse the missing lines from the cell content in File 1 Column C"""
        if pd.isna(text) or text == "":
            return []
            
        # List to store parsed lines
        lines = []
        
        # Handle web entries specifically - but the actual missing lines may contain valid entries
        if isinstance(text, str) and text.strip().lower() == "web":
            return []
            
        if isinstance(text, str):
            # First split by common line separators
            raw_lines = re.split(r'\r\n|\n|\\n', text)
            if len(raw_lines) == 1 and ',' in text:
                # If it's a single line with commas, it might contain multiple entries
                # Use a regex to extract all entries instead of splitting (avoids look-behind)
                entry_pattern = r'[^,]+,\s*[^,]+,\s*(?:RESELLER|DIRECT)(?:,\s*[a-zA-Z0-9-]+)?'
                raw_lines = re.findall(entry_pattern, text, re.IGNORECASE)
            
            for raw_line in raw_lines:
                if not raw_line.strip():
                    continue
                    
                # If the line contains multiple entries (bidder, ID, type pattern)
                # Extract each entry as separate line
                bidder_patterns = re.findall(r'([^,]+)\s*,\s*([^,]+)\s*,\s*(RESELLER|DIRECT)(?:\s*,\s*([a-zA-Z0-9-]+))?', raw_line, re.IGNORECASE)
                
                if bidder_patterns:
                    for pattern in bidder_patterns:
                        # Reconstruct the line from the matched groups
                        bidder, seller_id, seller_type, certificate = pattern[0], pattern[1], pattern[2], pattern[3] if len(pattern) > 3 else ""
                        reconstructed = f"{bidder.strip()},{seller_id.strip()},{seller_type.strip().upper()}"
                        if certificate:
                            reconstructed += f",{certificate.strip()}"
                        lines.append(reconstructed)
                else:
                    # Try a simpler pattern match
                    simple_patterns = re.findall(r'([^,]+)\s*,\s*([^,]+)\s*,\s*(RESELLER|DIRECT)', raw_line, re.IGNORECASE)
                    if simple_patterns:
                        for pattern in simple_patterns:
                            bidder, seller_id, seller_type = pattern[0], pattern[1], pattern[2]
                            reconstructed = f"{bidder.strip()},{seller_id.strip()},{seller_type.strip().upper()}"
                            lines.append(reconstructed)
                    # If still no match but commas exist, treat as a single line
                    elif ',' in raw_line and ('RESELLER' in raw_line.upper() or 'DIRECT' in raw_line.upper()):
                        parts = raw_line.split(',')
                        if len(parts) >= 3:  # At minimum: bidder, ID, type
                            cleaned_parts = [part.strip() for part in parts]
                            lines.append(','.join(cleaned_parts))
            
            # Only log non-web entries that failed to parse and contain commas (likely actual line data)
            if not lines and ',' in text and not text.strip().lower() == "web":
                self.log_status(f"Warning: Could not parse any lines from: {text[:50]}...")
                
        return lines

    def normalize_line(self, line):
        """Normalize a line to standard format for better matching"""
        if not line or not isinstance(line, str):
            return ""
        parts = line.split(',')
        if len(parts) < 3:
            return line.lower().strip()
        bidder = parts[0].strip().lower()
        seller_id = parts[1].strip()
        seller_type = parts[2].strip().upper()
        normalized = f"{bidder},{seller_id},{seller_type}"
        if len(parts) > 3 and parts[3].strip():
            normalized += f",{parts[3].strip()}"
        return normalized

    def process_files(self):
        try:
            # Validate inputs
            if not self.supply_chain_file.get() or not self.lines_referential_file.get() or not self.output_file.get():
                messagebox.showerror("Error", "Please select all required files")
                return
                
            self.log_status("Starting processing...")
            self.progress['value'] = 10
            
            # Load the files
            try:
                # File 1 is XLSX (read only first sheet)
                supply_chain_path = self.supply_chain_file.get()
                supply_chain_df = pd.read_excel(supply_chain_path, sheet_name=0, dtype=str)  # Only read first sheet, force all columns as str
                self.log_status(f"Supply Chain Excel (first sheet) loaded with {len(supply_chain_df)} rows")
                
                # Log the column structure to help with debugging
                self.log_status(f"Supply Chain columns: {list(supply_chain_df.columns)}")
                if len(supply_chain_df) > 0:
                    self.log_status(f"First row sample: {supply_chain_df.iloc[0].to_dict()}")
                
                # File 2 can be CSV or Excel
                lines_ref_path = self.lines_referential_file.get()
                if lines_ref_path.endswith('.csv'):
                    lines_ref_df = pd.read_csv(lines_ref_path, dtype=str)
                else:
                    lines_ref_df = pd.read_excel(lines_ref_path, dtype=str)
                self.log_status(f"Lines Referential file loaded with {len(lines_ref_df)} rows")
                
                # Log the column structure to help with debugging
                self.log_status(f"Lines Referential columns: {list(lines_ref_df.columns)}")
                if len(lines_ref_df) > 0:
                    self.log_status(f"First row sample: {lines_ref_df.iloc[0].to_dict()}")
                    
                self.log_status("Files loaded successfully")
            except Exception as e:
                self.log_status(f"Error loading files: {str(e)}")
                messagebox.showerror("Error", f"Failed to load files: {str(e)}")
                return
                
            self.progress['value'] = 30
                
            # Prepare the referential dictionary for faster lookups
            # Key: normalized line text, Value: (category, status)
            lines_ref_dict = {}
            bidder_category_dict = {}
            for _, row in lines_ref_df.iterrows():
                line_text = str(row.iloc[0]).strip()  # Column A: Line
                category_raw = str(row.iloc[1]).strip().lower()   # Column B: Line category
                status = str(row.iloc[2]).strip()     # Column C: Status
                # Map category
                if category_raw in ["main", "master"]:
                    category = "Primary"
                elif category_raw == "secondary":
                    category = "Secondary"
                else:
                    category = category_raw.capitalize()
                normalized_line = self.normalize_line(line_text)
                lines_ref_dict[normalized_line] = (category, status)
                # Also allow lookup by bidder for fallback
                if ',' in normalized_line:
                    bidder = normalized_line.split(',')[0].strip().lower()
                    if bidder not in bidder_category_dict:
                        bidder_category_dict[bidder] = []
                    bidder_category_dict[bidder].append((normalized_line, category, status))
            self.log_status(f"Processed {len(lines_ref_dict)} reference lines (normalized)")
            print(f"[DEBUG] First 5 normalized reference lines:")
            for i, k in enumerate(list(lines_ref_dict.keys())[:5]):
                print(f"[DEBUG] Ref {i+1}: {k}")
            self.progress['value'] = 50
            
            # Process supply chain validation report
            results = []
            
                        # Each row is treated as a unique (Name, Domain) pair
            for idx, row in supply_chain_df.iterrows():
                # Debug: show the unique pair being processed
                print(f"[UNIQUE PAIR] Processing row {idx}: Name={str(row.iloc[3]).strip()}, Domain={str(row.iloc[4]).strip()}")

                try:
                    status = str(row.iloc[0]).strip() # Column A: Status
                    Monthly_adcalls = str(row.iloc[1]).strip()  # Column B: Monthly adcalls
                    Platform = str(row.iloc[2]).strip()    # Column C: Platform
                    name = str(row.iloc[3]).strip()    # Column D: Name
                    domain = str(row.iloc[4]).strip()  # Column E: Domain
                    Id = str(row.iloc[5]).strip()  # Column F: Id
                    Bundle = str(row.iloc[6]).strip()  # Column G: Bundle
                    Created_at = str(row.iloc[7]).strip()  # Column H: Created at
                    Live_at = str(row.iloc[8]).strip()  # Column I: Live at
                    Ads_txt_status = str(row.iloc[9]).strip()  # Column J: Ads.txt status
                    Sellers_json_status = str(row.iloc[10]).strip()  # Column K: Sellers.json status
                    missing_lines_text = row.iloc[11]  # Column L: Missing ads.txt lines
                    print(f"[DEBUG] Row {idx}: Domain={domain}, Name={name}, Status={status}")

                    missing_lines = self.parse_missing_lines(missing_lines_text)
                    if idx == 0:
                        print(f"[DEBUG] Sample missing_lines for first row: {missing_lines}")

                    # Count missing lines by category and collect the lines
                    primary_missing = 0
                    secondary_missing = 0
                    primary_lines = []
                    secondary_lines = []

                    for i, line in enumerate(missing_lines):
                        normalized_line = self.normalize_line(line)
                        found = False
                        match_type = None
                        # Try exact normalized match
                        if normalized_line in lines_ref_dict:
                            category, _ = lines_ref_dict[normalized_line]
                            if category == "Primary":
                                primary_missing += 1
                                primary_lines.append(line)
                            elif category == "Secondary":
                                secondary_missing += 1
                                secondary_lines.append(line)
                            found = True
                            match_type = "exact"
                        # Fallback: try by bidder
                        if not found and ',' in normalized_line:
                            bidder = normalized_line.split(',')[0].strip().lower()
                            if bidder in bidder_category_dict:
                                for _norm, category, _ in bidder_category_dict[bidder]:
                                    if category == "Primary":
                                        primary_missing += 1
                                        primary_lines.append(line)
                                        found = True
                                        match_type = "bidder"
                                        break
                                    elif category == "Secondary":
                                        secondary_missing += 1
                                        secondary_lines.append(line)
                                        found = True
                                        match_type = "bidder"
                                        break
                        if idx < 3:  # Only print debug for first 3 rows
                            print(f"[DEBUG]   Missing line: '{line}' | Normalized: '{normalized_line}' | Found: {found} | MatchType: {match_type}")
                    if (primary_missing + secondary_missing) == 0 and len(missing_lines) > 0:
                        print(f"[DEBUG][WARNING] No missing lines matched for Domain={domain}, Name={name}, first missing line: {missing_lines[0] if missing_lines else ''}")
                    results.append({
                        'Domain': domain,
                        'Name': name,
                        'Number of missing Primary lines': primary_missing,
                        'Number of missing Secondary lines': secondary_missing,
                        'Status': status
                    })
                except Exception as e:
                    print(f"[DEBUG][ERROR] Row {idx} failed: {str(e)}")
                    self.log_status(f"Error processing row {idx}: {str(e)}")

            self.progress['value'] = 80

            # Create output DataFrame and save to CSV
            print(f"[DEBUG] Number of results: {len(results)}")
            output_df = pd.DataFrame(results)
            print(f"[DEBUG] Output DataFrame shape: {output_df.shape}")
            print(f"[DEBUG] Output DataFrame columns: {output_df.columns}")
            # Ensure columns are in the correct order
            output_df = output_df[['Domain', 'Name', 'Number of missing Primary lines', 'Number of missing Secondary lines', 'Status']]
            # Make sure the output path has csv extension
            output_path = self.output_file.get()
            if not output_path.lower().endswith('.csv'):
                output_path = os.path.splitext(output_path)[0] + '.csv'
                self.output_file.set(output_path)
            output_df.to_csv(output_path, index=False)
            self.log_status(f"Output saved as CSV file: {output_path}")

            self.progress['value'] = 100
            self.log_status(f"Processing complete. Output saved to {self.output_file.get()}")
            messagebox.showinfo("Success", "Processing completed successfully!")

        except Exception as e:
            self.log_status(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = SupplyChainValidator(root)
    root.mainloop()