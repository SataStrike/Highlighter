import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from pathlib import Path

class ErrorDistributionCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Error Distribution Calculator")
        self.root.geometry("600x400")
        
        # Variables
        self.input_file = None
        self.df = None
        self.processed_df = None
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Error Distribution Calculator", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input file section
        ttk.Label(main_frame, text="Input CSV File:", font=('Arial', 10)).grid(
            row=1, column=0, sticky=tk.W, pady=(10, 5))
        
        self.input_label = ttk.Label(main_frame, text="No file selected", 
                                    relief=tk.SUNKEN, padding=5, foreground="gray")
        self.input_label.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                             padx=(0, 10))
        
        ttk.Button(main_frame, text="Browse", command=self.browse_input).grid(
            row=2, column=2, padx=(0, 10))
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="Calculate Error Distribution", 
                                     command=self.process_data, state=tk.DISABLED)
        self.process_btn.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), 
                          pady=(0, 20))
        
        # Preview section
        preview_frame = ttk.LabelFrame(main_frame, text="Preview (First 10 rows)", 
                                      padding="10")
        preview_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), 
                          pady=(10, 0))
        
        # Text widget for preview
        self.preview_text = tk.Text(preview_frame, height=10, width=70, wrap=tk.NONE)
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars for preview
        v_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, 
                                   command=self.preview_text.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, 
                                   command=self.preview_text.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        self.preview_text.config(yscrollcommand=v_scrollbar.set, 
                                xscrollcommand=h_scrollbar.set)
        
        # Export button
        self.export_btn = ttk.Button(main_frame, text="Export to Excel", 
                                    command=self.export_data, state=tk.DISABLED)
        self.export_btn.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if filename:
            self.input_file = filename
            # Show full path on hover, filename in label
            self.input_label.config(text=Path(filename).name, foreground="black")
            self.input_label.bind("<Enter>", lambda e: self.show_tooltip(e, filename))
            self.input_label.bind("<Leave>", lambda e: self.hide_tooltip())
            self.process_btn.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.DISABLED)
            self.preview_text.delete(1.0, tk.END)
    
    def show_tooltip(self, event, text):
        """Show tooltip with full path"""
        self.tooltip = tk.Toplevel()
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
        label = tk.Label(self.tooltip, text=text, background="lightyellow", 
                        relief=tk.SOLID, borderwidth=1, font=('Arial', 9))
        label.pack()
    
    def hide_tooltip(self):
        """Hide tooltip"""
        if hasattr(self, 'tooltip'):
            self.tooltip.destroy()
            
    def calculate_error_distribution(self, df):
        """Calculate error distribution percentage for each website/app"""
        # Create a copy of the dataframe
        result_df = df.copy()
        
        # Calculate total ad calls for each website
        website_totals = df.groupby('Website/App Name')['Ad Calls'].sum()
        
        # Calculate percentage for each row
        percentages = []
        for idx, row in df.iterrows():
            website = row['Website/App Name']
            ad_calls = row['Ad Calls']
            total = website_totals[website]
            
            if total > 0:
                percentage = (ad_calls / total) * 100
            else:
                percentage = 0
                
            percentages.append(f"{percentage:.2f}%")
        
        result_df['Error Distribution'] = percentages
        
        return result_df
    
    def process_data(self):
        try:
            # Start progress bar
            self.progress.start()
            
            # Read CSV file
            self.df = pd.read_csv(self.input_file)
            
            # Validate required columns
            required_columns = ['Website/App Name', 'CSM Error', 'Type', 
                              'Website Ads Txt Reason', 'Ad Calls']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            
            if missing_columns:
                raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
            
            # Calculate error distribution
            self.processed_df = self.calculate_error_distribution(self.df)
            
            # Show preview
            self.show_preview()
            
            # Enable export button
            self.export_btn.config(state=tk.NORMAL)
            
            # Stop progress bar
            self.progress.stop()
            
            messagebox.showinfo("Success", "Error distribution calculated successfully!")
            
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def show_preview(self):
        """Show preview of processed data"""
        self.preview_text.delete(1.0, tk.END)
        
        if self.processed_df is not None:
            # Convert first 10 rows to string for preview
            preview_data = self.processed_df.head(10).to_string(index=False)
            self.preview_text.insert(1.0, preview_data)
            
            # Add summary
            total_websites = self.processed_df['Website/App Name'].nunique()
            total_rows = len(self.processed_df)
            summary = f"\n\nSummary: {total_websites} unique websites, {total_rows} total rows"
            self.preview_text.insert(tk.END, summary)
    
    def export_data(self):
        if self.processed_df is None:
            messagebox.showwarning("Warning", "No data to export!")
            return
        
        # Ask user to select the output directory and filename
        filename = filedialog.asksaveasfilename(
            title="Choose location and name for Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="error_distribution_results.xlsx"
        )
        
        if filename:
            try:
                # Show the selected path to user
                self.root.config(cursor="wait")
                self.export_btn.config(text="Exporting...", state=tk.DISABLED)
                self.root.update()
                
                # Export to Excel with formatting
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    self.processed_df.to_excel(writer, sheet_name='Error Distribution', 
                                             index=False)
                    
                    # Get the workbook and worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['Error Distribution']
                    
                    # Auto-adjust column widths
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                    
                    # Add filters
                    worksheet.auto_filter.ref = worksheet.dimensions
                
                # Reset cursor and button
                self.root.config(cursor="")
                self.export_btn.config(text="Export to Excel", state=tk.NORMAL)
                
                # Show success message with full path
                messagebox.showinfo("Success", 
                    f"File exported successfully!\n\nSaved to:\n{filename}")
                
            except Exception as e:
                self.root.config(cursor="")
                self.export_btn.config(text="Export to Excel", state=tk.NORMAL)
                messagebox.showerror("Error", f"Failed to export file: {str(e)}")

def main():
    root = tk.Tk()
    app = ErrorDistributionCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()