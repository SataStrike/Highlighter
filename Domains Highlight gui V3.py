"""
Domains Highlight GUI V3 - A modular approach to the Domains Highlighter tool
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import os
import pandas as pd
import threading

# Import our modules
from domains_highlighter import calculate_and_save_differences, percentage_to_decimal
from supply_chain_validator import process_supply_chain_files
from error_distribution import process_error_distribution
from excel_helper import write_to_excel_with_two_sheets

# Import the Bidrate improvement mail maker functionality
import sys
import os

# Import the Revenue Target Calculator functionality
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# We'll import the RevenueTargetCalculator class in the setup method to handle
# potential import errors gracefully

# Create a class that wraps the Bidrate improvement calculator for embedding in the GUI
class BidRateImprover(ttk.Frame):
    """Frame for the Bid Rate Improvement calculator"""
    
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        
        # Create the calculator UI
        self.calculator = AdRevenueCalculator(self)
        
        # Pack the calculator frame
        self.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Import the AdRevenueCalculator class directly from the script
class AdRevenueCalculator(ttk.Frame):
    """Ad Revenue Calculator embedded as a frame"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(fill=tk.BOTH, expand=True)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        self.create_widgets()
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Ad Revenue Calculator", 
                              font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Input fields
        row = 1
        
        # Monthly Ad Calls
        ttk.Label(main_frame, text="Monthly Ad Calls:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.ad_calls_var = tk.StringVar(value="30000000")
        ad_calls_entry = ttk.Entry(main_frame, textvariable=self.ad_calls_var, width=20)
        ad_calls_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Current Bid Rate
        ttk.Label(main_frame, text="Current Bid Rate (%):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.current_bid_var = tk.StringVar(value="2")
        current_bid_entry = ttk.Entry(main_frame, textvariable=self.current_bid_var, width=20)
        current_bid_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Target Bid Rate
        ttk.Label(main_frame, text="Target Bid Rate (%):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.target_bid_var = tk.StringVar(value="20")
        target_bid_entry = ttk.Entry(main_frame, textvariable=self.target_bid_var, width=20)
        target_bid_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Revenue per Billion Calls
        ttk.Label(main_frame, text="Revenue per Billion Calls ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.revenue_per_billion_var = tk.StringVar(value="16000")
        revenue_entry = ttk.Entry(main_frame, textvariable=self.revenue_per_billion_var, width=20)
        revenue_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Calculate button
        calculate_btn = ttk.Button(main_frame, text="Calculate Revenue Impact", 
                                command=self.calculate_revenue)
        calculate_btn.grid(row=row, column=0, columnspan=2, pady=20)
        row += 1
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="15")
        results_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        results_frame.columnconfigure(0, weight=1)
        row += 1
        
        # Results text widget with scrollbar
        text_frame = ttk.Frame(results_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.results_text = tk.Text(text_frame, height=25, width=80, wrap=tk.WORD, font=('Arial', 10))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights for results
        main_frame.rowconfigure(row - 1, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Clear button
        clear_btn = ttk.Button(main_frame, text="Clear Results", 
                            command=self.clear_results)
        clear_btn.grid(row=row, column=0, columnspan=2, pady=10)
    
    def calculate_revenue(self):
        try:
            # Get input values
            monthly_ad_calls = float(self.ad_calls_var.get().replace(',', ''))
            current_bid_rate = float(self.current_bid_var.get()) / 100
            target_bid_rate = float(self.target_bid_var.get()) / 100
            revenue_per_billion = float(self.revenue_per_billion_var.get().replace(',', ''))
            
            # Validate inputs
            if monthly_ad_calls <= 0 or revenue_per_billion <= 0:
                raise ValueError("Ad calls and revenue must be positive numbers")
            if current_bid_rate < 0 or target_bid_rate < 0 or current_bid_rate > 1 or target_bid_rate > 1:
                raise ValueError("Bid rates must be between 0% and 100%")
            
            # Calculate current performance
            current_biddable = monthly_ad_calls * current_bid_rate
            current_billion_equiv = current_biddable / 1_000_000_000
            current_monthly_revenue = current_billion_equiv * revenue_per_billion
            
            # Calculate target performance
            target_biddable = monthly_ad_calls * target_bid_rate
            target_billion_equiv = target_biddable / 1_000_000_000
            target_monthly_revenue = target_billion_equiv * revenue_per_billion
            
            # Calculate increases
            additional_revenue = target_monthly_revenue - current_monthly_revenue
            if current_monthly_revenue > 0:
                revenue_multiplier = target_monthly_revenue / current_monthly_revenue
                percentage_increase = (revenue_multiplier - 1) * 100
            else:
                revenue_multiplier = float('inf') if target_monthly_revenue > 0 else 1
                percentage_increase = float('inf') if target_monthly_revenue > 0 else 0
            
            # Format results with your custom text
            custom_summary = f"""
REVENUE OPTIMIZATION SUMMARY
============================

Implementing these changes might help increasing the bid rate.
Targeting {target_bid_rate*100:.1f}% from its current position of {current_bid_rate*100:.1f}%.
At the same level of ad calls it would represent an additional monthly revenue of ${additional_revenue:,.2f} or ${additional_revenue * 12:,.2f} annually.

"""
            
            detailed_results = f"""
=== DETAILED ANALYSIS ===

INPUT PARAMETERS:
• Monthly Ad Calls: {monthly_ad_calls:,.0f}
• Current Bid Rate: {current_bid_rate*100:.1f}%
• Target Bid Rate: {target_bid_rate*100:.1f}%
• Revenue per Billion Calls: ${revenue_per_billion:,.2f}

CURRENT PERFORMANCE:
• Current Biddable Ad Calls: {current_biddable:,.0f}
• Current Billion Equivalent: {current_billion_equiv:.5f}
• Current Monthly Revenue: ${current_monthly_revenue:,.2f}
• Current Annual Revenue: ${current_monthly_revenue * 12:,.2f}

TARGET PERFORMANCE:
• Target Biddable Ad Calls: {target_biddable:,.0f}
• Target Billion Equivalent: {target_billion_equiv:.5f}
• Target Monthly Revenue: ${target_monthly_revenue:,.2f}
• Target Annual Revenue: ${target_monthly_revenue * 12:,.2f}

REVENUE INCREASE:
• Additional Monthly Revenue: ${additional_revenue:,.2f}
• Additional Annual Revenue: ${additional_revenue * 12:,.2f}
• Revenue Increase Factor: {revenue_multiplier:.2f}x
• Percentage Increase: {percentage_increase:.2f}%
"""
            
            # Display results
            self.results_text.delete(1.0, tk.END)  # Clear previous results
            self.results_text.insert(tk.END, custom_summary + detailed_results)
            
        except ValueError as e:
            messagebox.showerror("Input Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
    
    def clear_results(self):
        """Clear the results text widget"""
        self.results_text.delete(1.0, tk.END)

class MetricRuleRow(ttk.Frame):
    """A single row for one metric rule"""
   
    def __init__(self, parent, metric_name, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        
        # Store metric name
        self.metric_name = metric_name
        
        # Define operators
        self.operators = [">", "<", "=", "Between"]
        
        # Create controls
        
        # Metric checkbox
        self.enabled = tk.BooleanVar(value=False)
        self.metric_check = ttk.Checkbutton(
            self, 
            text=metric_name, 
            variable=self.enabled,
            command=self.toggle_controls
        )
        self.metric_check.grid(row=0, column=0, padx=(0, 10), sticky="w")
        
        # Operator dropdown
        self.operator = tk.StringVar(value=self.operators[0])
        self.operator_combo = ttk.Combobox(
            self, 
            textvariable=self.operator, 
            values=self.operators, 
            width=10,
            state="disabled"
        )
        self.operator_combo.grid(row=0, column=1, padx=(0, 10))
        
        # Value entry
        vcmd = (self.register(self.validate_input), '%P')
        self.value = tk.StringVar(value="")
        self.value_entry = ttk.Entry(
            self, 
            textvariable=self.value, 
            width=20,
            state="disabled",
            validate="key", 
            validatecommand=vcmd
        )
        self.value_entry.grid(row=0, column=2, padx=(0, 10))
        
        # Add a label to indicate the unit
        unit = self.get_unit_label(metric_name)
        self.unit_label = ttk.Label(self, text=unit)
        self.unit_label.grid(row=0, column=3, padx=(0, 5))
        
        # Tooltip for the value field
        self.tooltip = None
        self.operator.trace("w", self.update_tooltip)
    
    def get_unit_label(self, metric_name):
        """Return the appropriate unit label for the metric"""
        if "Revenue" in metric_name:
            return "$"
        elif "Rate" in metric_name:
            return "%"
        elif "Requests" in metric_name:
            return "(count)"
        elif "RPB" in metric_name:
            return "(ratio)"
        return ""
    
    def validate_input(self, value):
        """Validate input based on the operator"""
        if value == "":
            return True
            
        if self.operator.get() == "Between":
            # Check for format like "10.5;20.5"
            parts = value.split(';')
            if len(parts) > 2:  # No more than one semicolon
                return False
                
            # Validate each part as a decimal
            pattern = r'^-?\d*\.?\d*$'
            return all(bool(re.match(pattern, part)) for part in parts)
        else:
            # For other operators, just validate as a decimal
            pattern = r'^-?\d*\.?\d*$'
            return bool(re.match(pattern, value))
    
    def toggle_controls(self):
        """Enable/disable controls based on checkbox state"""
        state = "normal" if self.enabled.get() else "disabled"
        self.operator_combo.config(state=state)
        self.value_entry.config(state=state)
    
    def update_tooltip(self, *args):
        """Update tooltip based on the selected operator"""
        pass

    def get_rule(self):
        """Return the rule configuration as a dictionary, converting percentage input to decimal if needed."""
        if not self.enabled.get():
            return None

        value = self.value.get()
        operator = self.operator.get()
        unit = self.unit_label.cget('text')

        # For 'Between', ensure it's properly formatted
        if operator == 'Between':
            # Verify that value contains a semicolon separator
            if ';' not in value:
                print(f"WARNING: Between value '{value}' missing semicolon separator")
                return None
                
            # Ensure both values can be converted to numbers
            try:
                min_val_str, max_val_str = value.split(';')
                min_val_str = min_val_str.strip()
                max_val_str = max_val_str.strip()
                
                # Just test that we can convert them to floats
                float(min_val_str)
                float(max_val_str)
                
                print(f"DEBUG: get_rule for metric={self.metric_name}, operator=Between, value={value}")
                rule = {
                    "metric": self.metric_name,
                    "operator": operator,
                    "value": value  # Keep as string
                }
                return rule
            except ValueError:
                print(f"WARNING: Between value '{value}' contains non-numeric values")
                return None

        # For other operators, keep existing logic
        if unit == "%" or "Rate" in self.metric_name:
            value = percentage_to_decimal(value)
        else:
            try:
                value = float(value)
            except (ValueError, TypeError):
                pass  # Keep as string if cannot convert
        rule = {
            "metric": self.metric_name,
            "operator": operator,
            "value": value
        }
        return rule


class PriorityRuleFrame(ttk.LabelFrame):
    """Frame for setting a single priority level rules"""
    
    def __init__(self, parent, priority_level, bg_color, fg_color, *args, **kwargs):
        super().__init__(parent, text=f"{priority_level} Priority Settings", *args, **kwargs)
        self.priority_level = priority_level
        self.bg_color = bg_color
        self.fg_color = fg_color
        
        # Define metrics
        self.metrics = ["Revenue", "Bid Rate", "Ad Requests", "Win Rate", "RPB"]
        self.metric_rows = []
        
        # Description
        desc = {
            "High": "Items matching these criteria will be highlighted with red background and white text",
            "Medium": "Items matching these criteria will be highlighted with orange background and black text",
            "Low": "Items matching these criteria will be highlighted with green background and black text"
        }
        
        # When "Between" is selected instruction
        if priority_level == "High":  # Only show on the first frame
            between_frame = ttk.Frame(self)
            between_frame.pack(fill="x", pady=5)
            between_label = ttk.Label(
                between_frame,
                text=("Note: For 'Between' operator, use semicolon (;) to separate min and max values. "
                      "Use a dot (.) as the decimal separator. Example: 10.5;20.5"),
                font=("Arial", 9, "italic"),
                wraplength=500
            )
            between_label.pack(anchor="w", padx=5)
        
        ttk.Label(self, text=desc.get(priority_level, "")).pack(anchor="w", padx=5, pady=(5,10))
        
        # Create a frame for the examples
        example_frame = ttk.Frame(self)
        example_frame.pack(fill="x", pady=5)
        
        # Example preview
        ttk.Label(example_frame, text="Preview: ").pack(side="left", padx=5)
        preview = ttk.Label(
            example_frame, 
            text="EXAMPLE.COM", 
            background=bg_color, 
            foreground=fg_color, 
            font=("Arial", 10, "bold"),
            relief="solid",
            borderwidth=1,
            padding=3
        )
        preview.pack(side="left", padx=5)
        
        # Instructions
        ttk.Label(
            self, 
            text="Check the metrics you want to use, select an operator, and enter values:",
            wraplength=500
        ).pack(anchor="w", padx=5, pady=(10,5))
        
        # Create frames for each metric
        for metric in self.metrics:
            metric_row = MetricRuleRow(self, metric)
            metric_row.pack(fill="x", padx=5, pady=3)
            self.metric_rows.append(metric_row)
    
    def get_rules(self):
        """Return all active rules for this priority level"""
        rules = []
        for row in self.metric_rows:
            rule = row.get_rule()
            if rule:
                rule["priority"] = self.priority_level
                rules.append(rule)
        return rules


class FileSelectionFrame(ttk.LabelFrame):
    """Frame for selecting input and output files"""
    
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, text="CSV File Selection", *args, **kwargs)
        
        # Variables to store file paths
        self.csv_a_path = tk.StringVar()
        self.csv_b_path = tk.StringVar()
        self.output_csv_path = tk.StringVar()
        
        # Create file selection controls
        file_frame = ttk.Frame(self, padding=5)
        file_frame.pack(fill="x", expand=True)
        
        # CSV A (Latest) - renamed to Inventory KPI CSV Reference Period
        ttk.Label(file_frame, text="Inventory KPI CSV Reference Period:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.csv_a_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_csv_a).grid(row=0, column=2, padx=5, pady=5)
        # Add legend for CSV A
        ttk.Label(file_frame, text="Reference period example: last 7 days, 30 days, Quarter etc...", 
                 font=("Arial", 8, "italic")).grid(row=1, column=0, columnspan=3, sticky="w", padx=5)
        
        # CSV B (Oldest) - renamed to Inventory KPI CSV Comparison Period
        ttk.Label(file_frame, text="Inventory KPI CSV Comparison Period:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.csv_b_path, width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_csv_b).grid(row=2, column=2, padx=5, pady=5)
        # Add legend for CSV B
        ttk.Label(file_frame, text="Comparison period example: 7 days before last 7 days", 
                 font=("Arial", 8, "italic")).grid(row=3, column=0, columnspan=3, sticky="w", padx=5)
        
        # Output CSV - renamed to Output File Path
        ttk.Label(file_frame, text="Output File Path:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.output_csv_path, width=50).grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_output_csv).grid(row=4, column=2, padx=5, pady=5)
    
    # Browse for CSV A file (latest)
    def browse_csv_a(self):
        filename = filedialog.askopenfilename(
            title="Select Inventory KPI CSV Reference Period",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_a_path.set(filename)
    
    # Browse for CSV B file (oldest)
    def browse_csv_b(self):
        filename = filedialog.askopenfilename(
            title="Select Inventory KPI CSV Comparison Period",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_b_path.set(filename)
    
    # Browse for output CSV file location
    def browse_output_csv(self):
        initial_file = ""
        if self.csv_a_path.get():
            # Default output filename based on input A
            base_dir = os.path.dirname(self.csv_a_path.get())
            base_name = os.path.basename(self.csv_a_path.get())
            initial_file = os.path.join(base_dir, f"output_{base_name}")
        
        filename = filedialog.asksaveasfilename(
            title="Save Output File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=initial_file
        )
        if filename:
            self.output_csv_path.set(filename)
    
    # Return all selected file paths
    def get_file_paths(self):
        return {
            "csv_a_path": self.csv_a_path.get(),
            "csv_b_path": self.csv_b_path.get(),
            "output_csv_path": self.output_csv_path.get()
        }




class SupplyChainFrame(ttk.LabelFrame):
    """Frame for Supply Chain Validation and Error Distribution"""
    
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, text="Supply Chain Validation & Error Distribution", *args, **kwargs)
        
        # File paths
        self.supply_chain_file = tk.StringVar()
        self.lines_referential_file = tk.StringVar()
        self.error_dist_file = tk.StringVar()
        
        # Create file selection section
        file_frame = ttk.Frame(self, padding=5)
        file_frame.pack(fill="x", expand=True)
        
        # Supply Chain File (XLSX) - renamed to Publisher's Supply Chain Validation Report
        ttk.Label(file_frame, text="Publisher's Supply Chain Validation Report:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.supply_chain_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_supply_chain).grid(row=0, column=2, padx=5, pady=5)
        
        # Lines Referential File (CSV) - renamed to Ads.txt lines reference file
        ttk.Label(file_frame, text="Ads.txt lines reference file:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.lines_referential_file, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_lines_referential).grid(row=1, column=2, padx=5, pady=5)
        # Add legend for lines referential
        ttk.Label(file_frame, text="See Alex CSM for updated version", 
                 font=("Arial", 8, "italic")).grid(row=2, column=0, columnspan=3, sticky="w", padx=5)
        
        # Error Distribution File (CSV) - renamed to CSM Error CSV
        ttk.Label(file_frame, text="CSM Error CSV:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.error_dist_file, width=50).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_error_dist).grid(row=3, column=2, padx=5, pady=5)
        # Add legend for error distribution
        ttk.Label(file_frame, text="Get the CSV for Adcalls by Website and CSM Error for this publisher in CSM: Config Error", 
                 font=("Arial", 8, "italic")).grid(row=4, column=0, columnspan=3, sticky="w", padx=5)
        
        # Note about output
        ttk.Label(file_frame, text="Output: Will use the Domains Highlight output file", font=("Arial", 9, "italic")).grid(row=5, column=0, columnspan=3, sticky="w", pady=5)
    
    def browse_supply_chain(self):
        filename = filedialog.askopenfilename(
            title="Select Publisher's Supply Chain Validation Report",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.supply_chain_file.set(filename)
    
    def browse_lines_referential(self):
        filename = filedialog.askopenfilename(
            title="Select Ads.txt lines reference file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.lines_referential_file.set(filename)
    
    def browse_error_dist(self):
        filename = filedialog.askopenfilename(
            title="Select CSM Error CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.error_dist_file.set(filename)
    
    def get_file_paths(self):
        """Return all selected file paths"""
        return {
            "supply_chain_file": self.supply_chain_file.get(),
            "lines_referential_file": self.lines_referential_file.get(),
            "error_dist_file": self.error_dist_file.get()
        }


class TargetRevenueCalculator(ttk.Frame):
    """Target Revenue Mail Calculator embedded as a frame"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(fill=tk.BOTH, expand=True)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        self.create_widgets()
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Target Revenue Mail Calculator", 
                              font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Input fields
        row = 1
        
        # Domain Name
        ttk.Label(main_frame, text="Domain Name:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.domain_var = tk.StringVar()
        domain_entry = ttk.Entry(main_frame, textvariable=self.domain_var, width=30)
        domain_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Current CPM
        ttk.Label(main_frame, text="Current CPM ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.current_cpm_var = tk.StringVar()
        current_cpm_entry = ttk.Entry(main_frame, textvariable=self.current_cpm_var, width=20)
        current_cpm_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Target CPM
        ttk.Label(main_frame, text="Target CPM ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.target_cpm_var = tk.StringVar()
        target_cpm_entry = ttk.Entry(main_frame, textvariable=self.target_cpm_var, width=20)
        target_cpm_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Monthly Impressions
        ttk.Label(main_frame, text="Monthly Impressions:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.impressions_var = tk.StringVar()
        impressions_entry = ttk.Entry(main_frame, textvariable=self.impressions_var, width=20)
        impressions_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Generate button
        generate_btn = ttk.Button(main_frame, text="Generate Mail Content", 
                                command=self.generate_mail)
        generate_btn.grid(row=row, column=0, columnspan=2, pady=20)
        row += 1
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Mail Content", padding="15")
        results_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        results_frame.columnconfigure(0, weight=1)
        
        # Text area for mail content
        self.mail_text = tk.Text(results_frame, height=15, wrap=tk.WORD)
        self.mail_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Scrollbar for text area
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.mail_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.mail_text.configure(yscrollcommand=scrollbar.set)
        
        # Copy button
        copy_btn = ttk.Button(results_frame, text="Copy to Clipboard", 
                             command=self.copy_to_clipboard)
        copy_btn.grid(row=1, column=0, columnspan=2, pady=10)
    
    def generate_mail(self):
        # Clear previous content
        self.mail_text.delete(1.0, tk.END)
        
        # Get input values
        try:
            domain = self.domain_var.get().strip()
            current_cpm = float(self.current_cpm_var.get().strip() or 0)
            target_cpm = float(self.target_cpm_var.get().strip() or 0)
            impressions = int(self.impressions_var.get().strip().replace(',', '') or 0)
            
            # Validate inputs
            if not domain:
                messagebox.showerror("Error", "Please enter a domain name")
                return
                
            if current_cpm <= 0 or target_cpm <= 0 or impressions <= 0:
                messagebox.showerror("Error", "Please enter valid positive numbers for CPM and impressions")
                return
            
            # Calculate values
            current_monthly_revenue = (current_cpm / 1000) * impressions
            target_monthly_revenue = (target_cpm / 1000) * impressions
            revenue_increase = target_monthly_revenue - current_monthly_revenue
            percentage_increase = (revenue_increase / current_monthly_revenue) * 100
            
            # Format numbers with commas and round to 2 decimal places
            formatted_current_revenue = f"${current_monthly_revenue:,.2f}"
            formatted_target_revenue = f"${target_monthly_revenue:,.2f}"
            formatted_increase = f"${revenue_increase:,.2f}"
            formatted_percentage = f"{percentage_increase:.1f}%"
            formatted_impressions = f"{impressions:,}"
            
            # Generate the email template
            mail_content = f"""Subject: {domain} - Target Revenue Opportunity

Hi Team,

I hope this email finds you well. I wanted to discuss an opportunity regarding {domain}.

Currently, {domain} is generating approximately {formatted_current_revenue} in monthly revenue based on {formatted_impressions} impressions at a CPM of ${current_cpm:.2f}.

We believe there is potential to increase the CPM to ${target_cpm:.2f}, which would result in a monthly revenue of {formatted_target_revenue}. This represents an increase of {formatted_increase} or {formatted_percentage} over the current revenue.

Can we discuss strategies to achieve this target? Some potential approaches could include:

1. Optimizing ad placement and formats
2. Reviewing price floors
3. Exploring new demand partners
4. Implementing header bidding improvements

Please let me know your thoughts and when we could schedule a call to discuss this opportunity further.

Best regards,
[Your Name]"""
            
            # Display the email template
            self.mail_text.insert(tk.END, mail_content)
            
        except ValueError as e:
            messagebox.showerror("Error", f"Please enter valid numbers: {str(e)}")
    
    def copy_to_clipboard(self):
        # Get mail content
        mail_content = self.mail_text.get(1.0, tk.END)
        
        # Copy to clipboard
        self.clipboard_clear()
        self.clipboard_append(mail_content)
        
        messagebox.showinfo("Success", "Mail content copied to clipboard!")


class MainApplication(ttk.Frame):
    """Main application window for Domains Highlighter"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        parent.title("Domains Highlighter Tool V3")
        parent.geometry("800x1700")
        
        # Status variables
        self.status_var = tk.StringVar(value="Ready")
        self.progress_var = tk.IntVar(value=0)
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.domains_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.domains_tab, text="Domains Highlight")
        
        # Create the Bidrate Improvement tab
        self.bidrate_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.bidrate_tab, text="Bid Rate Improvement")
        
        # Create the Target Revenue Mail tab
        self.target_revenue_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.target_revenue_tab, text="Target Revenue Mail")
        
        # Setup domains highlight tab
        self.setup_domains_tab()
        
        # Setup bidrate improvement tab
        self.setup_bidrate_tab()
        
        # Setup target revenue mail tab
        self.setup_target_revenue_tab()
        
        # Status frame at bottom
        status_frame = ttk.LabelFrame(self, text="Status")
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Status text widget
        self.status_text = tk.Text(status_frame, height=5, wrap=tk.WORD)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            self, 
            orient=tk.HORIZONTAL, 
            length=100, 
            mode='determinate',
            variable=self.progress_var
        )
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        
        self.pack(fill=tk.BOTH, expand=True)
    
    def setup_domains_tab(self):
        """Setup the Domains Highlight tab content"""
        # Create a canvas for scrolling
        canvas = tk.Canvas(self.domains_tab)
        scrollbar = ttk.Scrollbar(self.domains_tab, orient="vertical", command=canvas.yview)
        
        # Create a frame inside the canvas for all content
        self.domains_content = ttk.Frame(canvas)
        
        # Configure the canvas
        self.domains_content.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.domains_content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Create file selection frame
        self.file_selection = FileSelectionFrame(self.domains_content)
        self.file_selection.pack(fill="x", padx=10, pady=10)

        # Create Supply Chain validation frame
        self.supply_chain = SupplyChainFrame(self.domains_content)
        self.supply_chain.pack(fill="x", padx=10, pady=10)
        
        # Create priority rule frames
        priority_container = ttk.LabelFrame(self.domains_content, text="Metric Rules")
        priority_container.pack(fill="x", padx=10, pady=10)

        # Create a frame for each priority level
        self.priority_frames = []

        # High Priority Frame
        high_priority = PriorityRuleFrame(priority_container, "High", "red", "white")
        high_priority.pack(fill="x", padx=5, pady=5)
        self.priority_frames.append(high_priority)

        # Medium Priority Frame
        medium_priority = PriorityRuleFrame(priority_container, "Medium", "orange", "black")
        medium_priority.pack(fill="x", padx=5, pady=5)
        self.priority_frames.append(medium_priority)

        # Low Priority Frame
        low_priority = PriorityRuleFrame(priority_container, "Low", "green", "black")
        low_priority.pack(fill="x", padx=5, pady=5)
        self.priority_frames.append(low_priority)
        
        # Process button
        process_frame = ttk.Frame(self.domains_content)
        process_frame.pack(fill="x", padx=10, pady=20)
        
        process_button = ttk.Button(
            process_frame,
            text="Process Files",
            command=self.process_files
        )
        process_button.pack(pady=10)

    def setup_bidrate_tab(self):
        """Setup the Bid Rate Improvement tab content"""
        # Create the BidRateImprover frame in this tab
        self.bid_rate_improver = BidRateImprover(self.bidrate_tab)

        # The BidRateImprover class handles all the UI and functionality
        # We don't need to do anything else here as the class is self-contained
    
    def setup_target_revenue_tab(self):
        """Setup the Target Revenue Mail Calculator tab content"""
        # Create embedded version of the Revenue Target Calculator
        class EmbeddedRevenueCalculator(ttk.Frame):
            """Embedded version of the Revenue Target Calculator"""
            
            def __init__(self, parent):
                super().__init__(parent)
                self.pack(fill=tk.BOTH, expand=True)
                
                # Create notebook for tabs
                self.notebook = ttk.Notebook(self)
                self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                # Input tab
                self.input_frame = ttk.Frame(self.notebook, padding="20")
                self.notebook.add(self.input_frame, text="Input & Calculate")
                
                # Results tab
                self.results_frame = ttk.Frame(self.notebook, padding="20")
                self.notebook.add(self.results_frame, text="Results")
                
                # Configure grid weights
                self.input_frame.columnconfigure(1, weight=1)
                
                # Variables
                self.current_revenue_var = tk.StringVar(value="48000")
                self.ad_requests_var = tk.StringVar(value="30000")
                self.rpb_var = tk.StringVar(value="16000")
                self.cpm_var = tk.StringVar(value="2.50")
                self.bid_rate_var = tk.StringVar(value="2.0")
                self.win_rate_var = tk.StringVar(value="15.0")
                self.target_revenue_var = tk.StringVar(value="96000")
                
                self.create_input_tab()
                self.create_results_tab()
            
            def create_input_tab(self):
                # Current Metrics Section
                current_section = ttk.LabelFrame(self.input_frame, text="Current Metrics", padding="15")
                current_section.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
                current_section.columnconfigure(1, weight=1)
                
                row = 0
                
                # Current Revenue
                ttk.Label(current_section, text="Current Revenue ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
                ttk.Entry(current_section, textvariable=self.current_revenue_var, width=20).grid(
                    row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
                row += 1
                
                # Ad Requests (millions)
                ttk.Label(current_section, text="Ad Requests (millions):").grid(row=row, column=0, sticky=tk.W, pady=5)
                ttk.Entry(current_section, textvariable=self.ad_requests_var, width=20).grid(
                    row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
                row += 1
                
                # RPB (Revenue Per Billion)
                ttk.Label(current_section, text="RPB - Revenue Per Billion ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
                ttk.Entry(current_section, textvariable=self.rpb_var, width=20).grid(
                    row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
                row += 1
                
                # CPM
                ttk.Label(current_section, text="CPM ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
                ttk.Entry(current_section, textvariable=self.cpm_var, width=20).grid(
                    row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
                row += 1
                
                # Bid Rate (%)
                ttk.Label(current_section, text="Bid Rate (%):").grid(row=row, column=0, sticky=tk.W, pady=5)
                ttk.Entry(current_section, textvariable=self.bid_rate_var, width=20).grid(
                    row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
                row += 1
                
                # Win Rate (%)
                ttk.Label(current_section, text="Win Rate (%):").grid(row=row, column=0, sticky=tk.W, pady=5)
                ttk.Entry(current_section, textvariable=self.win_rate_var, width=20).grid(
                    row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
                
                # Target Section
                target_section = ttk.LabelFrame(self.input_frame, text="Target", padding="15")
                target_section.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 15))
                target_section.columnconfigure(1, weight=1)
                
                # Target Revenue
                ttk.Label(target_section, text="Target Revenue ($):").grid(row=0, column=0, sticky=tk.W, pady=5)
                ttk.Entry(target_section, textvariable=self.target_revenue_var, width=20).grid(
                    row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
                
                # Buttons
                button_frame = ttk.Frame(self.input_frame)
                button_frame.grid(row=2, column=0, columnspan=2, pady=20)
                
                ttk.Button(button_frame, text="Calculate Targets", command=self.calculate_targets, padding=10).grid(
                    row=0, column=0, padx=5)
                ttk.Button(button_frame, text="Clear All", command=self.clear_all, padding=10).grid(
                    row=0, column=1, padx=5)
            
            def create_results_tab(self):
                # Create a text widget for displaying results
                self.results_text = tk.Text(self.results_frame, height=30, width=80, wrap=tk.WORD)
                self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
                
                # Add a scrollbar
                scrollbar = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.results_text.yview)
                scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
                self.results_text.configure(yscrollcommand=scrollbar.set)
                
                # Configure grid weights
                self.results_frame.columnconfigure(0, weight=1)
                self.results_frame.rowconfigure(0, weight=1)
                
                # Copy button
                ttk.Button(self.results_frame, text="Copy Results", command=self.copy_results).grid(
                    row=1, column=0, pady=10)
            
            def calculate_revenue_cpm(self, ad_requests, cpm, bid_rate, win_rate):
                """Calculate revenue using CPM formula"""
                # Convert to correct units
                ad_requests_in_thousands = ad_requests * 1000  # Convert from millions to thousands
                bid_rate_decimal = bid_rate / 100
                win_rate_decimal = win_rate / 100
                
                # Calculate impressions
                impressions = ad_requests_in_thousands * bid_rate_decimal * win_rate_decimal
                
                # Calculate revenue using CPM formula
                revenue = (cpm * impressions) / 1000
                return revenue
            
            def calculate_revenue_rpb(self, ad_requests, rpb, bid_rate):
                """Calculate revenue using RPB formula"""
                # Convert to correct units
                ad_requests_in_billions = ad_requests / 1000  # Convert from millions to billions
                bid_rate_decimal = bid_rate / 100
                
                # Calculate revenue using RPB formula
                revenue = rpb * ad_requests_in_billions * bid_rate_decimal
                return revenue
            
            def calculate_targets(self):
                try:
                    # Get values from input fields
                    current_revenue = float(self.current_revenue_var.get())
                    ad_requests = float(self.ad_requests_var.get())
                    rpb = float(self.rpb_var.get())
                    cpm = float(self.cpm_var.get())
                    bid_rate = float(self.bid_rate_var.get())
                    win_rate = float(self.win_rate_var.get())
                    target_revenue = float(self.target_revenue_var.get())
                    
                    # Validate inputs
                    if current_revenue <= 0 or target_revenue <= 0:
                        messagebox.showerror("Error", "Revenue values must be positive")
                        return
                    
                    # Calculate multiplier needed
                    required_multiplier = target_revenue / current_revenue
                    
                    # Calculate current metrics
                    calculated_current_cpm = cpm
                    calculated_current_rpb = rpb
                    
                    # Calculate individual targets
                    targets = self.calculate_individual_targets(
                        ad_requests, rpb, cpm, bid_rate, win_rate,
                        current_revenue, calculated_current_cpm, calculated_current_rpb, 
                        target_revenue, required_multiplier
                    )
                    
                    # Display results
                    self.display_results(targets)
                    
                except ValueError as e:
                    messagebox.showerror("Error", f"Please enter valid numbers: {str(e)}")
            
            def calculate_individual_targets(self, ad_requests, rpb, cpm, bid_rate, win_rate,
                                            current_revenue, calculated_current_cpm, calculated_current_rpb, 
                                            target_revenue, required_multiplier):
                """Calculate what each individual metric would need to be to reach target"""
                targets = {}
                
                # Calculate Ad Requests target
                ar_target = ad_requests * required_multiplier
                targets['ad_requests'] = {
                    'current': ad_requests,
                    'target': ar_target,
                    'increase_abs': ar_target - ad_requests,
                    'increase_pct': ((ar_target / ad_requests) - 1) * 100
                }
                
                # Calculate RPB target
                rpb_target = rpb * required_multiplier
                targets['rpb'] = {
                    'current': rpb,
                    'target': rpb_target,
                    'increase_abs': rpb_target - rpb,
                    'increase_pct': ((rpb_target / rpb) - 1) * 100
                }
                
                # Calculate CPM target
                cpm_target = cpm * required_multiplier
                targets['cpm'] = {
                    'current': cpm,
                    'target': cpm_target,
                    'increase_abs': cpm_target - cpm,
                    'increase_pct': ((cpm_target / cpm) - 1) * 100
                }
                
                # Calculate Bid Rate target
                br_target = bid_rate * required_multiplier
                targets['bid_rate'] = {
                    'current': bid_rate,
                    'target': br_target,
                    'increase_abs': br_target - bid_rate,
                    'increase_pct': ((br_target / bid_rate) - 1) * 100
                }
                
                # Calculate Win Rate target
                wr_target = win_rate * required_multiplier
                targets['win_rate'] = {
                    'current': win_rate,
                    'target': wr_target,
                    'increase_abs': wr_target - win_rate,
                    'increase_pct': ((wr_target / win_rate) - 1) * 100
                }
                
                return targets
            
            def display_results(self, targets):
                """Display the calculated targets in the results tab"""
                # Clear previous results
                self.results_text.delete(1.0, tk.END)
                
                # Format results
                output = []
                
                # Header
                output.append("REVENUE TARGET ANALYSIS")
                output.append("=" * 30)
                output.append(f"Current Revenue: ${float(self.current_revenue_var.get()):,.2f}")
                output.append(f"Target Revenue: ${float(self.target_revenue_var.get()):,.2f}")
                required_multiplier = float(self.target_revenue_var.get()) / float(self.current_revenue_var.get())
                output.append(f"Required Increase: {(required_multiplier - 1) * 100:.1f}%")
                output.append("")
                output.append("INDIVIDUAL METRIC TARGETS")
                output.append("-" * 30)
                output.append("Each metric shown below would need to increase to the target value")
                output.append("if all other metrics remained constant.")
                output.append("")
                
                # Ad Requests
                ar = targets['ad_requests']
                output.append(f"AD REQUESTS (millions):")
                output.append(f"  Current: {ar['current']:,.1f}")
                output.append(f"  Target:  {ar['target']:,.1f}")
                output.append(f"  Increase: +{ar['increase_pct']:.1f}% (+{ar['increase_abs']:,.1f})")
                output.append("")
                
                # RPB
                rpb = targets['rpb']
                output.append(f"RPB (Revenue Per Billion):")
                output.append(f"  Current: ${rpb['current']:,.2f}")
                output.append(f"  Target:  ${rpb['target']:,.2f}")
                output.append(f"  Increase: +{rpb['increase_pct']:.1f}% (+${rpb['increase_abs']:,.2f})")
                output.append("")
                
                # CPM
                cpm = targets['cpm']
                output.append(f"CPM:")
                output.append(f"  Current: ${cpm['current']:.2f}")
                output.append(f"  Target:  ${cpm['target']:.2f}")
                output.append(f"  Increase: +{cpm['increase_pct']:.1f}% (+${cpm['increase_abs']:.2f})")
                output.append("")
                
                # Bid Rate
                br = targets['bid_rate']
                feasible_br = "✓ Feasible" if br['target'] <= 100 else "⚠ May not be feasible (>100%)"  
                output.append(f"BID RATE:")
                output.append(f"  Current: {br['current']:.1f}%")
                output.append(f"  Target:  {br['target']:.1f}%")
                output.append(f"  Increase: +{br['increase_pct']:.1f}% (+{br['increase_abs']:.1f} percentage points)")
                output.append(f"  Feasibility: {feasible_br}")
                output.append("")
                
                # Win Rate
                wr = targets['win_rate']
                feasible_wr = "✓ Feasible" if wr['target'] <= 100 else "⚠ May not be feasible (>100%)"
                output.append(f"WIN RATE:")
                output.append(f"  Current: {wr['current']:.1f}%")
                output.append(f"  Target:  {wr['target']:.1f}%")
                output.append(f"  Increase: +{wr['increase_pct']:.1f}% (+{wr['increase_abs']:.1f} percentage points)")
                output.append(f"  Feasibility: {feasible_wr}")
                output.append("")
                
                # Recommendations
                output.append("RECOMMENDATIONS:")
                output.append("-" * 20)
                
                # Find the most feasible improvements
                improvements = []
                if br['target'] <= 100:
                    improvements.append(('Bid Rate', br['increase_pct']))
                if wr['target'] <= 100:
                    improvements.append(('Win Rate', wr['increase_pct']))
                improvements.append(('CPM', cpm['increase_pct']))
                improvements.append(('RPB', rpb['increase_pct']))
                improvements.append(('Ad Requests', ar['increase_pct']))
                
                # Sort by required increase (ascending)
                improvements.sort(key=lambda x: x[1])
                
                output.append("Most feasible single-metric improvements (ranked by required increase):")
                for i, (metric, increase) in enumerate(improvements[:3], 1):
                    output.append(f"{i}. {metric}: +{increase:.1f}%")
                output.append("")
                output.append("Consider combining multiple smaller improvements for more realistic targets.")
                
                # Display in text widget
                result_text = "\n".join(output)
                self.results_text.insert(1.0, result_text)
                
                # Switch to results tab
                self.notebook.select(1)
            
            def copy_results(self):
                try:
                    results = self.results_text.get(1.0, tk.END)
                    self.clipboard_clear()
                    self.clipboard_append(results)
                    messagebox.showinfo("Success", "Results copied to clipboard!")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy results: {str(e)}")
            
            def clear_all(self):
                # Reset all input fields to default values
                self.current_revenue_var.set("48000")
                self.ad_requests_var.set("30000")
                self.rpb_var.set("16000")
                self.cpm_var.set("2.50")
                self.bid_rate_var.set("2.0")
                self.win_rate_var.set("15.0")
                self.target_revenue_var.set("96000")
                
                # Clear results
                self.results_text.delete(1.0, tk.END)
        
        # Create the embedded calculator
        self.revenue_calculator = EmbeddedRevenueCalculator(self.target_revenue_tab)
    

    
    def log_status(self, message):
        """Add a message to the status text widget"""
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.update_idletasks()
    
    def process_files(self):
        """Process the selected files with the defined rules"""
        # Disable the process button while processing
        for widget in self.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.configure(state="disabled")
        
        # Start the processing in a separate thread to keep the UI responsive
        processing_thread = threading.Thread(target=self._process_files_thread)
        processing_thread.daemon = True
        processing_thread.start()
    
    def _process_files_thread(self):
        """Thread function to process files"""
        try:
            # Get file paths
            file_paths = self.file_selection.get_file_paths()
            latest_csv_path = file_paths["csv_a_path"]
            oldest_csv_path = file_paths["csv_b_path"]
            output_csv_path = file_paths["output_csv_path"]
            
            # Validate input files
            if not latest_csv_path or not oldest_csv_path:
                messagebox.showerror("Error", "Please select both input CSV files")
                return
                
            if not output_csv_path:
                messagebox.showerror("Error", "Please select an output file")
                return
            
            # Get rules from each priority frame
            rules = {}
            for frame in self.priority_frames:
                priority_rules = frame.get_rules()
                if priority_rules:
                    rules[frame.priority_level] = priority_rules
            
            self.log_status("Starting Domains Highlight processing...")
            self.progress_var.set(10)
            
            # Process the domains files (data only, no formatting yet)
            output_path, domains_df = calculate_and_save_differences(
                latest_csv_path,
                oldest_csv_path,
                output_csv_path,
                rules,
                apply_formatting=False  # Don't apply formatting yet
            )
            
            self.log_status(f"Domains Highlight data processed and saved to {output_path}")
            self.progress_var.set(30)
            
            # Initialize variables for optional processing
            supply_chain_df = None
            error_dist_df = None
            error_dist_summary = None
            
            # Check if we need to process supply chain data
            supply_chain_paths = self.supply_chain.get_file_paths()
            supply_chain_file = supply_chain_paths["supply_chain_file"]
            lines_ref_file = supply_chain_paths["lines_referential_file"]
            
            if supply_chain_file and lines_ref_file:
                self.log_status("Processing Supply Chain data...")
                
                # Process the supply chain files
                supply_chain_df = process_supply_chain_files(
                    supply_chain_file,
                    lines_ref_file
                )
                
                if supply_chain_df is not None and not supply_chain_df.empty:
                    self.log_status(f"Supply Chain data processed. Shape: {supply_chain_df.shape}")
                else:
                    self.log_status("No Supply Chain data to process.")
            
            self.progress_var.set(50)
            
            # Check if we need to process error distribution data
            error_dist_file = supply_chain_paths.get("error_dist_file")
            
            if error_dist_file:
                self.log_status("Processing Error Distribution data...")
                
                try:
                    # Process the error distribution file
                    error_dist_df, error_dist_summary = process_error_distribution(error_dist_file)
                    
                    if error_dist_df is not None and not error_dist_df.empty:
                        self.log_status(f"Error Distribution data processed. Shape: {error_dist_df.shape}")
                    else:
                        self.log_status("No Error Distribution data to process.")
                except Exception as e:
                    self.log_status(f"Error processing Error Distribution data: {str(e)}")
                    error_dist_df = None
                    error_dist_summary = None
            else:
                self.log_status("No Error Distribution file selected, skipping Error Distribution processing.")
                error_dist_df = None
                error_dist_summary = None
            
            self.progress_var.set(70)
            
            # Write all data to Excel
            self.log_status("Writing all datasets to Excel...")
            
            # Import necessary functions from excel_helper
            from excel_helper import write_to_excel_with_two_sheets, apply_consistent_formatting, write_error_distribution_sheet
            
            # First write the domains and supply chain data
            if supply_chain_df is not None and not supply_chain_df.empty:
                result = write_to_excel_with_two_sheets(
                    output_path,
                    domains_df,
                    supply_chain_df,
                    apply_formatting=False,  # Don't apply formatting yet
                    rules=rules
                )
                
                if not result:
                    self.log_status("Failed to create Excel file with both sheets")
                    return
            
            # Then add error distribution data if available
            if error_dist_df is not None and not error_dist_df.empty:
                write_error_distribution_sheet(
                    output_path,
                    error_dist_df,
                    error_dist_summary,
                    apply_formatting=False  # Don't apply formatting yet
                )
                
                self.log_status("Error Distribution data added to Excel file")
            
            # Now apply consistent formatting to all sheets in a separate pass
            self.log_status("Applying consistent formatting to all sheets...")
            apply_consistent_formatting(output_path, domains_df, rules, error_dist_summary)
            
            self.log_status("Successfully created Excel file with all data!")
            self.progress_var.set(100)
            
            # Ask if user wants to open the file
            if messagebox.askyesno("Success", f"Processing complete! Open the output file?\n\n{output_path}"):
                os.startfile(output_path)
                
        except Exception as e:
            self.log_status(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            # Re-enable the process button
            for widget in self.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.configure(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApplication(root)
    root.mainloop()
