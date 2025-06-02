import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as messagebox
import math

class RevenueTargetCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Revenue Target Calculator")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame with notebook for tabs
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Revenue Target Calculator", 
                               font=('Arial', 18, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Input tab
        input_frame = ttk.Frame(notebook, padding="20")
        notebook.add(input_frame, text="Input & Calculate")
        
        # Results tab
        results_frame = ttk.Frame(notebook, padding="20")
        notebook.add(results_frame, text="Results")
        
        self.create_input_tab(input_frame)
        self.create_results_tab(results_frame)
        
    def create_input_tab(self, parent):
        parent.columnconfigure(1, weight=1)
        
        # Current Metrics Section
        current_section = ttk.LabelFrame(parent, text="Current Metrics", padding="15")
        current_section.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        current_section.columnconfigure(1, weight=1)
        
        row = 0
        
        # Current Revenue
        ttk.Label(current_section, text="Current Revenue ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.current_revenue_var = tk.StringVar(value="48000")
        ttk.Entry(current_section, textvariable=self.current_revenue_var, width=20).grid(
            row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Ad Requests (millions)
        ttk.Label(current_section, text="Ad Requests (millions):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.ad_requests_var = tk.StringVar(value="30000")
        ttk.Entry(current_section, textvariable=self.ad_requests_var, width=20).grid(
            row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # RPB (Revenue Per Billion)
        ttk.Label(current_section, text="RPB - Revenue Per Billion ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.rpb_var = tk.StringVar(value="16000")
        ttk.Entry(current_section, textvariable=self.rpb_var, width=20).grid(
            row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # CPM
        ttk.Label(current_section, text="CPM ($):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.cpm_var = tk.StringVar(value="2.50")
        ttk.Entry(current_section, textvariable=self.cpm_var, width=20).grid(
            row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Bid Rate (%)
        ttk.Label(current_section, text="Bid Rate (%):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.bid_rate_var = tk.StringVar(value="2.0")
        ttk.Entry(current_section, textvariable=self.bid_rate_var, width=20).grid(
            row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Win Rate (%)
        ttk.Label(current_section, text="Win Rate (%):").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.win_rate_var = tk.StringVar(value="15.0")
        ttk.Entry(current_section, textvariable=self.win_rate_var, width=20).grid(
            row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        row += 1
        
        # Target Section
        target_section = ttk.LabelFrame(parent, text="Target", padding="15")
        target_section.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=15)
        target_section.columnconfigure(1, weight=1)
        
        # Target Revenue
        ttk.Label(target_section, text="Target Revenue ($):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.target_revenue_var = tk.StringVar(value="96000")
        ttk.Entry(target_section, textvariable=self.target_revenue_var, width=20).grid(
            row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))
        
        # Buttons
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        calculate_btn = ttk.Button(button_frame, text="Calculate Required Increases", 
                                 command=self.calculate_targets, style='Accent.TButton')
        calculate_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_btn = ttk.Button(button_frame, text="Clear All", command=self.clear_all)
        clear_btn.pack(side=tk.LEFT)
        
        # Add help text
        help_frame = ttk.LabelFrame(parent, text="How it works", padding="10")
        help_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(15, 0))
        
        help_text = """This calculator determines how much each metric needs to increase individually to reach your target revenue.
        
Revenue Formulas:
• CPM Method: Revenue = Ad Requests (millions) × Bid Rate × Win Rate × CPM × 1000
• RPB Method: Revenue = Ad Requests (millions) × Bid Rate × RPB

The calculator uses your input current revenue as the baseline and shows what each metric needs to become to reach your target."""
        
        help_label = ttk.Label(help_frame, text=help_text, wraplength=800, justify=tk.LEFT)
        help_label.pack()
        
    def create_results_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        
        # Results text widget with scrollbar
        text_frame = ttk.Frame(parent)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.results_text = tk.Text(text_frame, wrap=tk.WORD, font=('Consolas', 10))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Copy results button
        copy_btn = ttk.Button(parent, text="Copy Results to Clipboard", 
                            command=self.copy_results)
        copy_btn.grid(row=1, column=0, pady=10)
        
    def calculate_revenue_cpm(self, ad_requests, cpm, bid_rate, win_rate):
        """Calculate revenue using CPM formula"""
        # Convert percentages to decimals
        bid_rate_decimal = bid_rate / 100
        win_rate_decimal = win_rate / 100
        
        # Revenue = (Ad Requests in millions × 1,000,000) × Bid Rate × Win Rate × (CPM / 1000)
        # Simplified: Ad Requests (millions) × Bid Rate × Win Rate × CPM × 1000
        revenue = ad_requests * bid_rate_decimal * win_rate_decimal * cpm * 1000
        return revenue
    
    def calculate_revenue_rpb(self, ad_requests, rpb, bid_rate):
        """Calculate revenue using RPB formula"""
        # Convert percentage to decimal
        bid_rate_decimal = bid_rate / 100
        
        # Revenue = (Ad Requests in millions × 1000) × Bid Rate × (RPB / 1000)
        # Simplified: Ad Requests (millions) × Bid Rate × RPB
        revenue = ad_requests * bid_rate_decimal * rpb
        return revenue
        
    def calculate_targets(self):
        try:
            # Get current values
            current_revenue = float(self.current_revenue_var.get().replace(',', ''))
            ad_requests = float(self.ad_requests_var.get().replace(',', ''))
            rpb = float(self.rpb_var.get().replace(',', ''))
            cpm = float(self.cpm_var.get().replace(',', ''))
            bid_rate = float(self.bid_rate_var.get())
            win_rate = float(self.win_rate_var.get())
            target_revenue = float(self.target_revenue_var.get().replace(',', ''))
            
            # Validate inputs
            if any(val <= 0 for val in [ad_requests, rpb, cpm, bid_rate, win_rate, target_revenue]):
                raise ValueError("All values must be positive numbers")
            if bid_rate > 100 or win_rate > 100:
                raise ValueError("Bid Rate and Win Rate must be percentages (≤ 100)")
            
            # Calculate current revenue using both formulas for comparison
            calculated_current_cpm = self.calculate_revenue_cpm(ad_requests, cpm, bid_rate, win_rate)
            calculated_current_rpb = self.calculate_revenue_rpb(ad_requests, rpb, bid_rate)
            
            # Use the user's input current revenue as the baseline, but show both calculations
            
            # Calculate required multiplier using input current revenue
            if current_revenue <= 0:
                raise ValueError("Current revenue must be positive")
            
            required_multiplier = target_revenue / current_revenue
            
            # Calculate individual metric targets
            results = self.calculate_individual_targets(
                ad_requests, rpb, cpm, bid_rate, win_rate, 
                current_revenue, calculated_current_cpm, calculated_current_rpb, target_revenue, required_multiplier
            )
            
            # Display results
            self.display_results(results)
            
        except ValueError as e:
            messagebox.showerror("Input Error", f"Please check your inputs:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred:\n{str(e)}")
    
    def calculate_individual_targets(self, ad_requests, rpb, cpm, bid_rate, win_rate,
                                   current_revenue, calculated_current_cpm, calculated_current_rpb, target_revenue, required_multiplier):
        
        results = {
            'current_values': {
                'ad_requests': ad_requests,
                'rpb': rpb,
                'cpm': cpm,
                'bid_rate': bid_rate,
                'win_rate': win_rate,
                'current_revenue': current_revenue,
                'calculated_current_cpm': calculated_current_cpm,
                'calculated_current_rpb': calculated_current_rpb,
                'target_revenue': target_revenue,
                'required_multiplier': required_multiplier
            },
            'targets': {}
        }
        
        # Calculate target for Ad Requests
        target_ad_requests = ad_requests * required_multiplier
        increase_ad_requests = ((target_ad_requests - ad_requests) / ad_requests) * 100
        results['targets']['ad_requests'] = {
            'current': ad_requests,
            'target': target_ad_requests,
            'increase_pct': increase_ad_requests,
            'increase_abs': target_ad_requests - ad_requests
        }
        
        # Calculate target for RPB
        target_rpb = rpb * required_multiplier
        increase_rpb = ((target_rpb - rpb) / rpb) * 100
        results['targets']['rpb'] = {
            'current': rpb,
            'target': target_rpb,
            'increase_pct': increase_rpb,
            'increase_abs': target_rpb - rpb
        }
        
        # Calculate target for CPM
        target_cpm = cpm * required_multiplier
        increase_cpm = ((target_cpm - cpm) / cpm) * 100
        results['targets']['cpm'] = {
            'current': cpm,
            'target': target_cpm,
            'increase_pct': increase_cpm,
            'increase_abs': target_cpm - cpm
        }
        
        # Calculate target for Bid Rate
        target_bid_rate = bid_rate * required_multiplier
        increase_bid_rate = ((target_bid_rate - bid_rate) / bid_rate) * 100
        results['targets']['bid_rate'] = {
            'current': bid_rate,
            'target': target_bid_rate,
            'increase_pct': increase_bid_rate,
            'increase_abs': target_bid_rate - bid_rate
        }
        
        # Calculate target for Win Rate
        target_win_rate = win_rate * required_multiplier
        increase_win_rate = ((target_win_rate - win_rate) / win_rate) * 100
        results['targets']['win_rate'] = {
            'current': win_rate,
            'target': target_win_rate,
            'increase_pct': increase_win_rate,
            'increase_abs': target_win_rate - win_rate
        }
        
        return results
    
    def display_results(self, results):
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        
        # Format results
        output = []
        output.append("REVENUE TARGET ANALYSIS")
        output.append("=" * 50)
        output.append("")
        
        # Current situation
        current = results['current_values']
        output.append("CURRENT SITUATION:")
        output.append(f"• Input Revenue: ${current['current_revenue']:,.2f}")
        output.append(f"• Calculated Revenue (CPM method): ${current['calculated_current_cpm']:,.2f}")
        output.append(f"• Calculated Revenue (RPB method): ${current['calculated_current_rpb']:,.2f}")
        output.append(f"• Target Revenue: ${current['target_revenue']:,.2f}")
        output.append(f"• Required Multiplier: {current['required_multiplier']:.2f}x")
        output.append("")
        
        # Individual metric targets
        output.append("INDIVIDUAL METRIC TARGETS:")
        output.append("(Each metric improved individually to reach target)")
        output.append("-" * 50)
        
        targets = results['targets']
        
        # Ad Requests
        ar = targets['ad_requests']
        output.append(f"AD REQUESTS:")
        output.append(f"  Current: {ar['current']:,.0f} million")
        output.append(f"  Target:  {ar['target']:,.0f} million")
        output.append(f"  Increase: +{ar['increase_pct']:.1f}% (+{ar['increase_abs']:,.0f} million)")
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
        self.root.nametowidget('.!frame.!notebook').select(1)
    
    def copy_results(self):
        try:
            results = self.results_text.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(results)
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

def main():
    root = tk.Tk()
    app = RevenueTargetCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()