import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as messagebox

class AdRevenueCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Ad Revenue Calculator")
        self.root.geometry("700x700")
        self.root.resizable(True, True)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
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
• Biddable Requests: {current_biddable:,.0f}
• Monthly Revenue: ${current_monthly_revenue:,.2f}
• Annual Revenue: ${current_monthly_revenue * 12:,.2f}

TARGET PERFORMANCE:
• Biddable Requests: {target_biddable:,.0f}
• Monthly Revenue: ${target_monthly_revenue:,.2f}
• Annual Revenue: ${target_monthly_revenue * 12:,.2f}

REVENUE IMPACT:
• Additional Monthly Revenue: ${additional_revenue:,.2f}
• Additional Annual Revenue: ${additional_revenue * 12:,.2f}
"""
            
            if revenue_multiplier != float('inf'):
                detailed_results += f"• Revenue Multiplier: {revenue_multiplier:.1f}x\n"
                detailed_results += f"• Percentage Increase: {percentage_increase:.0f}%\n"
            else:
                detailed_results += "• Revenue Multiplier: ∞ (from $0)\n"
                detailed_results += "• Percentage Increase: ∞%\n"
            
            # Combine custom summary with detailed results
            full_results = custom_summary + detailed_results
            
            # Display results
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(1.0, full_results)
            
        except ValueError as e:
            messagebox.showerror("Input Error", f"Please check your inputs:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred:\n{str(e)}")
    
    def clear_results(self):
        self.results_text.delete(1.0, tk.END)

def main():
    root = tk.Tk()
    app = AdRevenueCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()