"""
Domains Highlighter Module - Handles domain highlighting logic
"""
import pandas as pd
import numpy as np
import os
from excel_helper import apply_highlighting_rules

def percentage_to_decimal(value):
    """
    Converts a percentage string or float (e.g., '11' or 11) to a decimal (e.g., 0.11).
    """
    try:
        return float(value) / 100
    except (ValueError, TypeError):
        return None

def calculate_and_save_differences(latest_csv_path, oldest_csv_path, output_csv_path, rules, apply_formatting=True):
    """
    Calculate percentage differences between latest and oldest CSVs and save to output
    
    Parameters:
        latest_csv_path (str): Path to the latest CSV file
        oldest_csv_path (str): Path to the oldest CSV file
        output_csv_path (str): Path to save the output file
        rules (dict): Dictionary of rules to apply
        
    Returns:
        tuple: (output_path, result_df) - Path to the output file and the resulting DataFrame
    """
    try:
        # Load the CSV files
        latest_df = pd.read_csv(latest_csv_path)
        oldest_df = pd.read_csv(oldest_csv_path)
        
        # Ensure both dataframes have Website/App Name column
        if 'Website/App Name' not in latest_df.columns or 'Website/App Name' not in oldest_df.columns:
            raise ValueError("Both CSV files must contain 'Website/App Name' column")
        
        # Rename columns to standard format if needed
        column_map = {
            0: 'Website/App Name',
            1: 'Revenue',
            2: 'Ad Requests',
            3: 'RPB',
            4: 'CPM',
            5: 'Bid Rate',
            6: 'Win Rate',
            7: 'Fill Rate',
            8: 'Impressions',
            9: 'Viewability',
            10: 'Time in view',
            11: 'Platform'
        }
        
        # If columns are numbered instead of named, rename them
        if latest_df.columns[0] != 'Website/App Name' and len(latest_df.columns) >= 12:
            latest_df.columns = [column_map.get(i, col) for i, col in enumerate(latest_df.columns)]
        
        if oldest_df.columns[0] != 'Website/App Name' and len(oldest_df.columns) >= 12:
            oldest_df.columns = [column_map.get(i, col) for i, col in enumerate(oldest_df.columns)]
        
        # Sort both dataframes by Website/App Name to ensure consistent ordering
        latest_df = latest_df.sort_values('Website/App Name').reset_index(drop=True)
        oldest_df = oldest_df.sort_values('Website/App Name').reset_index(drop=True)
        
        # Create lists of domains for comparison
        latest_domains = set(latest_df['Website/App Name'].dropna().tolist())
        oldest_domains = set(oldest_df['Website/App Name'].dropna().tolist())
        
        # Merge dataframes on Website/App Name using outer join
        merged_df = pd.merge(
            latest_df, 
            oldest_df,
            on='Website/App Name', 
            how='outer',
            suffixes=('', '_oldest')
        )
        
        # Sort merged dataframe to maintain consistent ordering
        merged_df = merged_df.sort_values('Website/App Name').reset_index(drop=True)
        
        # Create status for each domain
        def get_domain_status(row):
            domain = row['Website/App Name']
            if domain in latest_domains and domain in oldest_domains:
                return "Present in both"
            elif domain in latest_domains:
                return "New"
            else:
                return "Deprecated"
        
        merged_df['New and Deprecated'] = merged_df.apply(get_domain_status, axis=1)
        
        # Create a new dataframe for output with proper alignment
        result_df = pd.DataFrame()
        
        # Add Website/App Name column first
        result_df['Website/App Name'] = merged_df['Website/App Name']
        
        # Columns to process (in the specified order)
        metric_columns = [
            ('Revenue', 'B'),
            ('Ad Requests', 'C'),
            ('RPB', 'D'),
            ('CPM', 'E'),
            ('Bid Rate', 'F'),
            ('Win Rate', 'G'),
            ('Fill Rate', 'H'),
            ('Impressions', 'I'),
            ('Viewability', 'J')
        ]
        
        # Process each column in the specified order
        for col, col_letter in metric_columns:
            if col in latest_df.columns:
                # For the metric column, use values from merged_df (which includes latest values)
                # This ensures proper alignment with the merged dataset
                result_df[col] = merged_df[col]
                
                # Calculate percentage difference column
                oldest_col = f"{col}_oldest"
                diff_col = f"{col} % Diff"
                
                if oldest_col in merged_df.columns:
                    # Convert to numeric, coercing errors to NaN
                    latest_values = pd.to_numeric(merged_df[col], errors='coerce')
                    oldest_values = pd.to_numeric(merged_df[oldest_col], errors='coerce')
                    
                    # Calculate percentage difference only where both values exist
                    pct_diff = np.where(
                        (pd.notna(latest_values)) & (pd.notna(oldest_values)) & (oldest_values != 0),
                        ((latest_values - oldest_values) / oldest_values) * 100,
                        np.nan
                    )
                    
                    result_df[diff_col] = pct_diff
                else:
                    # If oldest column doesn't exist, set all differences to NaN
                    result_df[diff_col] = np.nan
                
                # Handle special cases for new and deprecated domains
                mask_new = merged_df['New and Deprecated'] == 'New'
                mask_deprecated = merged_df['New and Deprecated'] == 'Deprecated'
                
                # For new domains, keep the value but set diff to NaN
                if mask_new.any():
                    result_df.loc[mask_new, diff_col] = np.nan
                
                # For deprecated domains, set both value and diff to NaN
                if mask_deprecated.any():
                    result_df.loc[mask_deprecated, col] = np.nan
                    result_df.loc[mask_deprecated, diff_col] = np.nan
        
        # Add any remaining columns from the merged dataframe
        remaining_columns = [col for col in merged_df.columns 
                        if col not in result_df.columns 
                        and col != 'Website/App Name'
                        and not col.endswith('_oldest')]
        
        for col in remaining_columns:
            if col in merged_df.columns:
                result_df[col] = merged_df[col]
        
        # Add the "New and Deprecated" column at the end
        result_df['New and Deprecated'] = merged_df['New and Deprecated']
        
        # Apply priority rules for highlighting
        if output_csv_path.lower().endswith('.xlsx'):
            output_path = output_csv_path
        else:
            # Change extension to xlsx if necessary
            output_path = os.path.splitext(output_csv_path)[0] + '.xlsx'
        
        # Save to Excel
        writer = pd.ExcelWriter(output_path, engine='openpyxl')
        result_df.to_excel(writer, index=False, sheet_name='Domains Highlight')
        writer.close()
        
        # Apply conditional formatting based on rules
        print(f"Applying rules to {output_path}:")
        for priority, rules_list in rules.items():
            print(f"  {priority} Priority: {len(rules_list)} rules")
            for rule in rules_list:
                print(f"    {rule['metric']} {rule['operator']} {rule['value']}")
                
        # Apply highlighting only if apply_formatting is True (for the two-phase approach)
        if apply_formatting:
            apply_highlighting_rules(output_path, result_df, rules)
        else:
            print("Skipping formatting in first phase - will be applied in separate formatting pass")
        
        print(f"Output saved to {output_path}")
        return output_path, result_df
            
    except Exception as e:
        print(f"Error processing files: {e}")
        raise
