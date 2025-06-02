"""
Error Distribution Module - Handles error distribution calculations
"""
import pandas as pd

def process_error_distribution(input_file):
    """
    Process error distribution data from an input CSV file.
    
    Parameters:
        input_file (str): Path to the input CSV file
        
    Returns:
        DataFrame: DataFrame with error distribution calculations
    """
    try:
        # Read CSV file
        df = pd.read_csv(input_file)
        
        # Validate required columns
        required_columns = ['Website/App Name', 'CSM Error', 'Type', 
                          'Website Ads Txt Reason', 'Ad Calls']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Calculate error distribution
        result_df = calculate_error_distribution(df)
        
        # Create summary data for the main summary sheet
        summary_data = create_summary_data(result_df)
        
        return result_df, summary_data
        
    except Exception as e:
        raise Exception(f"Error processing error distribution data: {str(e)}")

def calculate_error_distribution(df):
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

def create_summary_data(error_df):
    """
    Create summary data from error distribution results
    Returns a dictionary with Website/App Name as keys and (Status, Type, Percentage) as values
    """
    summary_data = {}
    
    # Group by Website/App Name
    for website, group in error_df.groupby('Website/App Name'):
        # Find the row with the most Ad Calls for each website
        max_adcalls_row = group.loc[group['Ad Calls'].idxmax()]
        
        # Extract the status (CSM Error), type, and percentage
        status = max_adcalls_row['CSM Error']
        error_type = max_adcalls_row['Type']
        percentage_str = max_adcalls_row['Error Distribution']
        
        # Store in the summary dictionary
        summary_data[website] = {
            'Most Adcalls status': status,
            'Most Ad Calls %': percentage_str,
            'Type': error_type  # Store the error type for conditional formatting
        }
    
    return summary_data
