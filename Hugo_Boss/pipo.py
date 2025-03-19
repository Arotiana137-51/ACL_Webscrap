import pandas as pd
import os

def process_excel_file(file_name="HugoBoss.xlsx"):
    """
    Process an Excel file according to specified requirements.
    
    Args:
        file_name (str): Name of the Excel file in the current directory
    
    Returns:
        str: Path to the processed file
    """
    # Current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Input and output paths
    input_path = os.path.join(current_dir, file_name)
    output_path = os.path.join(current_dir, f"HugoBoss_processed.xlsx")
    
    # Read the Excel file
    print(f"Reading file: {input_path}")
    df = pd.read_excel(input_path)
    
    # Step 1 & 2: Check column AD for specific fit values
    fit_patterns = ["Slim fit", "Regular fit", "Relaxed fit"]
    
    # Create a mask for rows that have these fit values in column AD
    mask = df["AD"].isin(fit_patterns)
    
    # Store the original fit values in column Q for matching rows
    df.loc[mask, "Q"] = df.loc[mask, "AD"]
    
    # Replace with "Normal sale" in column AD
    df.loc[mask, "AD"] = "Normal sale"
    
    print(f"Processed {mask.sum()} rows with fit values")
    
    # Step 3: Fill None values in column AB with values from AE where AE doesn't contain "Sale"
    # Create a mask for rows where AB is None/NaN and AE doesn't contain "Sale-"
    ab_none_mask = df["AB"].isna() | (df["AB"] == "None") | (df["AB"] == "")
    ae_no_sale_mask = ~df["AE"].str.contains("Sale-", na=False)
    fill_mask = ab_none_mask & ae_no_sale_mask
    
    # Fill the values
    df.loc[fill_mask, "AB"] = df.loc[fill_mask, "AE"]
    
    print(f"Filled {fill_mask.sum()} empty AB values with AE values")
    
    # Save the processed file
    df.to_excel(output_path, index=False)
    print(f"Saved processed file to: {output_path}")
    
    return output_path

if __name__ == "__main__":
    # Process the file
    processed_file = process_excel_file()
    print(f"Processing complete. File saved at: {processed_file}")