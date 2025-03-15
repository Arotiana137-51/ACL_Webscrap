import pandas as pd

def update_column_z_woven():
    # Load the merged Excel file
    df = pd.read_excel("All_Cos_Women_wear.xlsx")
    
    # Define column positions (0-based index)
    col_f_position = 5  # Excel column F
    col_z_position = 25  # Excel column Z
    
    # Ensure we have enough columns
    while len(df.columns) <= col_z_position:
        df[f'Column_{len(df.columns)+1}'] = None
    
    # Get column names by position
    col_f = df.columns[col_f_position]
    col_z = df.columns[col_z_position]
    
    # Convert to string type
    df[col_f] = df[col_f].astype(str)
    
    # Create boolean mask for cells containing 'woven'
    woven_mask = df[col_f].str.contains(
        'woven', 
        case=False, 
        na=False
    )
    
    # Update column Z with "woven" text where mask is True
    df.loc[woven_mask, col_z] = "woven"
    
    # Save the modified DataFrame
    df.to_excel("All_Cos_Women_wear.xlsx", index=False)
    print("Column Z updated successfully with 'woven' markers")

if __name__ == "__main__":
    update_column_z_woven()