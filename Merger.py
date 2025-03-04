import pandas as pd
import os

# Current directory
directory = '.'

# Create a list to store the dataframes
dfs = []

# Iterate over each batch file (from 1 to 10)
for i in range(1, 11):
    filename = f'Huckberry_Men_Shirts_batch_{i}.xlsx'
    if os.path.isfile(os.path.join(directory, filename)):
        # Read the file
        df = pd.read_excel(os.path.join(directory, filename))
        # Append the dataframe to the list
        dfs.append(df)

# Concatenate all dataframes
merged_df = pd.concat(dfs, ignore_index=False)

# Save the merged dataframe to a new Excel file with the index
merged_df.to_excel('Huckberry_Men_Shirts_Final.xlsx', index=True)
