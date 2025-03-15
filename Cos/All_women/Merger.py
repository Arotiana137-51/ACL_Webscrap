import pandas as pd

# Generate list of all file names
file_names = []

# Add numeric batches (1-4)
for i in range(1, 5):
    file_names.append(f"Cos_WOMEN_batch_{i}.xlsx")

# Add A batches (A1-A10)
for i in range(1, 11):
    file_names.append(f"Cos_WOMEN_batch_A{i}.xlsx")

# Add B batches (B1-B12)
for i in range(1, 13):
    file_names.append(f"Cos_WOMEN_batch_B{i}.xlsx")

# Add C batches (C1-C3)
for i in range(1, 4):
    file_names.append(f"Cos_WOMEN_batch_C{i}.xlsx")

# Add additional files
file_names.extend([
    "products_batch_Cos_final.xlsx",
    "cos_WOMEN_all.xlsx"
])

# Read and combine all Excel files
combined_df = pd.DataFrame()

for file in file_names:
    try:
        df = pd.read_excel(file)
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    except Exception as e:
        print(f"Error reading {file}: {str(e)}")

# Save the combined data
combined_df.to_excel("All_Cos_Women_wear.xlsx", index=False)

print("All files merged successfully into All_Cos_Women_wear.xlsx")