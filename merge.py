import pandas as pd

# Define the file paths for the original and converted Excel files
original_file = 'HPA organ specific gene list.xlsx'
converted_file = 'HPA_organ_specific_gene_list_converted.xlsx'

# Load both Excel files into pandas DataFrames
df_original = pd.read_excel(original_file, sheet_name=None)  # Load all sheets
df_converted = pd.read_excel(converted_file, sheet_name=None)  # Load all sheets

# Create a new Excel writer object
output_file = 'combined_gene_lists.xlsx'
with pd.ExcelWriter(output_file) as writer:
    # Iterate through each sheet in the original and converted DataFrames
    for sheet_name in df_original.keys():
        # Extract DataFrames for the current sheet from both original and converted files
        df_orig_sheet = df_original[sheet_name]
        df_conv_sheet = df_converted[sheet_name]
        
        # Initialize a new DataFrame to store the combined data
        df_combined = pd.DataFrame()
        
        # Iterate through each column in the original sheet
        for col_name in df_orig_sheet.columns:
            # Create new columns in the combined DataFrame
            df_combined[f'{col_name}_Gene_Symbol'] = df_orig_sheet[col_name]
            df_combined[f'{col_name}_BGEE'] = df_conv_sheet[col_name]
        
        # Write the combined DataFrame to the Excel file, under the current sheet name
        df_combined.to_excel(writer, sheet_name=sheet_name, index=False)

print(f'Combined file saved successfully as {output_file}')

