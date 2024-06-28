import pandas as pd

# Load the updated organ gene panel file
updated_file_path = './Organ gene panel Uni+HPA.xlsx'
updated_df = pd.read_excel(updated_file_path, sheet_name='Sheet1')

# Load the organ table 2 gene list file
organ_table_file_path = './organ table 2 gene list_2.xlsx'
organ_table_df = pd.read_excel(organ_table_file_path, sheet_name='Sheet1')

# Process each organ column in the organ table
for organ_name in organ_table_df.columns:
    new_col_name = f'Literature {organ_name} (Bgee)'
    updated_df[new_col_name] = ''  # Initialize the new column with empty strings
    
    for idx, bgee_value in organ_table_df[organ_name].dropna().items():
        if bgee_value in updated_df['Bgee'].values:
            # Get the row index in the updated_df where the Bgee value matches
            row_index = updated_df[updated_df['Bgee'] == bgee_value].index[0]
            # Set the Bgee value in the new column
            updated_df.at[row_index, new_col_name] = bgee_value

# Save the updated DataFrame to a new Excel file
output_file_path = './Literature_Organ_gene_panel_Uni+HPA.xlsx'
updated_df.to_excel(output_file_path, index=False)

print(f"Processing completed and saved to: {output_file_path}")

