import pandas as pd

# Load the Excel file
file_path = './Final_Literature_Organ_gene_panel_Uni+HPA_NoDuplicates.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# List of organs to process
organs_to_process = ['Brain', 'Prostate', 'Lymphoid tissue']

# Function to get the enriched column name for a given organ
def get_enriched_column_name(organ):
    return f'Enriched HPA {organ} (Bgee)'

# Function to generate the two columns for each organ
def generate_columns_for_organ(organ):
    enriched_bgee_col = get_enriched_column_name(organ)
    literature_col = f'Literature {organ} (Bgee)'
    bgee_sum_col = f'{organ} (bgee)'
    expression_col = f'Expression_{organ}'
    
    # Initialize columns
    df[bgee_sum_col] = ''
    df[expression_col] = ''
    
    # Iterate through rows
    for idx, row in df.iterrows():
        bgee_values = set()  # Use set to store unique Bgee values
        expression_types = []  # List to store expression types
        
        # Process enriched column
        if enriched_bgee_col in df.columns and pd.notna(row[enriched_bgee_col]) and row[enriched_bgee_col] != '':
            bgee_values.add(str(row[enriched_bgee_col]))
            expression_types.append('enriched')
        
        # Process literature column
        if literature_col in df.columns and pd.notna(row[literature_col]) and row[literature_col] != '':
            bgee_values.add(str(row[literature_col]))
            expression_types.append('literature')
        
        # Update DataFrame with concatenated values
        if bgee_values:
            df.at[idx, bgee_sum_col] = ' '.join(bgee_values)
            df.at[idx, expression_col] = '/'.join(expression_types)

# Generate columns for each organ
for organ in organs_to_process:
    generate_columns_for_organ(organ)

# Save the updated DataFrame to the same Excel file
output_file_path = './final.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Columns generated for organs: {', '.join(organs_to_process)}")
print(f"Updated file saved to: {output_file_path}")

