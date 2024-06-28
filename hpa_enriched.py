import pandas as pd

# File paths
hpa_scoring_file = 'updated_hpa_scoring.xlsx'
combined_gene_lists_file = 'combined_gene_lists.xlsx'

# Load Excel files into pandas dataframes
hpa_df = pd.read_excel(hpa_scoring_file, sheet_name='Sheet1')  # Assuming only one sheet
combined_df = pd.read_excel(combined_gene_lists_file, sheet_name='enriched')

# Get all organ columns from combined_df
organ_columns = [col for col in combined_df.columns if '_Gene_Symbol' in col]

# Create new enriched columns in hpa_df for each organ
for organ_col in organ_columns:
    organ_name = organ_col.split('_')[0]
    hpa_df[f'Enriched HPA {organ_name} (GS)'] = ''
    hpa_df[f'Enriched HPA {organ_name} (Bgee)'] = ''

# Process each organ column and update hpa_df
for organ_col in organ_columns:
    organ_name = organ_col.split('_')[0]
    symbol_col = f'{organ_name}_Gene_Symbol'
    bgee_col = f'{organ_name}_BGEE'
    
    print(f"Processing {symbol_col} and {bgee_col}...")

    for index, row in combined_df.iterrows():
        organ_symbol = row[symbol_col]
        organ_bgee = row[bgee_col]
        
        # Find matches in hpa_df and update enriched columns
        mask = hpa_df['Bgee'].apply(lambda x: x == organ_bgee if pd.notnull(x) and pd.notnull(organ_bgee) else False)
        hpa_df.loc[mask, f'Enriched HPA {organ_name} (GS)'] = organ_symbol
        hpa_df.loc[mask, f'Enriched HPA {organ_name} (Bgee)'] = organ_bgee
    
    print(f"Processed {symbol_col} and {bgee_col}.")

# Write the updated DataFrame to Excel
output_file = 'updated_hpa_scoring_with_enriched_organs.xlsx'
hpa_df.to_excel(output_file, index=False, engine='openpyxl')

print(f"Process complete. Output saved to {output_file}.")

