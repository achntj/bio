import pandas as pd

# Load the Excel file
file_path = './Literature_Organ_gene_panel_Uni+HPA.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Mappings between enriched and organ columns
organ_mapping = {
    'Urinary Bladder': 'bladder',
    'Bone marrow': ['bone', 'bone marrow'],
    'Breast': 'breast',
    'Stomach': 'gi',
    'Kidney': 'kidney',
    'Liver': 'liver',
    'Lung': 'lung',
    'Ovary': 'ovary',
    'Pancreas': 'pancreas',
    'Skin': 'skin',
    'Intestine': ['small intestine', 'crc'],
    'Testis': 'testis',
    'Thyroid Gland': 'thyroid'
}

# Function to get the column names for a given organ
def get_column_names(organ):
    high_gs_col = f'HIGH (GS) {organ}'
    high_bgee_col = f'HIGH (Bgee) {organ}'
    low_gs_col = f'LOW (GS) {organ}'
    low_bgee_col = f'LOW (Bgee) {organ}'
    ind_gs_col = f'INDETERMINATE (GS) {organ}'
    ind_bgee_col = f'INDETERMINATE (Bgee) {organ}'
    return [high_bgee_col, low_bgee_col, ind_bgee_col]

# Function to get the enriched column names for a given organ
def get_enriched_column_names(enriched_organ):
    enriched_gs_col = f'Enriched HPA {enriched_organ} (GS)'
    enriched_bgee_col = f'Enriched HPA {enriched_organ} (Bgee)'
    return enriched_bgee_col

# Insert columns after the specified columns
def insert_columns(df, new_cols, insert_after_col):
    for new_col in new_cols:
        df.insert(df.columns.get_loc(insert_after_col) + 1, new_col, '')

# Process each organ based on the mapping
for enriched_organ, organ_key in organ_mapping.items():
    enriched_bgee_col = get_enriched_column_names(enriched_organ)
    literature_col = f'Literature {enriched_organ} (Bgee)'
    
    # Multiple organ columns for some organs
    if isinstance(organ_key, list):
        for sub_organ in organ_key:
            columns_to_process = get_column_names(sub_organ)
            bgee_sum_col = f'{sub_organ} (bgee)'
            expression_col = f'Expression_{sub_organ}'
            
            # Insert new columns at the correct position
            insert_columns(df, [expression_col, bgee_sum_col], columns_to_process[-1])
            
            for idx, row in df.iterrows():
                bgee_values = set()  # Use set to store unique Bgee values
                expression_types = set()  # Use set to store unique expression types
                
                for col in columns_to_process:
                    if col in df.columns and pd.notna(row[col]) and row[col] != '':
                        bgee_values.add(str(row[col]))
                        expression_types.add(col.split()[0].lower())
                
                if enriched_bgee_col in df.columns and pd.notna(row[enriched_bgee_col]) and row[enriched_bgee_col] != '':
                    bgee_values.add(str(row[enriched_bgee_col]))
                    expression_types.add('enriched')
                
                if literature_col in df.columns and pd.notna(row[literature_col]) and row[literature_col] != '':
                    bgee_values.add(str(row[literature_col]))
                    expression_types.add('literature')
                
                if bgee_values:
                    df.at[idx, bgee_sum_col] = ' '.join(bgee_values)
                    df.at[idx, expression_col] = '/'.join(expression_types)
    
    else:
        columns_to_process = get_column_names(organ_key)
        bgee_sum_col = f'{organ_key} (bgee)'
        expression_col = f'Expression_{organ_key}'
        
        # Insert new columns at the correct position
        insert_columns(df, [expression_col, bgee_sum_col], columns_to_process[-1])
        
        for idx, row in df.iterrows():
            bgee_values = set()  # Use set to store unique Bgee values
            expression_types = set()  # Use set to store unique expression types
            
            for col in columns_to_process:
                if col in df.columns and pd.notna(row[col]) and row[col] != '':
                    bgee_values.add(str(row[col]))
                    expression_types.add(col.split()[0].lower())
            
            if enriched_bgee_col in df.columns and pd.notna(row[enriched_bgee_col]) and row[enriched_bgee_col] != '':
                bgee_values.add(str(row[enriched_bgee_col]))
                expression_types.add('enriched')
            
            if literature_col in df.columns and pd.notna(row[literature_col]) and row[literature_col] != '':
                bgee_values.add(str(row[literature_col]))
                expression_types.add('literature')
            
            if bgee_values:
                df.at[idx, bgee_sum_col] = ' '.join(bgee_values)
                df.at[idx, expression_col] = '/'.join(expression_types)

# Save the updated DataFrame to a new Excel file
output_file_path = './Final_Literature_Organ_gene_panel_Uni+HPA_NoDuplicates.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Processing completed and saved to: {output_file_path}")

