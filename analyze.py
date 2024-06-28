import pandas as pd
import os

def process_organ_data(organ_name, organ_df, uniprot_df):
    # Define the sheets we are interested in
    sheet_pairs = [('High-C', 'High'), ('Low-C', 'Low'), ('NS-C', 'NS')]
    column_suffixes = ['HIGH', 'LOW', 'INDETERMINATE']

    # Create new columns for the organ in the UniProt DataFrame
    for suffix in column_suffixes:
        uniprot_df[f'{suffix} (GS) {organ_name}'] = ''
        uniprot_df[f'{suffix} (Bgee) {organ_name}'] = ''

    for suffix, (primary_sheet, fallback_sheet) in zip(column_suffixes, sheet_pairs):
        # Try to get the primary sheet first, if it doesn't exist, use the fallback sheet
        if primary_sheet in organ_df:
            sheet_df = organ_df[primary_sheet]
        elif fallback_sheet in organ_df:
            sheet_df = organ_df[fallback_sheet]
        else:
            continue  # If neither sheet exists, skip to the next suffix

        # Ensure column names are strings and strip any leading/trailing spaces
        sheet_df.columns = sheet_df.columns.astype(str).str.strip()

        if 'Gene Names' not in sheet_df.columns or 'Bgee' not in sheet_df.columns:
            print(f"Warning: Expected columns not found in sheet '{primary_sheet}' or '{fallback_sheet}' for organ '{organ_name}'.")
            continue

        sheet_genes = sheet_df['Gene Names'].tolist()
        sheet_bgees = sheet_df['Bgee'].tolist()

        for index, row in uniprot_df.iterrows():
            gene_symbols = row['Gene symbols(GS)']
            bgee = row['Bgee']

            # Check for matches in the sheet's Bgee column
            if bgee in sheet_bgees:
                matched_index = sheet_bgees.index(bgee)
                uniprot_df.at[index, f'{suffix} (Bgee) {organ_name}'] = bgee
                uniprot_df.at[index, f'{suffix} (GS) {organ_name}'] = sheet_genes[matched_index]

            # Check for matches in the sheet's Gene Names column
            for gene in gene_symbols.split():
                if gene in sheet_genes:
                    matched_index = sheet_genes.index(gene)
                    uniprot_df.at[index, f'{suffix} (Bgee) {organ_name}'] = bgee
                    uniprot_df.at[index, f'{suffix} (GS) {organ_name}'] = sheet_genes[matched_index]
                    break

    return uniprot_df

# Load the UniProt HPA scoring file
uniprot_file_path = './hpa scoring/uniprot hpa scoring.xlsx'
uniprot_df = pd.read_excel(uniprot_file_path)

# Directory containing the organ files
organ_files_directory = './uniprot 15'
organ_files = [f for f in os.listdir(organ_files_directory) if f.endswith('.xlsx')]

# Process each organ file
for organ_file in organ_files:
    print("Processing", organ_file)
    organ_name = os.path.splitext(organ_file)[0]  # Get the organ name from the file name
    organ_file_path = os.path.join(organ_files_directory, organ_file)
    organ_df = pd.read_excel(organ_file_path, sheet_name=None)  # Load all sheets

    # Process the organ data
    uniprot_df = process_organ_data(organ_name, organ_df, uniprot_df)

# Save the updated UniProt DataFrame to a new Excel file
output_file_path = './updated_hpa_scoring.xlsx'
uniprot_df.to_excel(output_file_path, index=False)

print("Processing completed and file saved to:", output_file_path)

