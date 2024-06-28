import pandas as pd
import requests
import time
from gprofiler import GProfiler

# Initialize g:Profiler
gp = GProfiler(return_dataframe=True)

def convert_gene_symbols_to_bgee(gene_symbols, max_retries=5, wait_time=10):
    if not gene_symbols:
        return []

    attempt = 0
    while attempt < max_retries:
        try:
            # Convert gene symbols to Ensembl Gene IDs (ENSG) using g:Profiler
            result = gp.convert(organism='hsapiens', query=gene_symbols, target_namespace='ENSG')
            print(result)  # Debug print to see the conversion results

            if result.empty:
                return gene_symbols  # If conversion fails, return original gene symbols

            # Create a mapping from original gene symbols to ENSG IDs
            gene_to_ensg = dict(zip(result['incoming'], result['converted']))
            return [gene_to_ensg.get(gene, gene) for gene in gene_symbols]
        except requests.exceptions.RequestException as e:
            print(f"Request failed: {e}, retrying in {wait_time} seconds...")
            time.sleep(wait_time)
            attempt += 1

    # If all retries fail, return the original gene symbols
    print("Max retries reached, returning original gene symbols.")
    return gene_symbols

# Load the HPA organ specific gene list Excel file
input_file_path = './HPA organ specific gene list.xlsx'
output_file_path = './HPA_organ_specific_gene_list_converted.xlsx'
hpa_df = pd.read_excel(input_file_path, sheet_name=None)

# Create a new dictionary to hold the converted DataFrames
converted_dfs = {}

# Process each sheet
for sheet_name, df in hpa_df.items():
    print(f"Processing sheet: {sheet_name}")
    converted_df = df.copy()

    # Convert each column
    for column in df.columns:
        gene_symbols = df[column].dropna().tolist()
        converted_bgee_ids = convert_gene_symbols_to_bgee(gene_symbols)
        
        # Replace the column with converted Bgee IDs
        converted_df[column] = pd.Series(converted_bgee_ids, index=df[column].dropna().index)

    # Store the converted DataFrame
    converted_dfs[sheet_name] = converted_df

# Save the converted DataFrames to a new Excel file
with pd.ExcelWriter(output_file_path) as writer:
    for sheet_name, converted_df in converted_dfs.items():
        converted_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Conversion completed and saved to: {output_file_path}")

