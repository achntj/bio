import pandas as pd
import requests

# Read the updated Excel file
file_path = './Updated_Organ_gene_panel_Uni+HPA.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Function to convert Bgee IDs to Gene Symbols using g:Profiler API
def convert_bgee_to_gene_symbols(bgee_ids):
    url = "https://biit.cs.ut.ee/gprofiler/api/gconvert"
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Python client"
    }
    payload = {
        "organism": "hsapiens",
        "query": bgee_ids,
        "target": "ENSG",  # Assuming target is Ensembl Gene ID, adjust if needed
        "format": "json"
    }
    response = requests.post(url, json=payload, headers=headers)
    response_data = response.json()
    
    # Create a dictionary for mapping Bgee IDs to Gene Symbols
    conversion_dict = {item['incoming']: item['name'] for item in response_data['result']}
    return conversion_dict

# Get unique Bgee IDs from the Bgee column
unique_bgee_ids = df['Bgee'].dropna().unique().tolist()

# Convert Bgee IDs to Gene Symbols
conversion_dict = convert_bgee_to_gene_symbols(unique_bgee_ids)

# Replace the "Gene symbols(GS)" column with the converted Gene Symbols
df['Gene symbols(GS)'] = df['Bgee'].map(conversion_dict)

# Save the updated DataFrame to a new Excel file
output_file_path = './Final_Organ_gene_panel_Uni+HPA.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Gene symbol conversion completed and saved to: {output_file_path}")

