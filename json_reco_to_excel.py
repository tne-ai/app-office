import pandas as pd
from io import StringIO
import json

def markdown_tables_from_json_to_single_sheet(json_file, output_excel):
    """
    Extract Markdown tables from a JSON file and write them to a single sheet in an Excel file.

    :param json_file: Path to the JSON file containing Markdown tables.
    :param output_excel: Path to the output Excel file.
    """
    try:
        # Load the JSON file
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Navigate to the tables in the JSON
        tables = data.get("reco", {}).get("supporting_tables", [])
        
        if not tables:
            raise ValueError("No tables found in the JSON file.")

        # Initialize a list to store all data for the single sheet
        combined_data = []

        # Process each Markdown table
        for idx, table in enumerate(tables):
            # Extract the table header and body
            table_header = table.get("supporting_table_header", f"Table {idx + 1}")
            markdown_table = table.get("supporting_table_body", "")

            # Add the table header as a new row
            combined_data.append([table_header])

            # Remove Markdown alignment characters
            table_csv = markdown_table.replace('|', '').strip()
            table_csv = table_csv.replace(':---', '').replace('---:', '').replace(':---:', '')

            # Create a DataFrame from the cleaned table
            df = pd.read_csv(StringIO(table_csv), sep=r'\s{2,}', engine='python')  # Two or more spaces as delimiter

            # Append DataFrame content as a list of rows
            combined_data.extend(df.values.tolist())

            # Add a blank row as a separator between tables
            combined_data.append([])

        # Create a final DataFrame
        combined_df = pd.DataFrame(combined_data)

        # Write the combined data to a single Excel sheet
        combined_df.to_excel(output_excel, index=False, header=False)

        print(f"Successfully wrote all tables to a single sheet in '{output_excel}'.")
    
    except Exception as e:
        print(f"Error: {e}")

# Example usage
json_file = "data/newness-reco.json"  # Replace with your JSON file path
output_excel = "xlsx/newness-reco2.xlsx"  # Replace with your output Excel file path
markdown_tables_from_json_to_single_sheet(json_file, output_excel)