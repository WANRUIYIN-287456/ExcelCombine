import pandas as pd
from pathlib import Path
import openpyxl

def process_file(file):
    try:
        # Skip files starting with ~$
        if file.name.startswith("~$"):
            print(f"Skipping temporary file: {file}")
            return None, 0  # Return 0 rows for skipped files
        
        # Read the Excel file
        if file.suffix == ".xls":
            df = pd.read_excel(file, engine="xlrd")  # Use xlrd for .xls files
        elif file.suffix == ".xlsx":
            df = pd.read_excel(file, engine="openpyxl")  # Use openpyxl for .xlsx files
        else:
            print(f"Skipping {file}: Unsupported file type.")
            return None, 0  # Return 0 rows for unsupported files

        # Trim column names to avoid space issues
        df.columns = df.columns.str.strip()

        total_rows = 0  # Track the number of rows processed for the current file

        # Process Type 1 File
        if {'First Name', 'Last Name', 'Email Address', 'Cell Phone'}.issubset(df.columns):
            # Convert to string explicitly before concatenation
            df['First Name'] = df['First Name'].fillna('').astype(str)
            df['Last Name'] = df['Last Name'].fillna('').astype(str)
            
            df['Name'] = df['First Name'] + " " + df['Last Name']
            df['Email'] = df['Email Address'].fillna('').astype(str)
            df['Phone Number'] = df['Cell Phone'].fillna('').astype(str)  # Convert to string explicitly
            
            processed_df = df[['Name', 'Email', 'Phone Number']]
            total_rows = len(processed_df)
            print(f"Copied {total_rows} rows from {file}")

        # Process Type 2 File
        elif {'Buyer Name', 'Buyer Email', 'Buyer Contact', 'FULL NAME', 'E-MAIL ADDRESS', 'MOBILE NUMBER'}.issubset(df.columns):
            # Rows for Buyer Name, Buyer Email, Buyer Contact
            buyer_data = df[['Buyer Name', 'Buyer Email', 'Buyer Contact']].fillna('')
            buyer_data.columns = ['Name', 'Email', 'Phone Number']
            buyer_data['Phone Number'] = buyer_data['Phone Number'].astype(str)  # Ensure 'Phone Number' is text

            # Rows for FULL NAME, E-MAIL ADDRESS, MOBILE NUMBER
            full_name_data = df[['FULL NAME', 'E-MAIL ADDRESS', 'MOBILE NUMBER']].fillna('')
            full_name_data.columns = ['Name', 'Email', 'Phone Number']
            full_name_data['Phone Number'] = full_name_data['Phone Number'].astype(str)  # Ensure 'Phone Number' is text

            # Combine both into a single DataFrame
            processed_df = pd.concat([buyer_data, full_name_data], ignore_index=True)

            # Count rows for both sets of data (Buyer Data + Full Name Data)
            total_rows = len(buyer_data) + len(full_name_data)
            print(f"Copied {total_rows} rows from {file}")

        else:
            print(f"Skipping {file}: Missing required columns.")
            print(f"Columns in {file}: {df.columns}")  # Log columns of the skipped file
            return None, 0  # Return 0 rows for skipped files

        return processed_df, total_rows

    except Exception as e:
        print(f"Error processing {file}: {e}")
        return None, 0  # Return 0 rows in case of error

def combine_excel_files(input_files, output_file):
    combined_data = []
    total_rows_combined = 0  # Initialize the total rows counter

    for file in input_files:
        processed_df, rows = process_file(file)
        if processed_df is not None:
            combined_data.append(processed_df)
            total_rows_combined += rows  # Accumulate total rows copied

    # Combine all data into a single DataFrame
    if combined_data:
        final_df = pd.concat(combined_data, ignore_index=True)

        # Save to a new Excel file using openpyxl
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False)

            # Get the openpyxl workbook and sheet
            workbook = writer.book
            worksheet = workbook.active

            # Format the 'Phone Number' column as text (in Excel, 'Text' format will preserve leading zeros)
            phone_number_column = final_df.columns.get_loc("Phone Number") + 1  # Excel columns are 1-indexed
            for row in worksheet.iter_rows(min_col=phone_number_column, max_col=phone_number_column, min_row=2, max_row=worksheet.max_row):
                for cell in row:
                    # Set the cell format to text
                    cell.number_format = '@'  # '@' is the format code for text in Excel

        print(f"Combined file saved as {output_file}")
        print(f"Total rows copied: {total_rows_combined}")  # Print total rows copied
    else:
        print("No valid files to process.")

# Input Excel files (update the path to your folder)
input_folder = Path(r"C:\Users\User\Downloads\seminar list\seminar list\Ipoh")
input_files = list(input_folder.glob("*.xls")) + list(input_folder.glob("*.xlsx"))  # Convert to list before concatenating

# Output file name
output_file = input_folder / "combined_output.xlsx"

# Combine the files
combine_excel_files(input_files, output_file)
