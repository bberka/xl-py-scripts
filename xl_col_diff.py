import argparse
import asyncio
import os
import re
from xl_helper import read_excel, read_rows
import openpyxl


# Function to compare and sync columns between two files
async def compare_and_sync_columns(old_file, new_file, allow_delete=False, ignore_sheet_regex=None):
    old_sheet_data = await read_excel(old_file)
    new_sheet_data = await read_excel(new_file)

    # Iterate through the sheet names in both files
    for sheet_name in new_sheet_data:  # Check for all sheets in the new file
        if ignore_sheet_regex and re.match(ignore_sheet_regex, sheet_name):
            print(f"Ignoring sheet: {sheet_name} (matches ignore regex)")
            continue  # Skip this sheet if it matches the ignore regex

        if sheet_name in old_sheet_data:  # Check if the sheet exists in the old file
            print(f"Comparing columns in sheet: {sheet_name}")

            # Read only the header (first row) of the old and new sheet
            old_rows = await read_rows(old_file, sheet_name, read_header_only=True)
            new_rows = await read_rows(new_file, sheet_name, read_header_only=True)

            # Get the list of headers (columns) in both old and new sheet
            old_headers = old_rows[0] if old_rows else []
            new_headers = new_rows[0] if new_rows else []

            # Compare headers and add missing columns to the old file
            for new_header in new_headers:
                if new_header not in old_headers:
                    print(f"Column '{new_header}' is missing in the old file, adding it.")
                    # Add missing column to the old header row
                    old_headers.append(new_header)
                    
                    # Add the new column to all rows of the old sheet (initializing with empty values)
                    wb_old = openpyxl.load_workbook(old_file)
                    sheet_old = wb_old[sheet_name]

                    for row_idx in range(len(old_rows)):
                        sheet_old.cell(row=row_idx + 1, column=len(old_headers), value=None)  # Add new empty column

                    wb_old.save(old_file)

            # After ensuring the headers match, update the data (rows) accordingly
            # Now that we have the updated old_headers, update the rows in the old file
            for row_idx, row in enumerate(old_rows):
                for col_idx, new_header in enumerate(new_headers):
                    if new_header not in old_headers:
                        # Add the new value (from new file) to the corresponding row and column in the old sheet
                        sheet_old.cell(row=row_idx + 1, column=len(old_headers), value=new_rows[row_idx][col_idx])

            print(f"Columns synchronized for sheet: {sheet_name}")

        else:
            print(f"Sheet '{sheet_name}' does not exist in the old file. Skipping...")

    print("Comparison and sync completed.")


# Function to traverse directories and compare files
def compare_directory_files(old_dir, new_dir, allow_delete=False, ignore_file_regex=None, ignore_sheet_regex=None):
    for root, dirs, files in os.walk(new_dir):
        for file_name in files:
            if ignore_file_regex and re.match(ignore_file_regex, file_name):
                print(f"Ignoring file: {file_name} (matches ignore regex)")
                continue  # Skip this file if it matches the ignore regex

            if file_name.endswith(('.xlsx', '.xls')):  # Check if it's an Excel file
                # Get relative file path
                relative_path = os.path.relpath(os.path.join(root, file_name), new_dir)
                old_file_path = os.path.join(old_dir, relative_path)

                # Check if the file exists in both directories
                if os.path.exists(old_file_path):
                    print(f"Comparing: {old_file_path} and {os.path.join(root, file_name)}")
                    asyncio.run(compare_and_sync_columns(old_file_path, os.path.join(root, file_name), allow_delete, ignore_sheet_regex))


# Main function to parse arguments and execute the script
def main():
    parser = argparse.ArgumentParser(description="Compare and sync columns in two Excel files or directories.")
    parser.add_argument('--old-file', required=True, help="Path to the old Excel file or directory.")
    parser.add_argument('--new-file', required=True, help="Path to the new Excel file or directory.")
    parser.add_argument('--check-directory', action='store_true', help="Specify if the paths are directories to compare Excel files inside.")
    parser.add_argument('--allow-delete', action='store_true', help="Allow deletion of columns in the old file.")
    parser.add_argument('--ignore-sheet-regex', type=str, help="Regex pattern to ignore sheets during comparison.")
    parser.add_argument('--ignore-file-regex', type=str, help="Regex pattern to ignore files during comparison.")

    args = parser.parse_args()

    old_file = args.old_file
    new_file = args.new_file
    allow_delete = args.allow_delete
    check_directory = args.check_directory
    ignore_sheet_regex = args.ignore_sheet_regex
    ignore_file_regex = args.ignore_file_regex

    # If the paths are directories, compare all files inside them
    if check_directory:
        if not os.path.isdir(old_file) or not os.path.isdir(new_file):
            print("Both paths must be directories when using --check-directory.")
            return

        compare_directory_files(old_file, new_file, allow_delete, ignore_file_regex, ignore_sheet_regex)
    else:
        # Otherwise, compare the single files directly
        if not os.path.exists(old_file) or not os.path.exists(new_file):
            print("One or both of the provided files do not exist.")
            return

        asyncio.run(compare_and_sync_columns(old_file, new_file, allow_delete, ignore_sheet_regex))


if __name__ == "__main__":
    main()
