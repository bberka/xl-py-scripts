import argparse
import asyncio
import os
import openpyxl
from xl_helper import read_excel, read_rows


# Function to compare and sync columns between two files
async def compare_and_sync_columns(old_file, new_file, allow_delete=False):
    print(f"Comparing and syncing columns between {old_file} and {new_file}")
    old_sheet_data = await read_excel(old_file)
    new_sheet_data = await read_excel(new_file)

    # Iterate through the sheet names in both files
    for sheet_name in new_sheet_data:  # Check for all sheets in the new file
        if sheet_name in old_sheet_data:  # Check if the sheet exists in the old file
            print(f"Comparing columns in sheet: {sheet_name}")

            # Read the rows of the old and new sheet
            old_rows = await read_rows(old_file, sheet_name)
            new_rows = await read_rows(new_file, sheet_name)

            # Check column lengths and add missing columns in the old file
            old_column_count = len(old_rows[0]) if old_rows else 0
            new_column_count = len(new_rows[0]) if new_rows else 0

            if new_column_count > old_column_count:
                print(f"Adding {new_column_count - old_column_count} columns to {sheet_name}")
                
                for row_idx in range(len(old_rows)):
                    if allow_delete:
                        # Append new columns at the end if allow_delete is True
                        old_rows[row_idx] += new_rows[row_idx][old_column_count:new_column_count]
                    else:
                        # Insert new columns at the correct position (shift columns to the right)
                        for col_idx in range(old_column_count, new_column_count):
                            old_rows[row_idx].insert(col_idx, new_rows[row_idx][col_idx])

                # Update the old file with the new columns
                wb_old = openpyxl.load_workbook(old_file)
                sheet_old = wb_old[sheet_name]

                for row_idx, row in enumerate(old_rows):
                    for col_idx, value in enumerate(row):
                        sheet_old.cell(row=row_idx + 1, column=col_idx + 1, value=value)

                wb_old.save(old_file)
                print(f"Columns added to {sheet_name} in the old file")

            # If deletions are allowed, check if columns in the old file don't exist in the new file
            if allow_delete and new_column_count < old_column_count:
                columns_to_delete = old_column_count - new_column_count
                print(f"Deleting {columns_to_delete} columns from {sheet_name} in the old file")

                wb_old = openpyxl.load_workbook(old_file)
                sheet_old = wb_old[sheet_name]

                for col in range(old_column_count, new_column_count, -1):
                    sheet_old.delete_cols(col)

                wb_old.save(old_file)
                print(f"Columns deleted from {sheet_name} in the old file")

        else:
            print(f"Sheet '{sheet_name}' does not exist in the old file. Skipping...")

    print("Comparison and sync completed.")


# Function to traverse directories and compare files
def compare_directory_files(old_dir, new_dir, allow_delete=False):
    for root, dirs, files in os.walk(new_dir):
        for file_name in files:
            if file_name.endswith(('.xlsx', '.xls')):  # Check if it's an Excel file
                # Get relative file path
                relative_path = os.path.relpath(os.path.join(root, file_name), new_dir)
                old_file_path = os.path.join(old_dir, relative_path)

                # Check if the file exists in both directories
                if os.path.exists(old_file_path):
                    print(f"Comparing: {old_file_path} and {os.path.join(root, file_name)}")
                    asyncio.run(compare_and_sync_columns(old_file_path, os.path.join(root, file_name), allow_delete))


# Main function to parse arguments and execute the script
def main():
    parser = argparse.ArgumentParser(description="Compare and sync columns in two Excel files or directories.")
    parser.add_argument('--old-file', required=True, help="Path to the old Excel file or directory.")
    parser.add_argument('--new-file', required=True, help="Path to the new Excel file or directory.")
    parser.add_argument('--check-directory', action='store_true', help="Specify if the paths are directories to compare Excel files inside.")
    parser.add_argument('--allow-delete', action='store_true', help="Allow deletion of columns in the old file.")

    args = parser.parse_args()

    old_file = args.old_file
    new_file = args.new_file
    allow_delete = args.allow_delete
    check_directory = args.check_directory

    # If the paths are directories, compare all files inside them
    if check_directory:
        if not os.path.isdir(old_file) or not os.path.isdir(new_file):
            print("Both paths must be directories when using --check-directory.")
            return

        compare_directory_files(old_file, new_file, allow_delete)
    else:
        # Otherwise, compare the single files directly
        if not os.path.exists(old_file) or not os.path.exists(new_file):
            print("One or both of the provided files do not exist.")
            return

        asyncio.run(compare_and_sync_columns(old_file, new_file, allow_delete))


if __name__ == "__main__":
    main()
