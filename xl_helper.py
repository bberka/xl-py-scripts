import asyncio
import openpyxl
import aiofiles


async def read_excel(file_path):
    """
    Asynchronously reads an Excel file and returns the sheet names.
    :param file_path: str, path to the Excel file.
    :return: list of sheet names.
    """
    try:
        # Open the Excel file asynchronously
        async with aiofiles.open(file_path, mode='rb') as f:
            content = await f.read()

        # Write the content into a temporary file that openpyxl can read
        with open(file_path, 'wb') as temp_file:
            temp_file.write(content)

        # Load the Excel workbook using openpyxl
        wb = openpyxl.load_workbook(file_path)

        # Return sheet names
        return wb.sheetnames

    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return []


async def read_rows(file_path: str, sheet_name: str, read_header_only=False):
    """
    Reads rows from an Excel sheet, optionally only reading the first row.
    
    :param file_path: Path to the Excel file.
    :param sheet_name: Name of the sheet to read.
    :param read_header_only: If True, reads only the first row (header).
    :return: A list of rows (list of lists).
    """
    wb = openpyxl.load_workbook(file_path, read_only=True)
    sheet = wb[sheet_name]
    
    rows = []
    
    if read_header_only:
        # Read only the first row (header row)
        for row in sheet.iter_rows(min_row=1, max_row=1, values_only=True):
            rows.append(list(row))
    else:
        # Read all rows
        for row in sheet.iter_rows(values_only=True):
            rows.append(list(row))
    
    return rows

