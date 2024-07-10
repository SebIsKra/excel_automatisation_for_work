import os
import openpyxl
import pandas as pd 

# Function to remove write protection from an Excel file
def remove_write_protection(excel_file):
    try:
        # Load the Excel file
        wb = openpyxl.load_workbook(excel_file, read_only=False, keep_vba=False)

        # Remove write protection (if any)
        wb.security = None

        # Save the changes
        wb.save(excel_file)
    except Exception as e:
        print(f"Error removing write protection from '{excel_file}': {e}")

# Function to append rows 2 and 3 from one Excel file to another
def append_rows(source_file, target_file):
    try:
        source_wb = openpyxl.load_workbook(source_file, read_only=True)
        target_wb = openpyxl.load_workbook(target_file)

        source_ws = source_wb.active
        target_ws = target_wb.active

        # Append rows 2 and 3 from source to target
        for row in source_ws.iter_rows(min_row=2, max_row=3):
            target_ws.append([cell.value for cell in row])

        target_wb.save(target_file)
    except Exception as e:
        print(f"Error appending rows from '{source_file}' to '{target_file}': {e}")

# Function to copy the content of an Excel file to the clipboard
def copy_to_clipboard(excel_file):
    try:
        #need to add the specified columns of the work excels as parameters
        excel_file.pd.DataFrame.to_clipboard(excel=True, index=False, sep="\t", columns=['', ''])
    except Exception as e:
        print(f"Error copying content of '{excel_file}' to clipboard: {e}")

# Function to check if a file is a valid Excel file
def is_valid_excel_file(file_path):
    try:
        openpyxl.load_workbook(file_path, read_only=True)
        return True
    except Exception:
        return False


# Get the path to the user's home download folder 
downloads_folder = os.path.expanduser('~/Downloads')

# List all files in the Downloads folder
files = os.listdir(downloads_folder)

# Find the first valid Excel file in the Downloads folder
first_excel_file = None
processed_files = set()

for file in files:
    if file.startswith('picture_facebook'):
        full_path = os.path.join(downloads_folder, file)
        if is_valid_excel_file(full_path):
            if first_excel_file is None:
                first_excel_file = full_path
                print("First valid Excel file found:", first_excel_file)

                # Remove write protection from the first Excel file and add to processed_files list
                remove_write_protection(first_excel_file)
                processed_files.add(full_path)
            else:
                # Append rows from other valid Excel files to the first Excel file
                if full_path not in processed_files:
                    append_rows(full_path, first_excel_file)
                    processed_files.add(full_path)

if first_excel_file:
    print("Rows appended from other Excel files to the first Excel file.")

    # Copy the content of the first Excel file to the clipboard
    copy_to_clipboard(first_excel_file)

    print("Content of the first Excel file copied to the clipboard.")
else:
    print("No valid Excel file found in the Downloads folder.")

# Print out the processed files
print("Processed files:")
for file in processed_files:
    print(file)