import os
import pandas as pd
import glob
import tkinter as tk
from tkinter import filedialog

def count_images_in_folder(folder_path):
    image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.tiff']
    image_files = []
    for ext in image_extensions:
        image_files.extend(glob.glob(os.path.join(folder_path, ext)))
    return len(image_files)

def update_excel_with_image_count(excel_file, folder_to_copy):
    # Read the Excel file
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()  # Clean column names

    # Validate or initialize required columns
    for col in ['Image Count', 'Received Count', 'Total Count']:
        if col not in df.columns:
            df[col] = 0  # Initialize missing columns with 0

    # Ensure NaN values in relevant columns are treated as 0
    df['Image Count'] = df['Image Count'].fillna(0).astype(int)
    df['Received Count'] = df['Received Count'].fillna(0).astype(int)

    uqid_column = input("Enter the column name for uqid: ").strip()
    if uqid_column not in df.columns:
        print(f"Error: {uqid_column} column not found in Excel!")
        return

    folders_in_directory = [folder.strip() for folder in os.listdir(folder_to_copy) if os.path.isdir(os.path.join(folder_to_copy, folder))]

    for folder_name in folders_in_directory:
        folder_path = os.path.join(folder_to_copy, folder_name)
        image_count = count_images_in_folder(folder_path)

        if folder_name in df[uqid_column].values:
            row_index = df[df[uqid_column] == folder_name].index[0]
            received_count = df.at[row_index, 'Received Count']
            df.at[row_index, 'Image Count'] = image_count
            df.at[row_index, 'Total Count'] = received_count + image_count
        else:
            print(f"Adding new entry for folder: {folder_name}")
            new_row = {
                uqid_column: folder_name,
                'Image Count': image_count,
                'Received Count': 0,
                'Total Count': image_count,
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    # Recalculate Total Count for all rows
    df['Total Count'] = df['Received Count'] + df['Image Count']

    df.to_excel(excel_file, index=False)
    print(f"Excel file updated: {excel_file}")

def select_file_and_folder():
    root = tk.Tk()
    root.withdraw()
    excel_file = filedialog.askopenfilename(title="Select the Excel File", filetypes=[("Excel Files", "*.xlsx")])
    folder_to_copy = filedialog.askdirectory(title="Select the Root Folder")
    return excel_file, folder_to_copy

def main():
    excel_file, folder_to_copy = select_file_and_folder()
    if excel_file and folder_to_copy:
        update_excel_with_image_count(excel_file, folder_to_copy)

if __name__ == "__main__":
    main()
