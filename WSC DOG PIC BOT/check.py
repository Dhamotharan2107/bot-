import os
import pandas as pd
import glob
import tkinter as tk
from tkinter import filedialog

# Function to count image files in a folder
def count_images_in_folder(folder_path):
    # Define image extensions
    image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.tiff']
    image_files = []
    
    # Search for all image files in the folder
    for ext in image_extensions:
        image_files.extend(glob.glob(os.path.join(folder_path, ext)))
    
    return len(image_files)

# Function to update the Excel with the image count
def update_excel_with_image_count(excel_file, folder_to_copy):
    # Read the Excel file
    df = pd.read_excel(excel_file)
    
    # Strip any leading/trailing spaces from the column names
    df.columns = df.columns.str.strip()

    # Print the column names to help with debugging
    print("Columns in the Excel file:", df.columns)

    # Prompt user to manually enter the column name for uqid
    uqid_column = input("Enter the column name for uqid (as it appears in the Excel file): ").strip()

    # Get the list of folders in the root directory (folder_to_copy)
    folders_in_directory = [folder for folder in os.listdir(folder_to_copy) if os.path.isdir(os.path.join(folder_to_copy, folder))]

    # Iterate through each folder in the root directory
    for folder_name in folders_in_directory:
        image_count = count_images_in_folder(os.path.join(folder_to_copy, folder_name))
        if folder_name in df[uqid_column].values:  # Check if the folder name exists in the 'uqid' column of the Excel file
            # Update the "Image Count" column in the Excel file for existing uqid
            df.loc[df[uqid_column] == folder_name, 'Image Count'] = image_count
        else:
            print(f"Folder {folder_name} not found in Excel. Adding new entry.")
            # Add a new row for the missing folder and its image count
            new_row = {uqid_column: folder_name, 'Image Count': image_count}
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    # Save the updated DataFrame to the Excel file
    df.to_excel(excel_file, index=False)
    print(f"Excel file updated successfully: {excel_file}")

# Function to allow the user to manually select files and folders
def select_file_and_folder():
    # Create the Tkinter root window (it won't be shown)
    root = tk.Tk()
    root.withdraw()

    # Let the user select the Excel file
    excel_file = filedialog.askopenfilename(
        title="Select the Excel File", 
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    if not excel_file:
        print("No Excel file selected.")
        return None, None

    # Let the user select the root folder (the folder containing the uqid folders)
    folder_to_copy = filedialog.askdirectory(
        title="Select the Root Folder"
    )
    if not folder_to_copy:
        print("No folder selected.")
        return None, None

    return excel_file, folder_to_copy

# Main execution
def main():
    # Ask the user to select the Excel file and the root folder
    excel_file, folder_to_copy = select_file_and_folder()
    if excel_file and folder_to_copy:
        # Update the Excel file with image counts
        update_excel_with_image_count(excel_file, folder_to_copy)

if __name__ == "__main__":
    main()
