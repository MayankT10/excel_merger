import os
import pandas as pd
from pathlib import Path

def merge_excel_files():
    folder_path = input("Enter the folder path containing Excel files:").strip()
    
    if not os.path.exists(folder_path):
        print(f"Error: Folder {folder_path} does not exists.")
        return 
    
    if not os.path.isdir(folder_path):
        print(f"Error: Folder {folder_path} is not a directory")
        return
    
    output_file = input("Enter the output file name ( with .xlsx extension): ").strip()
    
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'


    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx','.xls'))]
    if not excel_files:
        print(f"No Excel files found in the {folder_path}")
        return 
    
    print(f"Found {len(excel_files)} Excel files to merge ...")

    try:
        with pd.ExcelWriter(output_file,engine="openpyxl") as writer:
            for filename in os.listdir(folder_path):
                if filename.endswith(".xlsx") or filename.endswith(".xls"):
                    file_path = os.path.join(folder_path,filename) 
                    try:
                        df = pd.read_excel(file_path)
                        sheet_name = os.path.splitext(filename)[0][:31]

                        df.to_excel(writer,sheet_name=sheet_name,index=False)
                        print(f"Added: {filename} -> Sheet: {sheet_name}")
                    except Exception as e:
                        print(f"Error with file {filename}: {e}")
        print(f"\nAll the files are merged into {output_file}")
    
    except Exception as e:
        print(f"Error creating output file: {e}")

if __name__ == "__main__":
    merge_excel_files()