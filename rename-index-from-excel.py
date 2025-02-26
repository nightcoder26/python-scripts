import pandas as pd
import os
import sys

def main():
    if len(sys.argv) != 3:
        print("Usage: python script.py <excel_file_path> <certs_directory>")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    certs_path = sys.argv[2]
    
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found.")
        sys.exit(1)
    
    if not os.path.exists(certs_path):
        print(f"Error: Directory '{certs_path}' not found.")
        sys.exit(1)
    
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()
    names = df['Name'].tolist()
    
    print("Files in the certificate directory:")
    for file in os.listdir(certs_path):
        print(file)
    
    for index, name in enumerate(names, start=1):
        old_filename = f'CERTFICATES-{index}.pdf'
        new_filename = f'{name}.pdf'
        
        old_file_path = os.path.join(certs_path, old_filename)
        new_file_path = os.path.join(certs_path, new_filename)
        
        if os.path.exists(old_file_path):
            os.rename(old_file_path, new_file_path)
            print(f'Renamed: {old_filename} to {new_filename}')
        else:
            print(f'File not found: {old_filename}')

if __name__ == "__main__":
    main()
