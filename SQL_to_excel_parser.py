import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def update_excel_with_sql_queries():
    """
    Reads an SQL file and an Excel file, maps queries from the SQL file
    to the Excel file based on a section number, and saves a new Excel file.
    """
    # Create and hide the tkinter root window
    root = tk.Tk()
    root.withdraw()
    
    # Select SQL file
    sql_file_path = filedialog.askopenfilename(
        title="Select SQL File",
        filetypes=(("SQL files", "*.sql"), ("All files", "*.*"))
    )
    if not sql_file_path:
        print("No SQL file selected. Exiting...")
        return
    
    # Select Excel file
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    if not excel_file_path:
        print("No Excel file selected. Exiting...")
        return
    
    # Select output directory
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    if not output_directory:
        print("No output directory selected. Exiting...")
        return
        
    # Generate output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file_path = os.path.join(output_directory, f"unified_testing_file_{timestamp}.xlsx")

    print("--- Starting Process ---")

    # --- Step 1: Read and Parse the SQL file ---
    encodings_to_try = ['cp1255', 'utf-8-sig', 'utf-8', 'windows-1255']
    sql_content = None
    
    for encoding in encodings_to_try:
        try:
            with open(sql_file_path, 'r', encoding=encoding) as f:
                sql_content = f.read()
            print(f"Successfully read SQL file using {encoding} encoding: {sql_file_path}")
            break
        except FileNotFoundError:
            print(f"ERROR: SQL file not found at: {sql_file_path}")
            print("Please make sure the file path is correct.")
            return
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f"An error occurred while reading the SQL file: {e}")
            continue
    
    if sql_content is None:
        print("ERROR: Could not read the SQL file with any of the attempted encodings.")
        return

        # The sections are marked with number comments like '--1.1', '--2.1' etc.
    lines = sql_content.split('\n')
    queries = {}
    current_section = None
    current_query = []
    
    # Regular expression to match section headers
    section_header = re.compile(r'^-{2,3}\s*(\d+(?:\.\d+)*(?:-)?)')
    
    # Process each line
    for line in lines:
        # Check if this line is a section header
        match = section_header.match(line.strip())
        if match:
            # If we were collecting a previous section, save it
            if current_section:
                queries[current_section] = '\n'.join(current_query).strip()
            
            # Start new section
            current_section = match.group(1).rstrip('-')  # Remove trailing dash if present
            current_query = []
        elif current_section:
            # Add line to current query
            current_query.append(line)
    
    # Don't forget to save the last section
    if current_section and current_query:
        queries[current_section] = '\n'.join(current_query).strip()
    
    if not queries:
        print("ERROR: Could not extract any queries from the SQL file.")
        print("Please ensure the SQL file contains numbered sections.")
        return
    
    # Print what we found
    print(f"\nFound {len(queries)} sections in the SQL file.")
    print("Section numbers:", sorted(queries.keys()))

    print(f"Found {len(queries)} queries in the SQL file.")

    # --- Step 2: Read and Update the Excel file ---
    try:
        # First read Excel file and show all columns
        df = pd.read_excel(excel_file_path)
        print(f"\nSuccessfully read Excel file: {excel_file_path}")
        print("\nAvailable columns in Excel file:")
        for col in df.columns:
            print(f"Column: '{col}' (type: {type(col)}, length: {len(str(col))}")
    except FileNotFoundError:
        print(f"ERROR: Excel file not found at: {excel_file_path}")
        print("Please make sure the file path is correct.")
        return
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    # Look for required columns (case-sensitive but flexible with spaces)
    section_col = None
    script_col = None

    for col in df.columns:
        col_str = str(col).strip()
        if 'סעיף' in col_str:
            section_col = col
            print(f"Found section column: '{col}'")
        if 'סקריפט' in col_str:
            script_col = col
            print(f"Found script column: '{col}'")

    if section_col is None or script_col is None:
        print("\nERROR: Required columns not found")
        print("Looking for columns containing 'סעיף' and 'סקריפט'")
        print("\nAvailable columns:")
        for col in df.columns:
            print(f"- '{col}'")
        return

    # Convert the section column to string to ensure matching with dictionary keys
    df[section_col] = df[section_col].astype(str)

    # --- Step 3: Map the queries to the DataFrame ---
    # Map the queries using the detected column names
    print(f"\nMapping queries from section column '{section_col}' to script column '{script_col}'")
    df[script_col] = df[section_col].map(queries)

    matches_found = df[script_col].notna().sum()
    print(f"Successfully mapped {matches_found} queries to the Excel data.")
    if matches_found < len(df):
        print(f"Warning: {len(df) - matches_found} rows in the Excel file did not have a corresponding query.")

    # --- Step 4: Save the new Excel file ---
    try:
        df.to_excel(output_file_path, index=False, engine='openpyxl')
        print(f"\n--- Success! ---")
        print(f"The new file has been saved to: {output_file_path}")
    except Exception as e:
        print(f"An error occurred while saving the new Excel file: {e}")

if __name__ == '__main__':
    update_excel_with_sql_queries()