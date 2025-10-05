import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

# --- !! DEBUG MODE !! ---
# If the script doesn't work, remove the '#' from the line below, save, and run again.
# A new sheet named 'DEBUG_RAW_QUERIES' will be created in the output file.
# DEBUG_MODE = True

def clean_sql_query(query_text):
    """
    Cleans an individual SQL query string by removing specific lines and comments.
    This version is more aggressive.
    """
    # General pattern to remove any INSERT INTO #RUNLOG statement
    insert_pattern = re.compile(r"INSERT\s+INTO\s+#RUNLOG.*?\)\s*;?\s*\n?", re.IGNORECASE | re.DOTALL)
    query_text = insert_pattern.sub('', query_text)

    # General pattern to remove any PRINT '[...]' statement
    print_pattern = re.compile(r"PRINT\s+'[\d\.]+';?\s*\n?", re.IGNORECASE)
    query_text = print_pattern.sub('', query_text)

    # Remove '--' from the beginning of each line
    lines = query_text.split('\n')
    cleaned_lines = [re.sub(r'^\s*--', '', line) for line in lines]
    
    cleaned_query = '\n'.join(cleaned_lines).strip()
    
    return cleaned_query

def update_excel_with_sql_queries():
    """
    Reads an SQL file and an Excel file, maps cleaned queries from the SQL file
    to the Excel file based on a section number, and saves a new Excel file.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # --- File Selection ---
    sql_file_path = filedialog.askopenfilename(
        title="Select SQL File",
        filetypes=(("SQL files", "*.sql"), ("All files", "*.*"))
    )
    if not sql_file_path: 
        print("No SQL file selected. Exiting...")
        return
    
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    if not excel_file_path: 
        print("No Excel file selected. Exiting...")
        return
    
    default_filename = f"unified_testing_file_with_queries_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    output_file_path = filedialog.asksaveasfilename(
        title="Choose where to save the output file",
        initialfile=default_filename,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not output_file_path: 
        print("No save location chosen. Exiting...")
        return

    print("--- Starting Process ---")

    # --- Read SQL File ---
    try:
        with open(sql_file_path, 'r', encoding='cp1255') as f:
            sql_content = f.read()
        print(f"Successfully read SQL file.")
    except Exception as e:
        print(f"ERROR: An error occurred while reading the SQL file: {e}")
        return

    # --- Parse SQL File ---
    # This pattern now looks for lines starting with '--' and then a number like '1.1.'
    section_header_pattern = re.compile(r"^\s*--(\d+\.[\d\.]*)")
    lines = sql_content.split('\n')
    
    raw_queries = {}
    queries = {}
    current_section = None
    current_query_lines = []

    for line in lines:
        match = section_header_pattern.match(line)
        if match:
            if current_section:
                raw_query_text = '\n'.join(current_query_lines)
                raw_queries[current_section] = raw_query_text
                queries[current_section] = clean_sql_query(raw_query_text)
            
            current_section = match.group(1).strip().strip('.')
            current_query_lines = []
        elif current_section:
            current_query_lines.append(line)
    
    if current_section:
        raw_query_text = '\n'.join(current_query_lines)
        raw_queries[current_section] = raw_query_text
        queries[current_section] = clean_sql_query(raw_query_text)
    
    if not queries:
        print("ERROR: Could not extract any queries. Please check if the file contains headers like '--1.1.' that start on a new line.")
        return
    
    print(f"Found and cleaned {len(queries)} queries.")

    # --- Read and Update Excel File ---
    try:
        df = pd.read_excel(excel_file_path)
        print(f"Successfully read Excel file.")
        df.columns = [str(col).strip() for col in df.columns]
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    if 'סעיף' not in df.columns or 'סקריפט' not in df.columns:
        print("ERROR: Required columns 'סעיף' and 'סקריפט' not found. Available columns:", df.columns.tolist())
        return

    df['סעיף'] = df['סעיף'].astype(str)
    df['סקריפט'] = df['סעיף'].map(queries)

    matches_found = df['סקריפט'].notna().sum()
    print(f"Successfully mapped {matches_found} queries.")

    # --- Save the Excel File ---
    try:
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Mapped_Queries', index=False)
            
            # If debug mode is enabled, write the raw queries to a separate sheet
            try:
                if DEBUG_MODE:
                    print("DEBUG MODE IS ON: Writing raw, uncleaned queries to a separate sheet.")
                    debug_df = pd.DataFrame(list(raw_queries.items()), columns=['סעיף', 'Raw_Script'])
                    debug_df.to_excel(writer, sheet_name='DEBUG_RAW_QUERIES', index=False)
            except NameError:
                pass # DEBUG_MODE is not defined, do nothing.

        print(f"\n--- Success! ---")
        print(f"The new file has been saved to: {output_file_path}")
    except Exception as e:
        print(f"An error occurred while saving the new Excel file: {e}")

if __name__ == '__main__':
    update_excel_with_sql_queries()