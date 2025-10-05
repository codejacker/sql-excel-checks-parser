import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def clean_sql_query(query_text):
    """
    Cleans an individual SQL query string by removing specific lines and comments.
    """
    # Remove the INSERT INTO #RUNLOG line (matches any section number)
    insert_pattern = re.compile(r"INSERT\s+INTO\s+#RUNLOG\s+VALUES\s*('[\\d\\.]+'.*?)\s*\n?", re.IGNORECASE | re.DOTALL)
    query_text = insert_pattern.sub('', query_text)

    # Remove the PRINT line (matches any section number)
    print_pattern = re.compile(r"PRINT\s+'[\\d\\.]+';?\s*\n?", re.IGNORECASE)
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
    
    # --- Step 1: Select Input Files ---
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
    
    # --- Step 2: Select Output File Location ---
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

    # --- Step 3: Read and Parse the SQL file ---
    try:
        with open(sql_file_path, 'r', encoding='cp1255') as f:
            sql_content = f.read()
        print(f"Successfully read SQL file using cp1255 encoding.")
    except Exception as e:
        print(f"An error occurred while reading the SQL file: {e}")
        return

    queries = {}
    current_section = None
    current_query_lines = []
    
    # FINAL, MORE SPECIFIC PATTERN: Looks for '--' at the start of a line,
    # followed by a number that MUST contain at least one dot.
    section_header_pattern = re.compile(r"^\s*--(\d+\.[\d\.]*)")

    for line in sql_content.split('\n'):
        match = section_header_pattern.match(line)
        if match:
            # If we were already building a query, save it before starting the new one
            if current_section:
                raw_query = '\n'.join(current_query_lines)
                queries[current_section] = clean_sql_query(raw_query)
            
            # Start the new section
            current_section = match.group(1).strip().strip('.')
            current_query_lines = []
        elif current_section:
            # If we are inside a section, append the line to its query
            current_query_lines.append(line)
    
    # Save the very last query in the file
    if current_section and current_query_lines:
        raw_query = '\n'.join(current_query_lines)
        queries[current_section] = clean_sql_query(raw_query)
    
    if not queries:
        print("ERROR: Could not extract any queries. Please check if the file contains headers like '--1.1.' that start on a new line.")
        return
    
    print(f"Found and cleaned {len(queries)} queries in the SQL file.")

    # --- Step 4: Read and Update the Excel file ---
    try:
        df = pd.read_excel(excel_file_path)
        print(f"Successfully read Excel file: {excel_file_path}")
        # Clean whitespace from all column names
        df.columns = [str(col).strip() for col in df.columns]
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    if 'סעיף' not in df.columns or 'סקריפט' not in df.columns:
        print("ERROR: Required columns 'סעיף' and 'סקריפט' not found after cleaning headers.")
        print("Available columns:", df.columns.tolist())
        return

    df['סעיף'] = df['סעיף'].astype(str)
    df['סקריפט'] = df['סעיף'].map(queries)

    matches_found = df['סקריפט'].notna().sum()
    print(f"Successfully mapped {matches_found} queries to the Excel data.")
    if matches_found < len(df):
        print(f"Warning: {len(df) - matches_found} rows did not have a corresponding query.")

    # --- Step 5: Save the new Excel file ---
    try:
        df.to_excel(output_file_path, index=False, engine='openpyxl')
        print(f"\n--- Success! ---")
        print(f"The new file has been saved to: {output_file_path}")
    except Exception as e:
        print(f"An error occurred while saving the new Excel file: {e}")

if __name__ == '__main__':
    update_excel_with_sql_queries()
