import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def clean_sql_query(query_text):
    """
    Cleans an individual SQL query string by removing the #RUNLOG insert and comments.
    The PRINT statement is already removed by the parsing logic.
    """
    # Remove the INSERT INTO #RUNLOG line
    insert_pattern = re.compile(r"INSERT\s+INTO\s+#RUNLOG.*\)\s*;?\s*\n?", re.IGNORECASE | re.DOTALL)
    query_text = insert_pattern.sub('', query_text)

    # Remove '--' from the beginning of each line
    lines = query_text.split('\n')
    cleaned_lines = [re.sub(r'^\s*--', '', line) for line in lines]
    
    cleaned_query = '\n'.join(cleaned_lines).strip()
    
    return cleaned_query

def update_excel_with_sql_queries():
    """
    Prompts user for files, reads SQL and Excel, maps cleaned queries based on PRINT statements,
    and saves to a user-chosen location.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # --- File Selection ---
    sql_file_path = filedialog.askopenfilename(title="Select SQL File", filetypes=(("SQL files", "*.sql"), ("All files", "*.*")))
    if not sql_file_path: return
    
    excel_file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
    if not excel_file_path: return
    
    default_filename = f"unified_testing_file_with_queries_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    output_file_path = filedialog.asksaveasfilename(title="Choose where to save the output file", initialfile=default_filename, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")] )
    if not output_file_path: return

    print("---" + "-" * 17 + " Starting Process " + "-" * 17 + "---")

    # --- Read SQL File ---
    try:
        with open(sql_file_path, 'r', encoding='cp1255') as f:
            sql_content = f.read()
        print(f"Successfully read SQL file.")
    except Exception as e:
        print(f"ERROR: An error occurred while reading the SQL file: {e}")
        return

    # --- CORRECTED PARSING LOGIC: Use PRINT as the delimiter ---
    delimiter_pattern = re.compile(r"PRINT\s+'([\d\.]+)'", re.IGNORECASE)
    
    # The split results in: [text_before, section_num1, query1, section_num2, query2, ...]
    split_content = delimiter_pattern.split(sql_content)
    
    if len(split_content) < 3:
        print("ERROR: Could not find any queries using the PRINT '[number]' pattern in the SQL file.")
        return

    # Extract section numbers and the raw queries that follow them
    section_numbers = split_content[1::2]
    raw_queries = split_content[2::2]

    queries = {}
    for i, section_number in enumerate(section_numbers):
        raw_query = raw_queries[i]
        # The raw_query no longer contains the PRINT, but it does contain the INSERT and comments.
        cleaned_query = clean_sql_query(raw_query)
        
        if cleaned_query:
            queries[section_number] = cleaned_query

    print(f"Found and cleaned {len(queries)} queries based on PRINT statements.")

    # --- Read and Update Excel File ---
    try:
        df = pd.read_excel(excel_file_path)
        print(f"Successfully read Excel file.")
        df.columns = [str(col).strip() for col in df.columns]
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    if 'סעיף' not in df.columns or 'סקריפט' not in df.columns:
        print("ERROR: Required columns 'סעיף' and 'סקריפט' not found.")
        return

    df['סעיף'] = df['סעיף'].astype(str)
    df['סקריפט'] = df['סעיף'].map(queries)

    matches_found = df['סקריפט'].notna().sum()
    print(f"Successfully mapped {matches_found} queries.")
    if matches_found < len(df):
        missing_sections = df[df['סקריפט'].isna()]['סעיף'].tolist()
        print(f"Warning: {len(missing_sections)} sections in Excel had no matching query: {missing_sections}")

    # --- Save the Excel File ---
    try:
        df.to_excel(output_file_path, index=False, engine='openpyxl')
        print(f"\n---" + "-" * 17 + " Success! " + "-" * 17 + "---")
        print(f"The new file has been saved to: {output_file_path}")
    except Exception as e:
        print(f"An error occurred while saving the new Excel file: {e}")

if __name__ == '__main__':
    update_excel_with_sql_queries()