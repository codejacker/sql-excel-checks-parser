import re
import tkinter as tk
from tkinter import filedialog
import sys

def run_ultimate_diagnostic():
    # Ensure console can print unicode characters
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except TypeError:
        # In some environments, this might not be needed or possible
        pass

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    print("Opening file dialog to select the SQL file to diagnose...")
    sql_file_path = filedialog.askopenfilename(
        title="Select SQL File to Diagnose",
        filetypes=(("SQL files", "*.sql"), ("All files", "*.*"))
    )
    if not sql_file_path:
        print("No file selected. Exiting.")
        return

    print(f"\n--- Running ULTIMATE Diagnostic on: {sql_file_path} ---")

    # Define a few patterns to test
    patterns_to_test = {
        "Starts with --[number]": re.compile(r'^\s*--([\d\.]+)'),
        "Contains --[number]": re.compile(r'--([\d\.]+)'),
        "Contains 'סעיף'": re.compile(r'סעיף'),
        "Contains PRINT '[number]'": re.compile(r"PRINT\s+'([\d\.]+)'", re.IGNORECASE)
    }

    print("\n--- Analyzing first 100 lines of the file ---\n")

    try:
        with open(sql_file_path, 'r', encoding='cp1255') as f:
            for i, line in enumerate(f):
                if i >= 100:
                    break

                print(f"----- Line {i+1:03d} -----")

                # 1. Print the raw line
                print(f"TEXT: {line.strip()}")

                # 2. Print the hex representation
                hex_repr = ' '.join(f'{ord(c):02x}' for c in line.strip())
                print(f"HEX : {hex_repr}")

                # 3. Test the patterns
                match_found = False
                for name, pattern in patterns_to_test.items():
                    # Use search() to find a match anywhere in the line
                    match = pattern.search(line)
                    if match:
                        print(f"  ✅ {name}: MATCHED! -> Found '{match.group(0)}'")
                        match_found = True

                if not match_found:
                    print("  ❌ No patterns matched.")

                print("") # Newline for readability

    except Exception as e:
        print(f"\nERROR reading file: {e}")

if __name__ == '__main__':
    run_ultimate_diagnostic()