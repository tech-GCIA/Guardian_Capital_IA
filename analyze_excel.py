import pandas as pd
import sys

def analyze_excel_headers():
    try:
        # Read the Excel file
        df = pd.read_excel('../Base Sheet.xlsx', header=None)
        
        print("=== BASE SHEET.XLSX HEADER ANALYSIS ===")
        print(f"Total shape: {df.shape}")
        print()
        
        # Analyze first 8 rows (header rows)
        for i in range(min(8, len(df))):
            print(f"=== ROW {i+1} ===")
            row_data = df.iloc[i]
            
            # Show non-empty cells with their column positions
            for j, val in enumerate(row_data):
                if pd.notna(val) and str(val).strip() != '':
                    print(f"  Column {j+1} (Excel: {chr(65 + j % 26)}): '{val}'")
            print()
        
        # Show column count
        print(f"Total columns: {df.shape[1]}")
        
        # Look for patterns in the first few data rows after headers
        print("\n=== SAMPLE DATA ROWS (9-12) ===")
        for i in range(8, min(12, len(df))):
            print(f"Row {i+1} sample (first 10 non-empty cells):")
            row_data = df.iloc[i]
            count = 0
            for j, val in enumerate(row_data):
                if pd.notna(val) and str(val).strip() != '' and count < 10:
                    print(f"  Col {j+1}: '{val}'")
                    count += 1
            print()
            
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    analyze_excel_headers()