import pandas as pd
import warnings
warnings.filterwarnings('ignore')

def process_file(filename):
    print(f"--- Processing {filename} ---")
    xl = pd.ExcelFile(filename)
    for sheet in xl.sheet_names:
        df = pd.read_excel(filename, sheet_name=sheet)
        print(f"Sheet: {sheet}, Shape: {df.shape}")
        
        # Check if Date exists in columns
        date_col = next((col for col in df.columns if 'date' in str(col).lower()), None)
        
        if date_col:
            try:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                april_data = df[(df[date_col].dt.month == 4) & (df[date_col].dt.year == 2026)]
                print(f"Found {len(april_data)} rows for April 2026 in column '{date_col}'")
                if not april_data.empty:
                    print(april_data.head().to_string())
            except:
                pass
        else:
            # Try to find date in first few columns if header isn't clear
            for col in df.columns[:3]:
                try:
                    dates = pd.to_datetime(df[col], errors='coerce')
                    april_data = df[(dates.dt.month == 4) & (dates.dt.year == 2026)]
                    if not april_data.empty:
                        print(f"Found {len(april_data)} rows for April 2026 in column '{col}'")
                        print(april_data.head().to_string())
                        break
                except:
                    continue
    print("-" * 30)

process_file("TAT - ALL TOKENS.xlsx")
process_file("TAT - ALL COMPLETED TOKENS.xlsx")
