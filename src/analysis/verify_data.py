import pandas as pd

file_path = r"C:\data_workspace\ICB6_Project2\japan_visit\data\Total_Dataset_analysis_20260321.xlsx"

try:
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    print("Full Sheet Names:")
    for name in sheet_names:
        print(f" - {name}")
    
    target_prefixes = ['P1-2', 'P4-4', 'P2-7', 'P2-1', 'P2-2']
    mapping = {}
    for prefix in target_prefixes:
        match = [s for s in sheet_names if s.startswith(prefix)]
        if match:
            mapping[prefix] = match[0]
            print(f"Found {prefix} -> {match[0]}")
        else:
            print(f"Prefix {prefix} NOT found")
            
    for prefix, full_name in mapping.items():
        df = pd.read_excel(file_path, sheet_name=full_name, header=None, nrows=5)
        print(f"\n--- {full_name} ---")
        print(df)
except Exception as e:
    print(f"Error: {e}")
