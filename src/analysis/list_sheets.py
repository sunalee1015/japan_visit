import pandas as pd
file_path = r"C:\data_workspace\ICB6_Project2\japan_visit\data\Total_Dataset_analysis_20260321.xlsx"
xl = pd.ExcelFile(file_path)
with open("sheet_names.txt", "w", encoding="utf-8") as f:
    for name in xl.sheet_names:
        f.write(name + "\n")
print("Done writing sheet names to sheet_names.txt")
