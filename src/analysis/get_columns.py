import pandas as pd
file_path = r"C:\data_workspace\ICB6_Project2\japan_visit\data\Total_Dataset_analysis_20260321.xlsx"

sheets = {
    'P1-2': 'P1-2 성별·연령·목적별 통계',
    'P4-4': 'P4-4 유튜브 영상 목록(전체)',
    'P2-7': 'P2-7 일본인 방문자수·증감률',
    'P2-1': 'P2-1 글로벌 지역별 방문비율',
    'P2-2': 'P2-2 일본인 지역별 방문비율'
}

with open("col_info.txt", "w", encoding="utf-8") as f:
    for prefix, full_name in sheets.items():
        try:
            df = pd.read_excel(file_path, sheet_name=full_name, header=None).iloc[2:]
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
            f.write(f"\n--- {prefix} Columns ---\n")
            f.write(", ".join([str(c) for c in df.columns.tolist()]) + "\n")
            f.write("Sample Head:\n")
            f.write(df.head(1).to_string() + "\n")
        except Exception as e:
            f.write(f"Error reading {full_name}: {e}\n")
print("Done writing column info to col_info.txt")
