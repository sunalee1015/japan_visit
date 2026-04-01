import pandas as pd
import numpy as np
import os

# 경로 설정
BASE_DIR = r"C:\data_workspace\ICB6_Project2\japan_visit"
DATA_PATH = os.path.join(BASE_DIR, "data", "Total_Dataset_analysis_20260321.xlsx")
OUT_DIR = os.path.join(BASE_DIR, "data", "processed")
os.makedirs(OUT_DIR, exist_ok=True)

def load_sheet(sheet_name):
    df = pd.read_excel(DATA_PATH, sheet_name=sheet_name, header=None).iloc[2:]
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    return df

# 1. 데이터 로드 및 전처리
# P1-2
p1_2 = load_sheet('P1-2 성별·연령·목적별 통계')
p1_2 = p1_2[(p1_2['성별'] != '승무원') & (p1_2['연령별'] != '승무원')].copy()
p1_2['인원수'] = pd.to_numeric(p1_2['인원수'], errors='coerce').fillna(0)
p1_2.rename(columns={'연령별': '연령', '목적별': '목적'}, inplace=True)

# P4-4
p4_4 = load_sheet('P4-4 유튜브 영상 목록(전체)')
for col in ['조회수', '좋아요수', '댓글수']:
    p4_4[col] = pd.to_numeric(p4_4[col], errors='coerce').fillna(0)
p4_4['조회수_log'] = np.log1p(p4_4['조회수'])

# P2-7
p2_7 = load_sheet('P2-7 일본인 방문자수·증감률')
p2_7['날짜'] = pd.to_datetime(p2_7['기준년월'].astype(str), format='%Y%m', errors='coerce')
p2_7['조회기간 방문자 수'] = pd.to_numeric(p2_7['조회기간 방문자 수'], errors='coerce').fillna(0)
p2_7['전년대비 증감 비율'] = pd.to_numeric(p2_7['전년대비 증감 비율'], errors='coerce').fillna(0)
p2_7['감소여부'] = (p2_7['전년대비 증감 비율'] < 0).astype(int)

# P2-1 / P2-2
p2_1 = load_sheet('P2-1 글로벌 지역별 방문비율')
p2_2 = load_sheet('P2-2 일본인 지역별 방문비율')
for df in [p2_1, p2_2]:
    df['시도 비율'] = pd.to_numeric(df['시도 방문자 비율'], errors='coerce')
    df['시군구 비율'] = pd.to_numeric(df['시군구 방문자 비율'], errors='coerce')
    df['절대비율'] = (df['시도 비율'] * df['시군구 비율']) / 100

df_gap = pd.merge(p2_1[['시도명', '시군구명', '절대비율']], p2_2[['시군구명', '절대비율']], 
                  on='시군구명', suffixes=('_글로벌', '_일본'))
df_gap['갭'] = df_gap['절대비율_글로벌'] - df_gap['절대비율_일본'] # 유저 스크립트 기반 (글로벌 - 일본)
df_gap['갭방향'] = np.where(df_gap['갭'] > 0, '글로벌↑', '일본↑')

# CSV 저장
p1_2.to_csv(os.path.join(OUT_DIR, "clean_p12_입국통계.csv"), index=False, encoding='utf-8-sig')
p1_2[p1_2['목적'] == '관광'].to_csv(os.path.join(OUT_DIR, "clean_p12_관광목적.csv"), index=False, encoding='utf-8-sig')
p4_4.to_csv(os.path.join(OUT_DIR, "clean_p44_유튜브.csv"), index=False, encoding='utf-8-sig')
p2_7.to_csv(os.path.join(OUT_DIR, "clean_p27_방문자수.csv"), index=False, encoding='utf-8-sig')
df_gap.to_csv(os.path.join(OUT_DIR, "clean_p21p22_지역갭.csv"), index=False, encoding='utf-8-sig')

print(f"Preprocessed CSV files saved to {OUT_DIR}")
