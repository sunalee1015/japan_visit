import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import koreanize_matplotlib
import os
import scipy.stats as stats
import statsmodels.api as sm
from statsmodels.tsa.seasonal import seasonal_decompose
from statsmodels.graphics.tsaplots import plot_acf
from sklearn.cluster import KMeans
from sklearn.preprocessing import LabelEncoder

# 경로 설정
BASE_DIR = r"C:\data_workspace\ICB6_Project2\japan_visit"
DATA_PATH = os.path.join(BASE_DIR, "data", "Total_Dataset_analysis_20260321.xlsx")
IMAGE_DIR = os.path.join(BASE_DIR, "images")
os.makedirs(IMAGE_DIR, exist_ok=True)

def load_sheet(sheet_name):
    df = pd.read_excel(DATA_PATH, sheet_name=sheet_name, header=None).iloc[2:]
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    return df

def detect_outliers_iqr(df, column):
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    outliers = df[(df[column] < lower_bound) | (df[column] > upper_bound)]
    return outliers, lower_bound, upper_bound

def cramers_v(contingency_table):
    chi2 = stats.chi2_contingency(contingency_table)[0]
    n = contingency_table.sum().sum()
    phi2 = chi2 / n
    r, k = contingency_table.shape
    phi2corr = max(0, phi2 - ((k-1)*(r-1))/(n-1))
    rcorr = r - ((r-1)**2)/(n-1)
    kcorr = k - ((k-1)**2)/(n-1)
    return np.sqrt(phi2corr / min((kcorr-1), (rcorr-1)))

# --- STEP 1: 전처리 ---
p1_2 = load_sheet('P1-2 성별·연령·목적별 통계')
p4_4 = load_sheet('P4-4 유튜브 영상 목록(전체)')
p2_7 = load_sheet('P2-7 일본인 방문자수·증감률')
p2_1 = load_sheet('P2-1 글로벌 지역별 방문비율')
p2_2 = load_sheet('P2-2 일본인 지역별 방문비율')

# P1-2 필터링 및 타입 변환
p1_2 = p1_2[(p1_2['성별'] != '승무원') & (p1_2['연령별'] != '승무원')].copy()
p1_2['인원수'] = pd.to_numeric(p1_2['인원수'], errors='coerce').fillna(0)

# 인코딩 (성별 0/1, 연령 순서형, 목적 더미)
p1_2['성별_enc'] = p1_2['성별'].map({'남성': 0, '여성': 1})
age_map = {'20세 이하': 0, '21-30세': 1, '31-40세': 2, '41-50세': 3, '51-60세': 4, '61세 이상': 5}
p1_2['연령_enc'] = p1_2['연령별'].map(age_map)
p1_2_dummies = pd.get_dummies(p1_2['목적별'], prefix='목적')

# P4-4: 조회수 로그변환 및 날짜 파싱
for col in ['조회수', '좋아요수', '댓글수']:
    p4_4[col] = pd.to_numeric(p4_4[col], errors='coerce').fillna(0)
    p4_4[f'log_{col}'] = np.log1p(p4_4[col])

p4_4['업로드 날짜'] = pd.to_datetime(p4_4['업로드 날짜'])
p4_4['year'] = p4_4['업로드 날짜'].dt.year
p4_4['month'] = p4_4['업로드 날짜'].dt.month

# P2-1/P2-2 병합 및 절대비율
for df in [p2_1, p2_2]:
    df['시도 비율'] = pd.to_numeric(df['시도 방문자 비율'], errors='coerce')
    df['시군구 비율'] = pd.to_numeric(df['시군구 방문자 비율'], errors='coerce')
    df['절대비율'] = (df['시도 비율'] * df['시군구 비율']) / 100

p2_merged = pd.merge(p2_1[['시군구명', '절대비율']], p2_2[['시군구명', '절대비율']], 
                      on='시군구명', suffixes=('_Global', '_Japan')).dropna()

# --- STEP 2: 단변량 분석 ---
# 수치형 기술통계 및 Shapiro-Wilk
print("\n[P4-4 기술통계]")
print(p4_4[['조회수', '좋아요수', '댓글수']].describe())
print("\n[P4-4 Shapiro-Wilk 정규성 검정 (로그 조회수)]")
stat, p = stats.shapiro(p4_4['log_조회수'])
print(f"Statistics={stat:.4f}, p={p:.4f}")

# 히스토그램 & 박스플롯
plt.figure(figsize=(12, 5))
plt.subplot(1, 2, 1)
plt.hist(p4_4['log_조회수'], bins=30, color='lightgreen', edgecolor='black')
plt.title('Log Video Views Distribution')
plt.subplot(1, 2, 2)
plt.boxplot(p4_4['log_조회수'])
plt.title('Log Video Views Boxplot')
plt.savefig(os.path.join(IMAGE_DIR, "step2_univariate_youtube.png"))
plt.close()

# 퍼센타일 분포
print("\n[유튜브 조회수 퍼센타일]")
print(p4_4['조회수'].quantile([0.25, 0.5, 0.75, 0.9, 0.95, 0.99]))

# --- STEP 3: 다변량 분석 ---
# Chi-square: 성별 x 목적
contingency = pd.crosstab(p1_2['성별'], p1_2['목적별'])
chi2, p, dof, ex = stats.chi2_contingency(contingency)
print(f"\n[Chi-square 성별 x 목적] chi2={chi2:.4f}, p={p:.4f}")
print(f"Cramer's V: {cramers_v(contingency):.4f}")

# 상관분석: Pearson vs Spearman
print("\n[P4-4 상관계수 비교]")
print("Pearson:\n", p4_4[['조회수', '좋아요수', '댓글수']].corr(method='pearson'))
print("Spearman:\n", p4_4[['조회수', '좋아요수', '댓글수']].corr(method='spearman'))

# 다중회귀 (OLS)
X = p4_4[['log_좋아요수', 'log_댓글수', 'year', 'month']]
X = sm.add_constant(X)
y = p4_4['log_조회수']
model_ols = sm.OLS(y, X).fit()
print("\n[OLS Regression Summary]")
print(model_ols.summary())

# K-Means (k=4): 글로벌비율 x 일본비율
kmeans = KMeans(n_clusters=4, random_state=42)
p2_merged['cluster'] = kmeans.fit_predict(p2_merged[['절대비율_Global', '절대비율_Japan']])

plt.figure(figsize=(8, 6))
plt.scatter(p2_merged['절대비율_Global'], p2_merged['절대비율_Japan'], c=p2_merged['cluster'], cmap='viridis')
plt.xlabel('Global Visit Ratio')
plt.ylabel('Japan Visit Ratio')
plt.title('Region Clustering (k=4): Global vs Japan')
plt.colorbar(label='Cluster')
plt.savefig(os.path.join(IMAGE_DIR, "step3_clustering_k4.png"))
plt.close()

# 시계열 ACF
p2_7['기준년월'] = pd.to_datetime(p2_7['기준년월'].astype(str), format='%Y%m')
p2_7_ts = p2_7.set_index('기준년월')['조회기간 방문자 수'].sort_index()
plt.figure(figsize=(10, 4))
plot_acf(p2_7_ts, lags=10, ax=plt.gca())
plt.title('ACF of Visitor Counts')
plt.savefig(os.path.join(IMAGE_DIR, "step3_time_series_acf.png"))
plt.close()

# 이동평균 추세
p2_7_ts.rolling(window=3).mean().plot(label='3-Month Moving Avg', color='red')
p2_7_ts.plot(label='Original', alpha=0.5)
plt.title('Visitor Counts with Moving Average')
plt.legend()
plt.savefig(os.path.join(IMAGE_DIR, "step3_moving_average.png"))
plt.close()

# 모든 통계 결과를 파일로 저장
with open(os.path.join(BASE_DIR, "docs", "stats_report.txt"), "w", encoding="utf-8") as f:
    f.write("[P4-4 기술통계]\n")
    f.write(p4_4[['조회수', '좋아요수', '댓글수']].describe().to_string() + "\n\n")
    
    f.write("[P4-4 Shapiro-Wilk (log 조회수)]\n")
    stat, p = stats.shapiro(p4_4['log_조회수'])
    f.write(f"Statistics={stat:.4f}, p={p:.4f}\n\n")
    
    f.write("[Chi-square 성별 x 목적]\n")
    contingency = pd.crosstab(p1_2['성별'], p1_2['목적별'])
    chi2, p_val, dof, ex = stats.chi2_contingency(contingency)
    f.write(f"chi2={chi2:.4f}, p={p_val:.4f}\n")
    f.write(f"Cramer's V: {cramers_v(contingency):.4f}\n\n")
    
    f.write("[P4-4 상관계수 비교]\n")
    f.write("Pearson:\n" + p4_4[['조회수', '좋아요수', '댓글수']].corr(method='pearson').to_string() + "\n")
    f.write("Spearman:\n" + p4_4[['조회수', '좋아요수', '댓글수']].corr(method='spearman').to_string() + "\n\n")
    
    f.write("[OLS Regression Summary]\n")
    f.write(model_ols.summary().as_text() + "\n")

print("\nRefined EDA Analysis complete. Advanced statistics results saved to docs/stats_report.txt")
