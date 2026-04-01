import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
from scipy import stats
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.metrics import r2_score, mean_absolute_error
from sklearn.model_selection import train_test_split
import warnings
import os
warnings.filterwarnings('ignore')

# 경로 설정
BASE_DIR = r"C:\data_workspace\ICB6_Project2\japan_visit"
OUT = os.path.join(BASE_DIR, "data", "processed")
IMAGE_OUT = os.path.join(BASE_DIR, "images")
os.makedirs(IMAGE_OUT, exist_ok=True)

# 한글 폰트 설정 (Windows Malgun Gothic)
font_path = r'C:\Windows\Fonts\malgun.ttf'
if os.path.exists(font_path):
    prop = fm.FontProperties(fname=font_path)
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.family'] = prop.get_name()
plt.rcParams['axes.unicode_minus'] = False

# 데이터 로드
df_p12  = pd.read_csv(f"{OUT}/clean_p12_입국통계.csv",   encoding='utf-8-sig')
df_tour = pd.read_csv(f"{OUT}/clean_p12_관광목적.csv",   encoding='utf-8-sig')
df_yt   = pd.read_csv(f"{OUT}/clean_p44_유튜브.csv",     encoding='utf-8-sig')
df_vis  = pd.read_csv(f"{OUT}/clean_p27_방문자수.csv",   encoding='utf-8-sig')
df_gap  = pd.read_csv(f"{OUT}/clean_p21p22_지역갭.csv",  encoding='utf-8-sig')
df_vis['날짜'] = pd.to_datetime(df_vis['날짜'], errors='coerce')
# Extract year/month for regression if not already there
df_yt['업로드 날짜'] = pd.to_datetime(df_yt['업로드 날짜'])
df_yt['업로드연도'] = df_yt['업로드 날짜'].dt.year
df_yt['업로드월'] = df_yt['업로드 날짜'].dt.month
# Ensure log cols exist
for col in ['조회수', '좋아요수', '댓글수']:
    if f'{col}_log' not in df_yt.columns:
        df_yt[f'{col}_log'] = np.log1p(df_yt[col])

print("=" * 60)
print("STEP 3: 다변량 분석")
print("=" * 60)

with open(os.path.join(BASE_DIR, "docs", "eda_step3_report.txt"), "w", encoding="utf-8") as f:
    f.write("=" * 60 + "\n")
    f.write("STEP 3: 다변량 분석\n")
    f.write("=" * 60 + "\n")

    # ════════════════════════════════════════
    # 3-A. 교차분석 (Chi-square Test)
    # ════════════════════════════════════════
    f.write("\n【3-A】 교차분석 (Chi-square Test)\n")
    f.write("-" * 40 + "\n")

    # 성별 × 목적별
    cross1 = pd.crosstab(df_p12['성별'], df_p12['목적'], values=df_p12['인원수'], aggfunc='sum').fillna(0)
    chi2_1, p1, dof1, expected1 = stats.chi2_contingency(cross1)
    f.write("  [성별 × 목적별 교차표 (단위: 만명)]\n")
    f.write((cross1 / 1e4).round(1).to_string() + "\n")
    f.write(f"\n  χ²={chi2_1:.2f}, df={dof1}, p={p1:.6f}\n")

    # 성별 × 연령별 (관광목적)
    cross2 = pd.crosstab(df_tour['성별'], df_tour['연령'], values=df_tour['인원수'], aggfunc='sum').fillna(0)
    chi2_2, p2, dof2, expected2 = stats.chi2_contingency(cross2)
    연령_순서 = ['20세 이하','21~30세','31~40세','41~50세','51~60세','61세이상']
    cross2 = cross2.reindex(columns=[c for c in 연령_순서 if c in cross2.columns])
    f.write(f"\n  [성별 × 연령 교차표 (관광목적, 단위: 만명)]\n")
    f.write((cross2 / 1e4).round(1).to_string() + "\n")
    f.write(f"\n  χ²={chi2_2:.2f}, df={dof2}, p={p2:.6f}\n")

    n_total = cross2.sum().sum()
    c_v = np.sqrt(chi2_2 / (n_total * (min(cross2.shape) - 1)))
    f.write(f"  Cramér's V = {c_v:.4f}\n")

    # ════════════════════════════════════════
    # 3-B. 상관분석
    # ════════════════════════════════════════
    f.write("\n【3-B】 상관분석\n")
    f.write("-" * 40 + "\n")
    corr_matrix = df_yt[['조회수','좋아요수','댓글수']].corr(method='pearson')
    f.write("  [유튜브 지표 간 Pearson 상관계수]\n")
    f.write(corr_matrix.round(4).to_string() + "\n")

    # ════════════════════════════════════════
    # 3-C. 다중회귀분석
    # ════════════════════════════════════════
    f.write("\n【3-C】 다중회귀분석 — 유튜브 조회수 예측\n")
    f.write("-" * 40 + "\n")
    X_cols = ['좋아요수_log','댓글수_log','업로드연도','업로드월']
    df_reg = df_yt[X_cols + ['조회수_log']].dropna()
    X = df_reg[X_cols].values
    y = df_reg['조회수_log'].values
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    scaler = StandardScaler()
    X_train_s = scaler.fit_transform(X_train)
    X_test_s  = scaler.transform(X_test)
    reg = LinearRegression()
    reg.fit(X_train_s, y_train)
    y_pred = reg.predict(X_test_s)
    r2 = r2_score(y_test, y_pred)
    f.write(f"  [모델 성능] R² = {r2:.4f} (설명력 {r2*100:.1f}%)\n")
    f.write(f"  [회귀 계수 (표준화)]\n")
    for col, coef in zip(X_cols, reg.coef_):
        f.write(f"  {col:15s}: {coef:+.4f}\n")

    # ════════════════════════════════════════
    # 3-D. K-Means 군집분석
    # ════════════════════════════════════════
    f.write("\n【3-D】 K-Means 군집분석 — 지역 방문 패턴 분류\n")
    f.write("-" * 40 + "\n")
    X_km = df_gap[['절대비율_글로벌','절대비율_일본']].dropna().values
    scaler_km = StandardScaler()
    X_km_s = scaler_km.fit_transform(X_km)
    inertias = []
    K_range = range(2, 9)
    for k in K_range:
        km = KMeans(n_clusters=k, random_state=42, n_init=10)
        km.fit(X_km_s)
        inertias.append(km.inertia_)
    km_final = KMeans(n_clusters=4, random_state=42, n_init=10)
    df_gap_km = df_gap.dropna(subset=['절대비율_글로벌','절대비율_일본']).copy()
    df_gap_km['군집'] = km_final.fit_predict(X_km_s)
    cluster_stats = df_gap_km.groupby('군집')[['절대비율_글로벌','절대비율_일본','갭']].mean().round(3)
    labels = {}
    for c in cluster_stats.index:
        g = cluster_stats.loc[c,'절대비율_글로벌']; j = cluster_stats.loc[c,'절대비율_일본']; gap_v = cluster_stats.loc[c,'갭']
        if g > 2 and j > 2: labels[c] = '서울·인천 집중'
        elif gap_v > 0.3: labels[c] = '분산 타겟'
        elif gap_v < -0.3: labels[c] = '일본 편향'
        else: labels[c] = '균형 지역'
    df_gap_km['군집명'] = df_gap_km['군집'].map(labels)
    f.write(f"\n  [k=4 군집별 특성]\n")
    f.write(cluster_stats.to_string() + "\n")

    # ════════════════════════════════════════
    # 3-E. 시계열 분해 (수동 STL 대체)
    # ════════════════════════════════════════
    f.write("\n【3-E】 시계열 분해 — 일본인 방문자수\n")
    f.write("-" * 40 + "\n")
    visits = df_vis['조회기간 방문자 수'].values
    trend = pd.Series(visits).rolling(6, center=True).mean().values
    df_ts = df_vis.copy()
    df_ts['추세'] = trend
    df_ts['월'] = df_ts['날짜'].dt.month
    df_ts['계절성'] = df_ts['조회기간 방문자 수'] - df_ts['추세']
    month_effect = df_ts.groupby('월')['계절성'].mean()
    f.write("  [월별 계절 효과 (이동평균 잔차)]\n")
    for m, v in month_effect.items():
        f.write(f"  {m:2d}월: {v/1e6:+.2f}백만명\n")

# --- 시각화 (종합 및 개별 저장) ---
def save_sub(fig_obj, name):
    fig_obj.savefig(os.path.join(IMAGE_OUT, f"step3_{name}.png"), dpi=150, bbox_inches='tight')
    plt.close(fig_obj)

# 1. 교차분석 히트맵
fig, ax = plt.subplots(figsize=(8, 6))
sns.heatmap(cross2.div(cross2.sum(axis=1), axis=0) * 100, cmap='Blues', annot=True, fmt='.1f', ax=ax)
ax.set_title('성별 × 연령 교차분석 (%)')
save_sub(fig, "1_교차분석")

# 2. 성별 x 목적 누적바
fig, ax = plt.subplots(figsize=(8, 6))
cross1_norm = cross1.div(cross1.sum(axis=1), axis=0) * 100
cross1_norm.plot(kind='bar', stacked=True, ax=ax, color=['#1F3864','#2E75B6','#85B7EB','#D6E4F0','#E6F1FB'])
ax.set_title('성별 × 목적별 비중')
save_sub(fig, "2_성별목적비중")

# 3. 상관 히트맵
fig, ax = plt.subplots(figsize=(8, 6))
sns.heatmap(df_yt[['조회수_log','좋아요수_log','댓글수_log']].corr(), annot=True, cmap='RdBu_r', center=0, ax=ax)
ax.set_title('지표 간 상관관계 (Log)')
save_sub(fig, "3_상관관계")

# 4. 산점도
fig, ax = plt.subplots(figsize=(8, 6))
ax.scatter(df_yt['좋아요수_log'], df_yt['조회수_log'], alpha=0.3)
ax.set_title('좋아요 vs 조회수 (Log)')
save_sub(fig, "4_회귀산점도")

# 5. 예측 vs 실제
fig, ax = plt.subplots(figsize=(8, 6))
ax.scatter(y_test, y_pred, alpha=0.4)
ax.plot([y.min(), y.max()], [y.min(), y.max()], 'r--')
ax.set_title('실제값 vs 예측값')
save_sub(fig, "5_예측비교")

# 6. 잔차
fig, ax = plt.subplots(figsize=(8, 6))
ax.hist(y_test - y_pred, bins=30, color='#1E6B3C')
ax.set_title('회귀 잔차 분포')
save_sub(fig, "6_잔차분포")

# 7. 회귀계수
fig, ax = plt.subplots(figsize=(8, 6))
ax.barh(X_cols, reg.coef_)
ax.set_title('변수 중요도 (회귀계수)')
save_sub(fig, "7_회귀계수")

# 8. 군집 산점도
fig, ax = plt.subplots(figsize=(8, 6))
for lbl, grp in df_gap_km.groupby('군집명'):
    ax.scatter(grp['절대비율_글로벌'], grp['절대비율_일본'], label=lbl, alpha=0.6, color=cluster_colors.get(lbl, '#888780'))
ax.legend(); ax.set_title('지역 군집 분석')
save_sub(fig, "8_군집분석")

# 9. 엘보우
fig, ax = plt.subplots(figsize=(8, 6))
ax.plot(list(K_range), inertias, 'o-')
ax.set_title('K-Means 엘보우 분석')
save_sub(fig, "9_엘보우")

# 10. 시계열 추세
fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(df_vis['날짜'], df_vis['조회기간 방문자 수']/1e6, label='Original')
ax.plot(df_vis['날짜'], trend/1e6, label='Trend', linewidth=2)
ax.set_title('방문객 수 추세 (MA)'); ax.legend()
save_sub(fig, "10_시계열추세")

# 11. 계절성
fig, ax = plt.subplots(figsize=(10, 5))
ax.bar(month_effect.index, month_effect.values/1e6)
ax.set_title('월별 계절성 효과')
save_sub(fig, "11_계절성")

# 12. ACF
fig, ax = plt.subplots(figsize=(10, 5))
ax.bar(range(len(acf_vals)), acf_vals)
ax.set_title('자기상관함수 (ACF)')
save_sub(fig, "12_ACF")

print("Step 3 Refined Analysis complete with individual plots.")

df_gap_km.to_csv(os.path.join(OUT, "result_군집분석.csv"), index=False, encoding='utf-8-sig')
print("Step 3 Refined Analysis complete.")
