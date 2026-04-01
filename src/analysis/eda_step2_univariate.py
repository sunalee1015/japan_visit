import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
from scipy import stats
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

# 전처리 결과 로드
df_p12  = pd.read_csv(f"{OUT}/clean_p12_입국통계.csv", encoding='utf-8-sig')
df_tour = pd.read_csv(f"{OUT}/clean_p12_관광목적.csv", encoding='utf-8-sig')
df_yt   = pd.read_csv(f"{OUT}/clean_p44_유튜브.csv",  encoding='utf-8-sig')
df_vis  = pd.read_csv(f"{OUT}/clean_p27_방문자수.csv", encoding='utf-8-sig')
df_gap  = pd.read_csv(f"{OUT}/clean_p21p22_지역갭.csv", encoding='utf-8-sig')
df_vis['날짜'] = pd.to_datetime(df_vis['날짜'], errors='coerce')

# --- 데이터 계산 ---
views = df_yt['조회수']
views_log = df_yt['조회수_log']

stats_tbl = pd.DataFrame({
    '통계량': ['평균','중앙값','표준편차','최솟값','최댓값','왜도','첨도'],
    '원본': [f"{views.mean():,.0f}", f"{views.median():,.0f}",
             f"{views.std():,.0f}", f"{views.min():,.0f}",
             f"{views.max():,.0f}", f"{views.skew():.2f}", f"{views.kurt():.2f}"],
    '로그변환': [f"{views_log.mean():.2f}", f"{views_log.median():.2f}",
                f"{views_log.std():.2f}", f"{views_log.min():.2f}",
                f"{views_log.max():.2f}", f"{views_log.skew():.2f}", f"{views_log.kurt():.2f}"]
})

if len(views) >= 50:
    sample = views.sample(50, random_state=42)
    stat_sw, p_sw = stats.shapiro(sample)
    stat_sw_log, p_sw_log = stats.shapiro(np.log1p(sample))

pct = [10, 25, 50, 75, 90, 95, 99]
gender_cnt = df_p12.groupby('성별')['인원수'].sum()
total = gender_cnt.sum()
age_cnt = df_tour.groupby('연령')['인원수'].sum()
연령_순서 = ['20세 이하','21-30세','31-40세','41-50세','51-60세','61세 이상']
age_cnt = age_cnt.reindex([a for a in 연령_순서 if a in age_cnt.index])
total_tour = age_cnt.sum() if not age_cnt.empty else 1
목적_cnt = df_p12.groupby('목적')['인원수'].sum().sort_values(ascending=False)
total_p12 = 목적_cnt.sum()
kw_stats = df_yt.groupby('검색 키워드')['조회수'].agg(['count','mean','sum']).round(0)
kw_stats.columns = ['영상수','평균조회수','총조회수']
kw_stats = kw_stats.sort_values('평균조회수', ascending=False)
visit_stats = df_vis['조회기간 방문자 수'].describe()

# --- 보고서 저장 ---
with open(os.path.join(BASE_DIR, "docs", "eda_step2_report.txt"), "w", encoding="utf-8") as f:
    f.write("=" * 60 + "\n")
    f.write("STEP 2: 단변량 분석\n")
    f.write("=" * 60 + "\n")
    f.write("\n【2-A】 유튜브 조회수 단변량 분석\n")
    f.write("-" * 40 + "\n")
    f.write("  [기술통계]\n")
    f.write(stats_tbl.to_string(index=False) + "\n")
    if len(views) >= 50:
        f.write("\n  [Shapiro-Wilk 정규성 검정 (n=50 샘플)]\n")
        f.write(f"    원본:      W={stat_sw:.4f}, p={p_sw:.6f}\n")
        f.write(f"    로그변환:  W={stat_sw_log:.4f}, p={p_sw_log:.6f}\n")
    f.write("\n  [퍼센타일 분포]\n")
    for p in pct:
        f.write(f"    {p:3d}%ile: {np.percentile(views, p):>12,.0f}회\n")
    f.write("\n【2-B】 성별·연령·목적 빈도 분석\n")
    f.write("-" * 40 + "\n")
    f.write("  [성별 입국자 합계]\n")
    for g, v in gender_cnt.items():
        f.write(f"    {g}: {v:>10,.0f}명 ({v/total*100:.1f}%)\n")
    f.write("\n  [연령별 관광 입국자 합계]\n")
    for a, v in age_cnt.dropna().items():
        f.write(f"    {a}: {v:>10,.0f}명 ({v/total_tour*100:.1f}%)\n")
    f.write("\n  [목적별 입국자 합계]\n")
    for m, v in 목적_cnt.items():
        f.write(f"    {m}: {v:>10,.0f}명 ({v/total_p12*100:.1f}%)\n")
    f.write("\n  [유튜브 키워드별 평균 조회수 TOP10]\n")
    f.write(kw_stats.head(10).to_string() + "\n")
    f.write("\n【2-C】 일본인 방문자수 시계열 기술통계\n")
    f.write("-" * 40 + "\n")
    f.write(f"  평균: {visit_stats['mean']:>12,.0f}명/월\n")
    if not df_vis.empty:
        f.write(f"  최대: {visit_stats['max']:>12,.0f}명 ({df_vis.loc[df_vis['조회기간 방문자 수'].idxmax(),'기준년월']})\n")
        f.write(f"  최소: {visit_stats['min']:>12,.0f}명 ({df_vis.loc[df_vis['조회기간 방문자 수'].idxmin(),'기준년월']})\n")
        f.write(f"\n  증감률 통계:\n")
        f.write(f"  평균 증감률: {df_vis['전년대비 증감 비율'].mean():.1f}%\n")
    f.write("\n【2-D】 글로벌-일본 방문비율 갭 분포\n")
    f.write("-" * 40 + "\n")
    f.write(df_gap['갭'].describe().round(4).to_string() + "\n")

# --- 시각화 (종합 및 개별 저장) ---
def save_sub(fig_obj, name):
    fig_obj.savefig(os.path.join(IMAGE_OUT, f"step2_{name}.png"), dpi=150, bbox_inches='tight')
    plt.close(fig_obj)

# 1. 조회수 히스토그램
fig, ax = plt.subplots(figsize=(6, 4))
ax.hist(df_yt['조회수']/1e6, bins=40, color='#2E75B6', edgecolor='white', density=True)
ax_t = ax.twinx()
views_s = df_yt['조회수']/1e6
x_kde = np.linspace(0, views_s.max(), 200)
ax_t.plot(x_kde, stats.gaussian_kde(views_s)(x_kde), color='#C00000')
ax.set_title('조회수 분포 (백만)')
save_sub(fig, "1_조회수분포")

# 2. 로그 조회수
fig, ax = plt.subplots(figsize=(6, 4))
ax.hist(df_yt['조회수_log'], bins=40, color='#1E6B3C', edgecolor='white')
ax.axvline(df_yt['조회수_log'].mean(), color='r', linestyle='--')
ax.set_title('조회수 로그분포')
save_sub(fig, "2_로그조회수")

# 3. 키워드 박스플롯
fig, ax = plt.subplots(figsize=(8, 5))
kw_order = df_yt.groupby('검색 키워드')['조회수'].median().sort_values(ascending=False).index
kw_short = [str(k)[:6] for k in kw_order]
ax.boxplot([df_yt[df_yt['검색 키워드']==k]['조회수']/1e6 for k in kw_order], labels=kw_short, vert=False, patch_artist=True)
ax.set_title('키워드별 조회수')
save_sub(fig, "3_키워드박스")

# 4. 좋아요/댓글 로그
fig, ax = plt.subplots(figsize=(6, 4))
ax.hist(np.log1p(df_yt['좋아요수']), alpha=0.6, label='좋아요(log)')
ax.hist(np.log1p(df_yt['댓글수']), alpha=0.6, label='댓글(log)')
ax.legend(); ax.set_title('좋아요/댓글 로그분포')
save_sub(fig, "4_소셜지표로그")

# 5. 성별
fig, ax = plt.subplots(figsize=(6, 4))
g_data = df_p12.groupby('성별')['인원수'].sum() / 1e6
ax.bar(g_data.index, g_data.values, color=['#2E75B6', '#C00000'])
ax.set_title('성별 입국자 규모'); ax.set_ylabel('백만명')
save_sub(fig, "5_성별규모")

# 6. 연령별
fig, ax = plt.subplots(figsize=(8, 4))
age_data = age_cnt
ax.bar(range(len(age_data)), age_data.values/1e6, color='#2E75B6')
ax.set_xticks(range(len(age_data)))
ax.set_xticklabels([str(a) for a in age_data.index], rotation=20)
ax.set_title('연령별 관광객'); ax.set_ylabel('백만명')
save_sub(fig, "6_연령별규모")

# 7. 목적별
fig, ax = plt.subplots(figsize=(6, 6))
ax.pie(목적_cnt.values, labels=목적_cnt.index, autopct='%1.1f%%', startangle=90)
ax.set_title('방문 목적 비중')
save_sub(fig, "7_목적비중")

# 8. 월별 시계열
fig, ax = plt.subplots(figsize=(8, 4))
if not df_vis.empty:
    ax.plot(df_vis['날짜'], df_vis['조회기간 방문자 수']/1e6, marker='o')
ax.set_title('월별 방문자 추이'); ax.set_ylabel('백만명')
save_sub(fig, "8_월별추이")

# 9. YoY 증감
fig, ax = plt.subplots(figsize=(8, 4))
if not df_vis.empty:
    ax.bar(df_vis['날짜'], df_vis['전년대비 증감 비율'], color=['#C00000' if v < 0 else '#1E6B3C' for v in df_vis['전년대비 증감 비율']])
ax.set_title('YoY 증감률'); ax.set_ylabel('%')
save_sub(fig, "9_YoY증감")

# 10. 갭 분포
fig, ax = plt.subplots(figsize=(6, 4))
ax.hist(df_gap['갭'], bins=40, color='#4B0082')
ax.set_title('글로벌-일본 방문비율 갭')
save_sub(fig, "10_갭분포")

# 11. 갭 TOP10
fig, ax = plt.subplots(figsize=(8, 5))
top10 = df_gap.nlargest(10,'갭').copy()
top10['지역'] = top10['시도명'].str[:3] + ' ' + top10['시군구명']
ax.barh(top10['지역'], top10['갭'])
ax.invert_yaxis(); ax.set_title('갭 상위 지역')
save_sub(fig, "11_갭상위")

# 12. Q-Q
fig, ax = plt.subplots(figsize=(6, 4))
stats.probplot(df_yt['조회수_log'], plot=ax)
ax.set_title('Q-Q 플롯')
save_sub(fig, "12_QQ플롯")

print("EDA Step 2 individual plots complete.")
print("EDA Step 2 complete.")
