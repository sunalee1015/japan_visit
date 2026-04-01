# 🗾 일본인 관광객 지방분산 대시보드 (Japan Visit Dashboard)

본 프로젝트는 일본인 20~30대 여성 관광객의 한국 내 방문 및 소비 패턴을 분석하고, 서울에 집중된 관광 수요를 지방으로 분산시키기 위한 전략적 인사이트를 제공하는 **Streamlit 기반 인터랙티브 대시보드**입니다.

## 🚀 주요 기능

대시보드( `src/app.py` )는 다음 5가지 핵심 탭으로 구성되어 있습니다:

1.  **📊 KPI 요약**: 연령별/목적별 입국자 수, 월별 신용카드 소비 추이 등 주요 거버넌스 지표 시각화.
2.  **🗺️ 지역 분석**: 전국 광역자치구별 방문객 현황 및 서울 집중도(Gauge 차트) 진단.
3.  **💳 소비 패턴**: 업종별 소비 비중(Pie 차트) 및 서울 자치구별 글로벌 vs 일본인 소비 격차 분석.
4.  **📱 콘텐츠 트렌드**: SNS 해시태그 누적 추이, 유튜브 관광 키워드 및 업로드 트렌드 빅데이터 분석.
5.  **🎯 전략 인사이트**: 데이터 분석 결과에 기반한 3대 핵심 추진 전략(콘텐츠 성지순례, 웰니스, 지방 프로모션) 제안.

## 🛠️ 기술 스택

- **Frontend/Dashboard**: Streamlit
- **Data Analysis**: Pandas, Openpyxl
- **Visualization**: Plotly (Express & Graph Objects)
- **Environment**: uv (Python Package Manager)

## 📁 프로젝트 구조

```text
japan_visit/
├── .streamlit/          # Streamlit 설정 (테마 등)
├── data/                # 분석용 데이터 (※ 별도 준비 필요)
├── docs/                # 마케팅 및 EDA 최종 리포트 (PDF, Markdown)
├── images/              # EDA 과정에서 생성된 시각화 이미지
└── src/
    ├── app.py           # 메인 대시보드 실행 스크립트
    └── analysis/        # 데이터 분석 및 전처리에 사용된 보조 스크립트
```

## ⚙️ 설치 및 실행 방법

### 1. 가상환경 및 패키지 설치
`uv`를 사용하여 가상환경을 활성화하고 필요한 패키지를 설치합니다:

```powershell
pip install uv
uv venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
uv pip install -r requirements.txt
```

### 2. 데이터 파일 배치
본 저장소에는 보안 및 용량 문제로 원본 데이터(`Total_Dataset_analysis_20260321.xlsx`)가 포함되어 있지 않습니다.
- `japan_visit/data/` 폴더를 생성하고, 해당 엑셀 파일을 위치시켜 주십시오.

### 3. 대시보드 실행
```powershell
uv run streamlit run src/app.py
```

## 📄 주요 리포트
- [데이터 EDA 분석 리포트](docs/eda_report.md)
- [지방 분산 마케팅 전략 보고서 (v2)](docs/marketing_report_v2.md)
- [지방 분산 유도 구체화 전략](docs/dispersion_strategy.md)

---
**출처:** 한국관광데이터랩, 법무부 출입국 통계, YouTube Data API 기반 자체 분석
