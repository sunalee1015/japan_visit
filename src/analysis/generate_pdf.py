import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import inch

# 경로 설정
BASE_DIR = r"C:\data_workspace\ICB6_Project2\japan_visit"
IMAGE_DIR = os.path.join(BASE_DIR, "images")
DOC_DIR = os.path.join(BASE_DIR, "docs")
PDF_PATH = os.path.join(DOC_DIR, "EDA_Visual_Report_Detailed.pdf")

# 폰트 등록 (Malgun Gothic)
FONT_PATH = r"C:\Windows\Fonts\malgun.ttf"
pdfmetrics.registerFont(TTFont('Malgun', FONT_PATH))

def generate_pdf():
    doc = SimpleDocTemplate(PDF_PATH, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
    styles = getSampleStyleSheet()
    
    # 커스텀 스타일 정의
    title_style = ParagraphStyle('TitleStyle', fontName='Malgun', fontSize=22, spaceAfter=30, alignment=1, textColor='#1F3864', bold=True)
    h2_style = ParagraphStyle('H2Style', fontName='Malgun', fontSize=16, spaceBefore=20, spaceAfter=10, textColor='#2E75B6', bold=True)
    h3_style = ParagraphStyle('H3Style', fontName='Malgun', fontSize=12, spaceBefore=15, spaceAfter=8, textColor='#1F3864', bold=True)
    body_style = ParagraphStyle('BodyStyle', fontName='Malgun', fontSize=10, leading=14, spaceAfter=10)
    insight_style = ParagraphStyle('InsightStyle', fontName='Malgun', fontSize=10, leading=14, leftIndent=20, textColor='#C00000', spaceAfter=20)

    story = []

    # 제목
    story.append(Paragraph("일본인 방한 관광객 EDA 심층 결과 보고서", title_style))
    story.append(Spacer(1, 0.2*inch))

    sections = [
        {
            "title": "1. 유튜브 성과 및 콘텐츠 분석",
            "points": [
                {"t": "[1-1] 조회수 분포 및 왜도 개선", "img": "step2_1_조회수분포.png", "txt": "조회수가 극단적 우편향(Long-tail) 분포를 보이며, 소수의 영상이 전체 조회수를 견인하고 있습니다."},
                {"t": "", "img": "step2_2_로그조회수.png", "txt": "로그 변환을 통해 분포의 정규성을 확보하였으며, 이는 통계적 추론 및 예측 모델 수립에 적합한 상태임을 의미합니다."},
                {"t": "[1-2] 지표 간 상관성 및 소셜 반응", "img": "step3_3_상관관계.png", "txt": "조회수와 좋아요수(r=0.76) 간에 강한 정적 상관관계가 확인되어, 시청자의 긍정적 반응이 전파력의 핵심임을 알 수 있습니다."},
                {"t": "", "img": "step2_4_소셜지표로그.png", "txt": "좋아요와 댓글수가 유사한 분포 패턴을 보이며, 적극적인 소셜 참여가 높은 조회수로 이어지는 경향이 뚜렷합니다."}
            ]
        },
        {
            "title": "2. 인구통계 및 방문 목적 분석",
            "points": [
                {"t": "[2-1] 성별 및 연령별 핵심 타겟", "img": "step2_5_성별규모.png", "txt": "여성이 남성에 비해 약 1.5배 이상 많이 방문하여 주요 소비 주체임을 보여줍니다."},
                {"t": "", "img": "step2_6_연령별규모.png", "txt": "21~30세 연령층이 가장 큰 비중을 차지하여, MZ 세대를 겨냥한 마케팅 전략이 유효함을 시사합니다."},
                {"t": "[2-2] 방문 목적 및 교차 분석", "img": "step2_7_목적비중.png", "txt": "입국객의 88%가 순수 '관광' 목적이며, 기타 목적 대비 압도적인 비중을 차지합니다."},
                {"t": "", "img": "step3_1_교차분석.png", "txt": "2030 여성층이 주를 이루는 '관광' 비중이 가장 높으며, 성별에 따른 연령대 분포 차이가 통계적으로 유의미합니다."}
            ]
        },
        {
            "title": "3. 지역별 방문 패턴 및 기회 지역",
            "points": [
                {"t": "[3-1] 방문 비율 갭 및 잠재 지역", "img": "step2_10_갭분포.png", "txt": "글로벌 방문객 대비 일본인의 방문이 적은 '갭' 지역들이 존재하여 새로운 시장 발굴의 기회를 제공합니다."},
                {"t": "", "img": "step2_11_갭상위.png", "txt": "창원, 신안 등 글로벌 대비 일본인 방문 비중이 낮은 지역들이 상위에 랭크되어 타겟팅 보완이 필요함을 보여줍니다."},
                {"t": "[3-2] K-Means를 활용한 지역 군집화", "img": "step3_8_군집분석.png", "txt": "지역들은 4개의 그룹으로 분류되며, 특히 '분산 타겟' 군집은 향후 일본 관광객의 방문 확장을 위한 핵심 전략 지역입니다."},
                {"t": "", "img": "step3_9_엘보우.png", "txt": "K=4 수준에서 클러스터링의 효율성이 최적화됨을 확인하여 군집 분류의 타당성을 입증했습니다."}
            ]
        },
        {
            "title": "4. 시계열 트렌드 및 예측성 분석",
            "points": [
                {"t": "[4-1] 방문객 추세 및 시계열 분석", "img": "step2_8_월별추이.png", "txt": "전반적으로 우상향 추세가 뚜렷하며, 특정 시점에서 관광 수요가 급증하는 패턴이 관찰됩니다."},
                {"t": "", "img": "step3_10_시계열추세.png", "txt": "변동성을 제거한 이동평균(Trend) 선이 꾸준한 상승 곡선을 그리고 있어 방한 일본 시장의 성장동력이 견고함을 입증합니다."},
                {"t": "[4-2] 계절성 및 자기상관", "img": "step3_11_계절성.png", "txt": "4월, 10월 등 봄/가을 시즌에 뚜렷한 양(+)의 계절 효과가 있으며, 하계 휴가 시즌보다 쾌적한 날씨의 방문 선호도가 높습니다."},
                {"t": "", "img": "step3_12_ACF.png", "txt": "자기상관 분석을 통해 과거 방문 정보가 미래 수요와 연관성이 높음을 확인하여 향후 방문객 예측의 가능성을 열어줍니다."}
            ]
        },
        {
            "title": "5. 예측 모델 및 품질 관리",
            "points": [
                {"t": "[5-1] 회귀 모델 성능 및 중요도", "img": "step3_5_예측비교.png", "txt": "실제값과 예측값이 선형적으로 정렬되어 있어 모델의 예측 신뢰도가 확보되었습니다."},
                {"t": "", "img": "step3_7_회귀계수.png", "txt": "'좋아요수'가 조회수 예측에 가장 큰 기여를 하는 핵심 지표임을 수치적으로 확인했습니다."},
                {"t": "[5-2] 데이터 분포 검증", "img": "step2_12_QQ플롯.png", "txt": "정규성 검정 결과 데이터의 질이 우수하며, 통계 모델 적용을 위한 기초 가정이 잘 충족되었습니다."},
                {"t": "", "img": "step3_6_잔차분포.png", "txt": "회귀 모델의 잔차가 0을 중심으로 대칭을 이루어 특정 편향 없이 학습이 잘 이루어졌음을 보여줍니다."}
            ]
        }
    ]

    for sec in sections:
        story.append(Paragraph(sec["title"], h2_style))
        for p in sec["points"]:
            if p["t"]:
                story.append(Paragraph(p["t"], h3_style))
            
            p_img = os.path.join(IMAGE_DIR, p["img"])
            if os.path.exists(p_img):
                story.append(Image(p_img, width=4*inch, height=3*inch))
                story.append(Spacer(1, 0.1*inch))
            
            story.append(Paragraph(f"• 인사이트: {p['txt']}", insight_style))
            
        story.append(PageBreak())

    # 빌드
    doc.build(story)
    print(f"Detailed PDF Report generated: {PDF_PATH}")

if __name__ == "__main__":
    generate_pdf()
