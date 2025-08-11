# -*- coding: utf-8 -*-
import io
import os
import pandas as pd
from datetime import datetime

# reportlab 관련 임포트
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Plotly 사용 가능 여부 체크
try:
    import plotly
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

def create_enhanced_pdf_report(
    financial_data=None,
    news_data=None,
    insights: str | None = None,
    selected_charts: list | None = None,
    quarterly_df: pd.DataFrame | None = None,
    show_footer: bool = False,
    report_target: str = "SK이노베이션 경영진",
    report_author: str = "보고자 미기재"
):
    """
    Streamlit용 PDF 생성 함수 (로컬 환경용 절대경로 폰트 등록 포함)
    """

    # 1. 폰트 절대경로 (본인 로컬 환경에 맞게 반드시 수정하세요)
    base_dir = r"C:\Users\songo\OneDrive\써니C\예시"
    font_paths = {
        "Korean": os.path.join(base_dir, "nanum-gothic", "NanumGothic.ttf"),
        "KoreanBold": os.path.join(base_dir, "nanum-gothic", "NanumGothicBold.ttf"),
        "KoreanSerif": os.path.join(base_dir, "nanum-myeongjo", "NanumMyeongjo.ttf")
    }

    # 2. 폰트 등록
    for font_name, font_path in font_paths.items():
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
            except Exception as e:
                print(f"폰트 등록 실패: {font_name} - {e}")
        else:
            print(f"폰트 파일 없음: {font_path}")

    # 3. 스타일 정의
    styles = getSampleStyleSheet()
    styles['Normal'].fontName = 'Korean'
    TITLE_STYLE = ParagraphStyle('Title', fontName='KoreanBold', fontSize=20, leading=30, spaceAfter=15)
    HEADING_STYLE = ParagraphStyle('Heading', fontName='KoreanBold', fontSize=14, leading=23, textColor=colors.HexColor('#E31E24'), spaceBefore=16, spaceAfter=10)
    BODY_STYLE = ParagraphStyle('Body', fontName='KoreanSerif', fontSize=12, leading=18, spaceAfter=6)

    # 4. PDF 문서 생성 준비
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)

    story = []

    # 표지 및 메타 정보
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 20))

    # 1. 재무분석 결과 (예: 데이터프레임 표 출력)
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), 'KoreanBold'),
            ('FONTNAME', (0, 1), (-1, -1), 'KoreanSerif'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 2. 시각화 차트 삽입 (Plotly → PNG 변환 필요)
    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception:
            return None

    import tempfile

    charts_added = False
    if PLOTLY_AVAILABLE and selected_charts:
        story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
        for idx, fig in enumerate(selected_charts, start=1):
            img_bytes = _fig_to_png_bytes(fig)
            if img_bytes:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                story.append(Paragraph(f"2-{idx}. 추가 차트", BODY_STYLE))
                story.append(RLImage(tmp_path, width=500, height=280))
                story.append(Spacer(1, 16))
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass
                charts_added = True

    # 3. 최신 뉴스 하이라이트
    if news_data is not None and hasattr(news_data, "empty") and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # 4. AI 인사이트
    if insights:
        story.append(PageBreak())
        story.append(Paragraph("4. AI 인사이트", HEADING_STYLE))
        for line in insights.splitlines():
            story.append(Paragraph(line, BODY_STYLE))

    # 5. 선택적 푸터
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # PDF 빌드
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
