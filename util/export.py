# -*- coding: utf-8 -*-
import io
import os
import tempfile
import pandas as pd
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False


def create_pdf_report(
    financial_data: pd.DataFrame | None = None,
    news_data: pd.DataFrame | None = None,
    insights: str | None = None,
    selected_charts: list | None = None,
    quarterly_df: pd.DataFrame | None = None,
    show_footer: bool = False,
    report_target: str = "SK이노베이션 경영진",
    report_author: str = "보고자 미기재"
) -> bytes:
    """
    한글 폰트 등록 및
    재무데이터 표, plotly 차트, 뉴스, AI 인사이트 포함 PDF 보고서 생성 함수
    
    반드시 font_path는 본인 PC 환경에 맞게 절대경로 또는 정확한 상대경로로 수정하세요.
    """

    # 1. 폰트 경로 설정 (본인 PC 환경에 맞게 변경)
    font_path = r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothic.ttf"
    font_bold_path = r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothicBold.ttf"
    font_serif_path = r"C:\Users\songo\OneDrive\써니C\예시\nanum-myeongjo\NanumMyeongjo.ttf"

    # 2. 폰트 등록 (기본 Helvetica fallback)
    fonts = {}
    try:
        pdfmetrics.registerFont(TTFont("NanumGothic", font_path))
        fonts["regular"] = "NanumGothic"
    except Exception as e:
        print("NanumGothic 등록 실패:", e)
        fonts["regular"] = "Helvetica"
    try:
        pdfmetrics.registerFont(TTFont("NanumGothicBold", font_bold_path))
        fonts["bold"] = "NanumGothicBold"
    except Exception as e:
        print("NanumGothicBold 등록 실패:", e)
        fonts["bold"] = "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("NanumMyeongjo", font_serif_path))
        fonts["serif"] = "NanumMyeongjo"
    except Exception as e:
        print("NanumMyeongjo 등록 실패:", e)
        fonts["serif"] = "Times-Roman"

    # 3. 스타일 정의
    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle('Title', fontName=fonts["bold"], fontSize=20, leading=30, spaceAfter=15)
    HEADING_STYLE = ParagraphStyle('Heading', fontName=fonts["bold"], fontSize=14, leading=23, textColor=colors.HexColor('#E31E24'), spaceBefore=16, spaceAfter=10)
    BODY_STYLE = ParagraphStyle('Body', fontName=fonts["serif"], fontSize=12, leading=18, spaceAfter=6)

    # 4. PDF 준비
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    story = []

    # 5. 표지
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 20))

    # 6. 재무분석 결과 표 추가
    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        # '_원시값' 컬럼 제외하고 출력
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), fonts["bold"]),
            ('FONTNAME', (0, 1), (-1, -1), fonts["serif"]),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 7. Plotly 차트 삽입 함수
    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception:
            return None

    # 8. plotly 차트 삽입
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

    # 9. 분기별 데이터가 있을 경우 꺾은선 그래프 자동 추가
    if PLOTLY_AVAILABLE and quarterly_df is not None and not quarterly_df.empty:
        if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
            fig_line = go.Figure()
            for comp in quarterly_df['회사'].dropna().unique():
                cdf = quarterly_df[quarterly_df['회사'] == comp]
                fig_line.add_trace(go.Scatter(x=cdf['분기'], y=cdf['영업이익률'], mode='lines+markers', name=comp))
            fig_line.update_layout(title="분기별 영업이익률 추이", xaxis_title="분기", yaxis_title="영업이익률(%)")
            img_bytes = _fig_to_png_bytes(fig_line)
            if img_bytes:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                story.append(Paragraph("2-1. 분기별 영업이익률 추이 (꺾은선)", BODY_STYLE))
                story.append(RLImage(tmp_path, width=500, height=280))
                story.append(Spacer(1, 16))
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

    # 10. 최신 뉴스 하이라이트
    if news_data is not None and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # 11. AI 인사이트
    if insights:
        story.append(PageBreak())
        story.append(Paragraph("4. AI 인사이트", HEADING_STYLE))
        for line in insights.splitlines():
            story.append(Paragraph(line, BODY_STYLE))

    # 12. 푸터
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # 13. PDF 빌드
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
