# -*- coding: utf-8 -*-
import io
import os
import re
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


def register_fonts():
    """
    폰트 등록 (본인 PC 절대 경로로 바꿔주세요)
    """
    font_paths = {
        "Korean": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothic.ttf",
        "KoreanBold": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothicBold.ttf",
        "KoreanSerif": r"C:\Users\songo\OneDrive\써니C\예시\nanum-myeongjo\NanumMyeongjo.ttf",
    }
    fallback_fonts = {
        "Korean": "Helvetica",
        "KoreanBold": "Helvetica-Bold",
        "KoreanSerif": "Times-Roman",
    }

    registered_fonts = {}

    for name, path in font_paths.items():
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(name, path))
                registered_fonts[name] = name
            except Exception as e:
                print(f"폰트 등록 실패: {name} - {e}")
                registered_fonts[name] = fallback_fonts[name]
        else:
            print(f"폰트 파일 없음: {path}")
            registered_fonts[name] = fallback_fonts[name]

    return registered_fonts


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
    한글 폰트 등록 후 PDF 보고서 생성 함수
    """
    fonts = register_fonts()

    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'Title', fontName=fonts["KoreanBold"], fontSize=20, leading=30, spaceAfter=15)
    HEADING_STYLE = ParagraphStyle(
        'Heading', fontName=fonts["KoreanBold"], fontSize=14, leading=23,
        textColor=colors.HexColor('#E31E24'), spaceBefore=16, spaceAfter=10)
    BODY_STYLE = ParagraphStyle(
        'Body', fontName=fonts["KoreanSerif"], fontSize=12, leading=18, spaceAfter=6)

    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception:
            return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    story = []

    # 표지
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 20))

    # 1. 재무분석 결과 테이블
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), fonts["KoreanBold"]),
            ('FONTNAME', (0, 1), (-1, -1), fonts["KoreanSerif"]),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 2. 시각화 차트 (Plotly)
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

    # 푸터
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def create_excel_report(financial_data: pd.DataFrame) -> bytes:
    """
    간단한 Excel 보고서 생성
    """
    if financial_data is None or financial_data.empty:
        raise ValueError("financial_data가 없습니다.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        financial_data.to_excel(writer, index=False, sheet_name='재무데이터')
        writer.save()
    output.seek(0)
    return output.read()
