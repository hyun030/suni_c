# -*- coding: utf-8 -*-
import io
import os
import re
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

# plotly 관련
try:
    import plotly.graph_objects as go
    import plotly.express as px
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
    PDF 보고서 생성 (한글 포함, 메모리 버퍼 이미지 삽입)
    """

    # 1. 폰트 경로 (상대경로 또는 절대경로 본인 환경에 맞게 수정)
    base_dir = os.path.dirname(os.path.abspath(__file__))  # 이 파일 기준 폴더
    font_paths = {
        "Korean": os.path.join(base_dir, "fonts", "NanumGothic.ttf"),
        "KoreanBold": os.path.join(base_dir, "fonts", "NanumGothicBold.ttf"),
        "KoreanSerif": os.path.join(base_dir, "fonts", "NanumMyeongjo.ttf"),
    }
    fallback_fonts = {
        "Korean": "Helvetica",
        "KoreanBold": "Helvetica-Bold",
        "KoreanSerif": "Times-Roman",
    }

    # 2. 폰트 등록 (없는 파일은 fallback 처리)
    registered_fonts = {}
    for fam, path in font_paths.items():
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(fam, path))
                registered_fonts[fam] = fam
            except Exception as e:
                print(f"[폰트 등록 실패] {fam}: {e}")
                registered_fonts[fam] = fallback_fonts[fam]
        else:
            print(f"[폰트 파일 없음] {path}")
            registered_fonts[fam] = fallback_fonts[fam]

    # 3. 스타일 정의
    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'TITLE',
        fontName=registered_fonts.get("KoreanBold", "Helvetica-Bold"),
        fontSize=20,
        leading=34,
        spaceAfter=18
    )
    HEADING_STYLE = ParagraphStyle(
        'HEADING',
        fontName=registered_fonts.get("KoreanBold", "Helvetica-Bold"),
        fontSize=14,
        leading=23.8,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'BODY',
        fontName=registered_fonts.get("KoreanSerif", "Times-Roman"),
        fontSize=12,
        leading=20.4,
        spaceAfter=6
    )

    def _page_number(canvas, doc):
        canvas.setFont('Helvetica', 9)
        canvas.drawCentredString(A4[0] / 2, 18, f"- {canvas.getPageNumber()} -")

    # 4. plotly figure → io.BytesIO PNG 변환
    def _fig_to_image_buffer(fig, width=900, height=450):
        try:
            img_bytes = fig.to_image(format="png", width=width, height=height)
            return io.BytesIO(img_bytes)
        except Exception as e:
            print(f"[차트 이미지 변환 오류] {e}")
            return None

    # 5. PDF 문서 생성
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=54, rightMargin=54, topMargin=54, bottomMargin=54)
    story = []

    # 표지
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 12))

    # 1. 재무분석 결과
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), registered_fonts.get("KoreanBold", "Helvetica-Bold")),
            ('FONTNAME', (0, 1), (-1, -1), registered_fonts.get("KoreanSerif", "Times-Roman")),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 2. 시각화 차트
    if PLOTLY_AVAILABLE and selected_charts:
        story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
        for idx, fig in enumerate(selected_charts, start=1):
            img_buffer = _fig_to_image_buffer(fig)
            if img_buffer:
                img = RLImage(img_buffer, width=500, height=280)
                story.append(Paragraph(f"2-{idx}. 추가 차트", BODY_STYLE))
                story.append(img)
                story.append(Spacer(1, 16))
            else:
                story.append(Paragraph(f"2-{idx}. 차트 이미지 변환 실패", BODY_STYLE))

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
        # 간단하게 줄별 Paragraph 처리
        for line in insights.splitlines():
            story.append(Paragraph(line, BODY_STYLE))

    # 5. 푸터
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # PDF 빌드
    doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
    buffer.seek(0)
    return buffer.getvalue()


def create_excel_report(financial_data: pd.DataFrame) -> bytes:
    """
    Excel 보고서 생성 (financial_data 필수)
    """
    if financial_data is None or financial_data.empty:
        raise ValueError("financial_data가 없습니다.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        financial_data.to_excel(writer, index=False, sheet_name='재무데이터')
        writer.save()
    output.seek(0)
    return output.read()
