import os
import io
import tempfile
import pandas as pd
import streamlit as st
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 폰트 폴더 경로 (상대경로)
FONT_DIR = "./fonts"

# 폰트 파일 절대경로 (상대경로 기반)
font_files = {
    "Korean": os.path.join(FONT_DIR, "NanumGothic.ttf"),
    "KoreanBold": os.path.join(FONT_DIR, "NanumGothicBold.ttf"),
    "KoreanSerif": os.path.join(FONT_DIR, "NanumMyeongjo.ttf"),
}

# 폰트 등록 함수
def register_fonts():
    for font_name, font_path in font_files.items():
        if os.path.exists(font_path):
            try:
                # 중복 등록 대비 예외 처리
                if font_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(font_name, font_path))
                    st.write(f"폰트 등록 성공: {font_name} ({font_path})")
            except Exception as e:
                st.warning(f"폰트 등록 실패: {font_name} ({font_path}) - {e}")
        else:
            st.warning(f"폰트 파일 없음: {font_path}")

# AI 인사이트 텍스트 마크다운 간단 클린 함수
def _clean_ai_text(raw: str):
    import re
    raw = re.sub(r'[*_`#>~]', '', raw)  # 마크다운 특수문자 제거
    blocks = []
    for ln in raw.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        if re.match(r'^\d+(\.\d+)*\s', ln):
            blocks.append(('title', ln))
        else:
            blocks.append(('body', ln))
    return blocks

# ASCII 표 → ReportLab Table 변환
def _ascii_block_to_table(lines):
    header = [c.strip() for c in lines[0].split('|') if c.strip()]
    data = []
    for ln in lines[2:]:  # 구분선 제외
        cols = [c.strip() for c in ln.split('|') if c.strip()]
        if len(cols) == len(header):
            data.append(cols)
    if not data:
        return None
    tbl = Table([header] + data)
    tbl.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND',(0,0),(-1,0), colors.HexColor('#E31E24')),
        ('TEXTCOLOR',(0,0),(-1,0), colors.white),
        ('ALIGN',(0,0),(-1,-1), 'CENTER'),
        ('FONTNAME',(0,0),(-1,0), 'KoreanBold'),
        ('FONTNAME',(0,1),(-1,-1), 'KoreanSerif'),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('ROWBACKGROUNDS',(0,1),(-1,-1), [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
    ]))
    return tbl

def create_enhanced_pdf_report(
    financial_data=None,
    news_data=None,
    insights: str | None = None,
    selected_charts: list | None = None
):
    # 폰트 등록
    register_fonts()

    # 폰트 등록 확인 (fallback 처리)
    available_fonts = pdfmetrics.getRegisteredFontNames()
    korean_font = "Korean" if "Korean" in available_fonts else "Helvetica"
    korean_bold_font = "KoreanBold" if "KoreanBold" in available_fonts else "Helvetica-Bold"
    korean_serif_font = "KoreanSerif" if "KoreanSerif" in available_fonts else "Times-Roman"

    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        "TITLE",
        fontName=korean_bold_font,
        fontSize=20,
        leading=34,
        spaceAfter=18
    )
    HEADING_STYLE = ParagraphStyle(
        "HEADING",
        fontName=korean_bold_font,
        fontSize=14,
        leading=23.8,
        textColor=colors.HexColor("#E31E24"),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        "BODY",
        fontName=korean_serif_font,
        fontSize=12,
        leading=20.4,
        spaceAfter=6
    )

    buff = io.BytesIO()

    def _page_no(canvas, doc):
        canvas.setFont("Helvetica", 9)
        canvas.drawCentredString(letter[0] / 2, 18, f"- {canvas.getPageNumber()} -")

    doc = SimpleDocTemplate(buff, pagesize=letter,
                            leftMargin=54, rightMargin=54,
                            topMargin=54, bottomMargin=54)

    story = []

    # 제목 및 메타 정보
    story.append(Paragraph("SK에너지 경쟁사 분석 보고서", TITLE_STYLE))
    story.append(Paragraph("보고일자: 2024년 10월 26일    보고대상: SK에너지 전략기획팀    보고자: 전략기획팀", BODY_STYLE))
    story.append(Spacer(1, 12))

    # 재무분석 표
    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not c.endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), korean_bold_font),
            ('FONTNAME', (0, 1), (-1, -1), korean_serif_font),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER')
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 뉴스 요약
    if news_data is not None and not news_data.empty:
        story.append(Paragraph("2. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # AI 인사이트
    if insights:
        story.append(PageBreak())
        story.append(Paragraph("3. AI 인사이트", HEADING_STYLE))
        blocks = _clean_ai_text(insights)
        ascii_buf = []
        for typ, ln in blocks:
            if "|" in ln:
                ascii_buf.append(ln)
                continue
            if ascii_buf:
                tbl = _ascii_block_to_table(ascii_buf)
                if tbl:
                    story.append(tbl)
                story.append(Spacer(1, 12))
                ascii_buf.clear()
            if typ == "title":
                story.append(Paragraph(f"<b>{ln}</b>", BODY_STYLE))
            else:
                story.append(Paragraph(ln, BODY_STYLE))
        if ascii_buf:
            tbl = _ascii_block_to_table(ascii_buf)
            if tbl:
                story.append(tbl)

    # 선택된 Plotly 차트 이미지 삽입
    if selected_charts:
        try:
            import plotly
            if not selected_charts:
                selected_charts = []
        except ImportError:
            selected_charts = []

    if selected_charts:
        story.append(PageBreak())
        story.append(Paragraph("4. 시각화 차트", HEADING_STYLE))
        for fig in selected_charts:
            try:
                img_bytes = fig.to_image(format="png", width=700, height=400)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                story.append(RLImage(tmp_path, width=500, height=280))
                story.append(Spacer(1, 16))
                os.unlink(tmp_path)
            except Exception as e:
                story.append(Paragraph(f"차트 삽입 오류: {e}", BODY_STYLE))

    # 빌드
    doc.build(story, onFirstPage=_page_no, onLaterPages=_page_no)
    buff.seek(0)
    return buff.getvalue()
