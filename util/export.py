# -*- coding: utf-8 -*-
import io
import os
import re
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Plotly 필요 시 체크
try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False


def create_enhanced_pdf_report(
        financial_data=None,
        news_data=None,
        insights: str | None = None,
        selected_charts: list | None = None
):
    """
    PDF 보고서 생성 (폰트 깨짐 방지 + 쪽번호 포함)
    
    - financial_data : pandas DataFrame (재무데이터)
    - news_data      : pandas DataFrame (뉴스데이터)
    - insights       : str (AI 인사이트 텍스트)
    - selected_charts: list of Plotly figures
    
    반환 : PDF 바이너리(bytes) 또는 None (오류 시)
    """

    # --- 1. 내부 헬퍼 함수들 ---

    def _clean_ai_text(raw: str) -> list[tuple[str, str]]:
        """
        AI 인사이트 텍스트 마크다운 기호 제거 후
        ('title'|'body', text) 리스트 반환
        """
        raw = re.sub(r'[*_`#>~\-]', '', raw)  # 일부 마크다운 심볼 제거
        blocks = []
        for ln in raw.splitlines():
            ln = ln.strip()
            if not ln:
                continue
            if re.match(r'^\d+(\.\d+)*\s', ln):  # 1. 또는 1.1 처럼 시작
                blocks.append(('title', ln))
            else:
                blocks.append(('body', ln))
        return blocks

    def _ascii_block_to_table(lines: list[str]):
        """
        ASCII 테이블 파이프(|) 문자열 리스트를
        ReportLab Table 객체로 변환
        """
        if not lines or len(lines) < 3:
            return None
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
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E31E24')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'KoreanBold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Korean'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
        ]))
        return tbl

    # --- 2. 폰트 등록 (여러 경로 시도) ---
    font_paths = {
        "Korean": [
            r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothic.ttf",
            "./fonts/NanumGothic.ttf",
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "/System/Library/Fonts/AppleSDGothicNeo.ttc",
            "C:/Windows/Fonts/malgun.ttf"
        ],
        "KoreanBold": [
            r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothicBold.ttf",
            "./fonts/NanumGothicBold.ttf",
            "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
            "/System/Library/Fonts/AppleSDGothicNeo.ttc",
            "C:/Windows/Fonts/malgunbd.ttf"
        ],
        "KoreanSerif": [
            r"C:\Users\songo\OneDrive\써니C\예시\nanum-myeongjo\NanumMyeongjoBold.ttf",
            "./fonts/NanumMyeongjoBold.ttf",
            "/usr/share/fonts/truetype/nanum/NanumMyeongjoBold.ttf",
            "C:/Windows/Fonts/batang.ttc",
            "/System/Library/Fonts/Supplemental/Batang.ttc"
        ]
    }

    registered_fonts = {}

    for font_name, paths in font_paths.items():
        for path in paths:
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont(font_name, path))
                    registered_fonts[font_name] = font_name
                    break
                except Exception:
                    # 이미 등록됐거나 오류면 무시하고 다음 시도
                    continue

    # 폰트가 없으면 기본 폰트로 대체
    korean_font = registered_fonts.get("Korean", "Helvetica")
    korean_bold_font = registered_fonts.get("KoreanBold", "Helvetica-Bold")
    korean_serif_font = registered_fonts.get("KoreanSerif", "Times-Roman")

    # --- 3. 스타일 정의 ---
    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'TITLE',
        fontName=korean_bold_font,
        fontSize=20,
        leading=34,
        spaceAfter=18,
        textColor=colors.HexColor('#E31E24')
    )
    HEADING_STYLE = ParagraphStyle(
        'HEADING',
        fontName=korean_bold_font,
        fontSize=14,
        leading=23.8,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'BODY',
        fontName=korean_serif_font,
        fontSize=12,
        leading=20.4,
        spaceAfter=6
    )

    # --- 4. PDF 문서 준비 ---
    buff = io.BytesIO()

    def _page_no(canvas, doc):
        canvas.setFont('Helvetica', 9)
        page_num = canvas.getPageNumber()
        canvas.drawCentredString(letter[0] / 2, 18, f"- {page_num} -")

    doc = SimpleDocTemplate(buff, pagesize=letter,
                            leftMargin=54, rightMargin=54,
                            topMargin=54, bottomMargin=54)

    story = []

    # --- 5. 제목 및 메타 ---
    story.append(Paragraph("SK에너지 경쟁사 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    "
        "보고대상: SK에너지 전략기획팀    보고자: 전략기획팀", BODY_STYLE))
    story.append(Spacer(1, 12))

    # --- 6. 재무분석 ---
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
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # --- 7. 뉴스 요약 ---
    if news_data is not None and not news_data.empty:
        story.append(Paragraph("2. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # --- 8. AI 인사이트 ---
    if insights:
        story.append(PageBreak())
        story.append(Paragraph("3. AI 인사이트", HEADING_STYLE))

        blocks = _clean_ai_text(insights)
        ascii_buf = []
        for typ, ln in blocks:
            if '|' in ln:  # ASCII 표 후보
                ascii_buf.append(ln)
                continue

            # 버퍼 비우기
            if ascii_buf:
                tbl = _ascii_block_to_table(ascii_buf)
                if tbl:
                    story.append(tbl)
                story.append(Spacer(1, 12))
                ascii_buf.clear()

            # 제목과 본문 구분
            if typ == 'title':
                story.append(Paragraph(f"<b>{ln}</b>", BODY_STYLE))
            else:
                story.append(Paragraph(ln, BODY_STYLE))

        # 남은 표 비우기
        if ascii_buf:
            tbl = _ascii_block_to_table(ascii_buf)
            if tbl:
                story.append(tbl)

    # --- 9. 차트 삽입 ---
    if selected_charts and PLOTLY_AVAILABLE:
        story.append(PageBreak())
        story.append(Paragraph("4. 시각화 차트", HEADING_STYLE))
        for fig in selected_charts:
            try:
                img_bytes = fig.to_image(format="png", width=700, height=400)
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                story.append(RLImage(tmp_path, width=500, height=280))
                story.append(Spacer(1, 16))
                os.unlink(tmp_path)
            except Exception as e:
                story.append(Paragraph(f"차트 삽입 오류: {e}", BODY_STYLE))

    # --- 10. 푸터 ---
    story.append(Spacer(1, 50))
    story.append(Paragraph("🔗 본 보고서는 SK에너지 경쟁사 분석 대시보드에서 자동 생성되었습니다.", BODY_STYLE))
    story.append(Paragraph("📊 실제 DART API 데이터 + Google Gemini AI 분석 기반", BODY_STYLE))

    # --- 11. PDF 빌드 (쪽번호 포함) ---
    try:
        doc.build(story, onFirstPage=_page_no, onLaterPages=_page_no)
        buff.seek(0)
        return buff.getvalue()
    except Exception as e:
        st.error(f"PDF 생성 오류: {e}")
        return None

