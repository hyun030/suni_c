import os
import io
import tempfile
import pandas as pd
from datetime import datetime

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

PDF_AVAILABLE = True
PLOTLY_AVAILABLE = True  # 이건 Plotly 설치 상황에 맞게 조절

# 프로젝트 내 폰트 상대경로 설정 (fonts 폴더 기준)
FONT_DIR = os.path.join(os.path.dirname(__file__), 'fonts')

font_files = {
    "Korean": os.path.join(FONT_DIR, "NanumGothic.ttf"),
    "KoreanBold": os.path.join(FONT_DIR, "NanumGothicBold.ttf"),
    "KoreanSerif": os.path.join(FONT_DIR, "NanumMyeongjo.ttf"),
}

def register_fonts():
    """한글폰트 등록 - 실패해도 무시, 등록 여부 반환"""
    success = True
    try:
        for font_name, path in font_files.items():
            if os.path.exists(path):
                pdfmetrics.registerFont(TTFont(font_name, path))
            else:
                success = False
                print(f"폰트 파일 없음: {path}")
    except Exception as e:
        print(f"폰트 등록 에러: {e}")
        success = False
    return success

def create_enhanced_pdf_report(
        financial_data=None,
        news_data=None,
        insights:str|None=None,
        selected_charts:list|None=None
):
    if not PDF_AVAILABLE:
        import streamlit as st
        st.error("reportlab 라이브러리가 필요합니다.")
        return None

    # 폰트 등록 시도
    register_fonts()
    registered_fonts = pdfmetrics.getRegisteredFontNames()

    # 스타일 설정
    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'TITLE',
        fontName='KoreanBold' if 'KoreanBold' in registered_fonts else 'Helvetica-Bold',
        fontSize=20,
        leading=34,
        spaceAfter=18
    )
    HEADING_STYLE = ParagraphStyle(
        'HEADING',
        fontName='KoreanBold' if 'KoreanBold' in registered_fonts else 'Helvetica-Bold',
        fontSize=14,
        leading=23.8,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'BODY',
        fontName='KoreanSerif' if 'KoreanSerif' in registered_fonts else 'Times-Roman',
        fontSize=12,
        leading=20.4,
        spaceAfter=6
    )

    # 쪽번호 함수
    def _page_no(canvas, doc):
        canvas.setFont('Helvetica', 9)
        canvas.drawCentredString(letter[0]/2, 18, f"- {canvas.getPageNumber()} -")

    buff = io.BytesIO()
    doc = SimpleDocTemplate(buff, pagesize=letter,
                            leftMargin=54, rightMargin=54,
                            topMargin=54, bottomMargin=54)

    story = []

    # 제목 & 메타
    story.append(Paragraph("SK에너지 경쟁사 분석 보고서", TITLE_STYLE))
    story.append(Paragraph("보고일자: 2024년 10월 26일    보고대상: SK에너지 전략기획팀    보고자: 전략기획팀", BODY_STYLE))
    story.append(Spacer(1, 12))

    # 재무 데이터 표 (원시값 컬럼 제외)
    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not c.endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0,0), (-1,0), 'KoreanBold' if 'KoreanBold' in registered_fonts else 'Helvetica-Bold'),
            ('FONTNAME', (0,1), (-1,-1), 'KoreanSerif' if 'KoreanSerif' in registered_fonts else 'Times-Roman'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 뉴스 요약
    if news_data is not None and not news_data.empty:
        story.append(Paragraph("2. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # AI 인사이트 텍스트 처리 (마크다운 기호 제거 & 표 변환 가능)
    if insights:
        import re
        story.append(PageBreak())
        story.append(Paragraph("3. AI 인사이트", HEADING_STYLE))

        def _clean_ai_text(raw:str):
            raw = re.sub(r'[*_`#>~]', '', raw)
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

        def _ascii_block_to_table(lines):
            header = [c.strip() for c in lines[0].split('|') if c.strip()]
            data = []
            for ln in lines[2:]:
                cols = [c.strip() for c in ln.split('|') if c.strip()]
                if len(cols) == len(header):
                    data.append(cols)
            if not data:
                return None
            tbl = Table([header] + data)
            tbl.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E31E24')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'KoreanBold' if 'KoreanBold' in registered_fonts else 'Helvetica-Bold'),
                ('FONTNAME', (0,1), (-1,-1), 'KoreanSerif' if 'KoreanSerif' in registered_fonts else 'Times-Roman'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
            ]))
            return tbl

        blocks = _clean_ai_text(insights)
        ascii_buf = []
        for typ, ln in blocks:
            if '|' in ln:
                ascii_buf.append(ln)
                continue
            if ascii_buf:
                tbl = _ascii_block_to_table(ascii_buf)
                if tbl: story.append(tbl)
                story.append(Spacer(1,12))
                ascii_buf.clear()
            if typ == 'title':
                story.append(Paragraph(f"<b>{ln}</b>", BODY_STYLE))
            else:
                story.append(Paragraph(ln, BODY_STYLE))
        if ascii_buf:
            tbl = _ascii_block_to_table(ascii_buf)
            if tbl: story.append(tbl)

    # 차트 이미지 삽입 (Plotly)
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

    # 푸터 문구
    story.append(Spacer(1, 50))
    story.append(Paragraph("🔗 본 보고서는 SK에너지 경쟁사 분석 대시보드에서 자동 생성되었습니다.", BODY_STYLE))
    story.append(Paragraph("📊 실제 DART API 데이터 + Google Gemini AI 분석 기반", BODY_STYLE))

    # PDF 빌드 (쪽번호 추가)
    doc.build(story, onFirstPage=_page_no, onLaterPages=_page_no)
    buff.seek(0)
    return buff.getvalue()
