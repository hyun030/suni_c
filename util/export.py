# -*- coding: utf-8 -*-
import io
import os
import logging
import pandas as pd
from datetime import datetime

# --- 로그 대신 간단 출력 ---
def log_info(msg):
    print("[INFO]", msg)
def log_warn(msg):
    print("[WARN]", msg)
def log_error(msg):
    print("[ERROR]", msg)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak, Image as RLImage
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    PDF_AVAILABLE = True
except ImportError:
    log_warn("ReportLab 라이브러리 없음. PDF 생성 불가")
    PDF_AVAILABLE = False

try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    log_warn("Plotly 라이브러리 없음. 차트 생성 불가")
    PLOTLY_AVAILABLE = False

def create_enhanced_pdf_report(
    financial_data=None,
    news_data=None,
    insights=None,
    selected_charts=None,
    quarterly_df=None,
    show_footer=False,
    report_target="SK이노베이션 경영진",
    report_author="보고자 미기재"
):
    if not PDF_AVAILABLE:
        log_error("PDF 생성 불가능")
        return None

    import tempfile
    import re

    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception as e:
            log_warn(f"차트 PNG 변환 실패: {e}")
            return None

    def _clean_ai_text(raw):
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
            ('FONTNAME', (0,0), (-1,0), 'KoreanBold'),
            ('FONTNAME', (0,1), (-1,-1), 'KoreanSerif'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
        ]))
        return tbl

    # 폰트 경로 설정 (사용자 환경에 맞게 경로 수정 필요)
    base_dir = r"C:\Users\songo\OneDrive\써니C\예시"
    font_paths = {
        "Korean":      [os.path.join(base_dir, "nanum-gothic", "NanumGothic.ttf")],
        "KoreanBold":  [os.path.join(base_dir, "nanum-gothic", "NanumGothicBold.ttf")],
        "KoreanSerif": [os.path.join(base_dir, "nanum-myeongjo", "NanumMyeongjoBold.ttf")]
    }

    for font_name, paths in font_paths.items():
        registered = False
        for path in paths:
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont(font_name, path))
                    log_info(f"폰트 등록 성공: {font_name} ({path})")
                    registered = True
                    break
                except Exception as e:
                    log_warn(f"폰트 등록 실패: {font_name} ({path}) → {e}")
            else:
                log_warn(f"폰트 파일 없음: {path}")
        if not registered:
            log_warn(f"{font_name} 글꼴 등록 실패. 기본 폰트로 대체됩니다.")

    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'TITLE',
        fontName='KoreanBold' if 'KoreanBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold',
        fontSize=20,
        leading=34,
        spaceAfter=18
    )
    HEADING_STYLE = ParagraphStyle(
        'HEADING',
        fontName='KoreanBold' if 'KoreanBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold',
        fontSize=14,
        leading=23.8,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'BODY',
        fontName='KoreanSerif' if 'KoreanSerif' in pdfmetrics.getRegisteredFontNames() else 'Times-Roman',
        fontSize=12,
        leading=20.4,
        spaceAfter=6
    )

    buff = io.BytesIO()

    def _page_no(canvas, doc):
        canvas.setFont('Helvetica', 9)
        canvas.drawCentredString(A4[0] / 2, 18, f"- {canvas.getPageNumber()} -")

    doc = SimpleDocTemplate(
        buff, pagesize=A4,
        leftMargin=54, rightMargin=54, topMargin=54, bottomMargin=54
    )

    story = []

    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE
    ))
    story.append(Spacer(1, 12))

    # 재무분석 예시 (빈 데이터일 경우 스킵)
    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        display_cols = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
        df_disp = financial_data[display_cols].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0,0), (-1,0), 'KoreanBold'),
            ('FONTNAME', (0,1), (-1,-1), 'KoreanSerif'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 이후 차트, 뉴스, AI 인사이트 등 생략 가능

    try:
        doc.build(story, onFirstPage=_page_no, onLaterPages=_page_no)
        buff.seek(0)
        log_info("PDF 생성 성공")
        return buff.getvalue()
    except Exception as e:
        log_error(f"PDF 생성 실패: {e}")
        return None
