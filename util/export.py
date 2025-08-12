# -*- coding: utf-8 -*-
import os
import io
import tempfile
import re
import pandas as pd
from datetime import datetime
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
    Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

try:
    import plotly
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# ====================
# 폰트 등록 함수 (절대경로 하드코딩)
# ====================
def register_fonts_safe():
    # 절대경로로 수정 (필요시 환경에 맞게 변경)
    font_paths = {
        "Korean": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothic.ttf",
        "KoreanBold": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothicBold.ttf",
        "KoreanSerif": r"C:\Users\songo\OneDrive\써니C\예시\nanum-myeongjo\NanumMyeongjo.ttf"
    }

    registered_fonts = {}

    default_fonts = {
        "Korean": "Helvetica",
        "KoreanBold": "Helvetica-Bold",
        "KoreanSerif": "Times-Roman"
    }

    for font_name, path in font_paths.items():
        if os.path.exists(path):
            try:
                if font_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(font_name, path))
                    st.success(f"✅ {font_name} 폰트 등록 성공: {path}")
                else:
                    st.info(f"ℹ️ {font_name} 폰트 이미 등록됨")
                registered_fonts[font_name] = font_name
            except Exception as e:
                st.error(f"❌ {font_name} 폰트 등록 실패: {e}")
                registered_fonts[font_name] = default_fonts[font_name]
        else:
            st.warning(f"⚠️ {font_name} 폰트 파일을 찾을 수 없음: {path}")
            registered_fonts[font_name] = default_fonts[font_name]

    st.write("📝 최종 사용 폰트:", registered_fonts)
    return registered_fonts


# ====================
# AI 인사이트 텍스트 마크다운 제거 및 블록 분리 함수
# ====================
def clean_ai_text(raw: str):
    raw = re.sub(r'[*_#>~]', '', raw)
    blocks = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        if re.match(r'^\d+(\.\d+)*\s', line):
            blocks.append(('title', line))
        else:
            blocks.append(('body', line))
    return blocks

# ====================
# ASCII 표 → ReportLab Table 변환 함수
# ====================
def ascii_to_table(lines, registered_fonts, style_type="default"):
    if not lines:
        return None
    header = [c.strip() for c in lines[0].split('|') if c.strip()]
    if not header:
        return None
    data = []
    for ln in lines[2:]:
        cols = [c.strip() for c in ln.split('|') if c.strip()]
        if len(cols) == len(header):
            data.append(cols)
    if not data:
        return None

    # 스타일별 색상 차별화
    if style_type == "news":
        bg_color = colors.HexColor('#0066CC')
        row_bgcolors = [colors.whitesmoke, colors.HexColor('#F0F8FF')]
        text_color = colors.white
    elif style_type == "standalone":
        bg_color = colors.HexColor('#228B22')
        row_bgcolors = [colors.whitesmoke, colors.HexColor('#F0FFF0')]
        text_color = colors.white
    else:
        bg_color = colors.HexColor('#E31E24')
        row_bgcolors = [colors.whitesmoke, colors.HexColor('#F7F7F7')]
        text_color = colors.white

    tbl = Table([header] + data)
    tbl.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), bg_color),
        ('TEXTCOLOR', (0,0), (-1,0), text_color),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
        ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), row_bgcolors),
    ]))
    return tbl


# ====================
# Plotly Figure → PNG Bytes 변환 (임시파일 저장용)
# ====================
def _fig_to_png_bytes(fig, width=900, height=450):
    try:
        return fig.to_image(format="png", width=width, height=height)
    except Exception as e:
        st.warning(f"차트 이미지 변환 실패: {e}")
        return None


# ====================
# Excel 보고서 생성 함수
# ====================
def create_excel_report(financial_data=None, news_data=None, insights=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if financial_data is not None and not financial_data.empty:
            financial_data.to_excel(writer, sheet_name='재무분석', index=False)
        if news_data is not None and not news_data.empty:
            news_data.to_excel(writer, sheet_name='뉴스분석', index=False)
        if insights:
            pd.DataFrame({'AI 인사이트': [insights]}).to_excel(writer, sheet_name='AI인사이트', index=False)
    output.seek(0)
    return output.getvalue()


# ====================
# 향상된 PDF 보고서 생성 함수 (부가기능 + 뉴스섹션 완전 포함)
# ====================
def create_enhanced_pdf_report(
    financial_data=None,
    news_data=None,
    insights: str | None = None,
    selected_charts: list | None = None,
    quarterly_df: pd.DataFrame | None = None,
    show_footer: bool = False,
    report_target: str = "SK이노베이션 경영진",
    report_author: str = "보고자 미기재",
):
    registered_fonts = register_fonts_safe()
    styles = getSampleStyleSheet()

    TITLE_STYLE = ParagraphStyle(
        'Title',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=20,
        leading=30,
        spaceAfter=15,
        alignment=1,
    )
    HEADING_STYLE = ParagraphStyle(
        'Heading',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=14,
        leading=23,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10,
    )
    BODY_STYLE = ParagraphStyle(
        'Body',
        fontName=registered_fonts.get('KoreanSerif', 'Times-Roman'),
        fontSize=12,
        leading=18,
        spaceAfter=6,
    )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    story = []

    # 표지
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Spacer(1, 20))

    # 보고서 정보
    report_info = f"""
    <b>보고일자:</b> {datetime.now().strftime('%Y년 %m월 %d일')}<br/>
    <b>보고대상:</b> {report_target}<br/>
    <b>보고자:</b> {report_author}
    """
    story.append(Paragraph(report_info, BODY_STYLE))
    story.append(Spacer(1, 30))

    # 1. 재무분석 결과
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        cols_to_show = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
        df_disp = financial_data[cols_to_show].copy()
        max_rows_per_table = 25
        total_rows = len(df_disp)

        if total_rows <= max_rows_per_table:
            table_data = [df_disp.columns.tolist()] + df_disp.values.tolist()
            tbl = Table(table_data, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F2F2F2')),
                ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
                ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            story.append(tbl)
        else:
            for i in range(0, total_rows, max_rows_per_table):
                end_idx = min(i + max_rows_per_table, total_rows)
                chunk = df_disp.iloc[i:end_idx]
                if i > 0:
                    story.append(Paragraph(f"1-{i//max_rows_per_table + 1}. 재무분석 결과 (계속)", BODY_STYLE))
                table_data = [df_disp.columns.tolist()] + chunk.values.tolist()
                tbl = Table(table_data, repeatRows=1)
                tbl.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F2F2F2')),
                    ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
                    ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ]))
                story.append(tbl)
                story.append(Spacer(1, 8))

        story.append(Spacer(1, 20))

    # 2. 뉴스 및 벤치마킹 사례
    if news_data is not None and hasattr(news_data, "empty") and not news_data.empty:
        story.append(Paragraph("2. 뉴스 및 벤치마킹 사례", HEADING_STYLE))

        # 뉴스 요약 테이블 (최대 15개)
        max_news_rows = 15
        news_for_table = news_data.head(max_news_rows)
        news_columns = news_for_table.columns.tolist()
        table_data = [news_columns] + news_for_table.values.tolist()
        news_tbl = Table(table_data, repeatRows=1, colWidths=[120, 350, 80] if "date" in news_columns else None)
        news_tbl.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0066CC')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
            ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        story.append(news_tbl)
        story.append(Spacer(1, 15))

        # 뉴스 요약 텍스트 및 링크
        if "summary" in news_data.columns:
            news_summary = news_data.loc[0, "summary"]
            if isinstance(news_summary, str):
                news_summary = news_summary.strip()
                if news_summary:
                    story.append(Paragraph(f"<b>요약:</b> {news_summary}", BODY_STYLE))
                    story.append(Spacer(1, 15))

        # 링크 목록
        if "link" in news_data.columns:
            links = news_data["link"].dropna().unique().tolist()
            if links:
                story.append(Paragraph("관련 뉴스 링크:", BODY_STYLE))
                for link in links:
                    story.append(Paragraph(f'<a href="{link}">{link}</a>', BODY_STYLE))
                story.append(Spacer(1, 15))

    # 3. AI 인사이트 및 개선 전략
    if insights:
        story.append(Paragraph("3. AI 인사이트 및 개선 전략", HEADING_STYLE))
        blocks = clean_ai_text(insights)
        for block_type, content in blocks:
            if block_type == 'title':
                style = ParagraphStyle(
                    name="InsightTitle",
                    fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
                    fontSize=14,
                    leading=22,
                    textColor=colors.HexColor('#228B22'),
                    spaceBefore=14,
                    spaceAfter=8,
                )
                story.append(Paragraph(content, style))
            else:
                style = ParagraphStyle(
                    name="InsightBody",
                    fontName=registered_fonts.get('KoreanSerif', 'Times-Roman'),
                    fontSize=12,
                    leading=18,
                    spaceAfter=5,
                )
                story.append(Paragraph(content, style))
        story.append(Spacer(1, 20))

    # 4. 분기별 차트 삽입 (Plotly → 이미지)
    if PLOTLY_AVAILABLE and quarterly_df is not None and not quarterly_df.empty:
        story.append(Paragraph("4. 분기별 재무 추이", HEADING_STYLE))

        import plotly.graph_objects as go
        fig = go.Figure()

        # 예시 : 분기 컬럼 기준 선 그래프
        x_vals = quarterly_df.iloc[:, 0].tolist()
        for col in quarterly_df.columns[1:]:
            fig.add_trace(go.Scatter(x=x_vals, y=quarterly_df[col], mode='lines+markers', name=col))

        img_bytes = _fig_to_png_bytes(fig, width=900, height=350)
        if img_bytes:
            tmp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            try:
                tmp_img.write(img_bytes)
                tmp_img.close()
                story.append(RLImage(tmp_img.name, width=400, height=180))
            finally:
                try:
                    os.unlink(tmp_img.name)
                except:
                    pass
        story.append(Spacer(1, 20))

    # 5. 푸터(선택)
    if show_footer:
        footer_text = "본 보고서는 내부 참고용입니다."
        footer_style = ParagraphStyle(
            'Footer',
            fontName=registered_fonts.get('KoreanSerif', 'Times-Roman'),
            fontSize=9,
            leading=12,
            textColor=colors.grey,
            alignment=1,
        )
        story.append(Spacer(1, 30))
        story.append(Paragraph(footer_text, footer_style))

    # PDF 생성
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer.read()

