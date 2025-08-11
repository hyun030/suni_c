import os
import io
import pandas as pd
import datetime
import tempfile
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import streamlit as st

# 상단에 상대경로 기준 fonts 폴더 위치 지정
FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'fonts')

font_files = {
    "Korean": os.path.join(FONT_DIR, "NanumGothic.ttf"),
    "KoreanBold": os.path.join(FONT_DIR, "NanumGothicBold.ttf"),
    "KoreanSerif": os.path.join(FONT_DIR, "NanumMyeongjo.ttf"),
}

def register_fonts():
    """
    fonts 폴더 내 ttf 파일을 reportlab에 등록.
    실패해도 예외 무시하되 등록된 폰트 목록 확인 가능.
    """
    for font_name, font_path in font_files.items():
        try:
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont(font_name, font_path))
            else:
                st.warning(f"폰트 파일 경로를 찾을 수 없습니다: {font_path}")
        except Exception as e:
            st.warning(f"폰트 등록 실패 ({font_name}): {e}")

def create_enhanced_pdf_report(financial_data=None, news_data=None, insights=None, selected_charts=None):
    """
    한글 폰트 완벽 지원 PDF 생성 함수
    """

    register_fonts()  # 폰트 등록 먼저 수행

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
        canvas.drawCentredString(letter[0]/2, 18, f"- {canvas.getPageNumber()} -")

    doc = SimpleDocTemplate(buff, pagesize=letter,
                            leftMargin=54, rightMargin=54,
                            topMargin=54, bottomMargin=54)

    story = []

    story.append(Paragraph("SK에너지 경쟁사 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(f"보고일자: {datetime.datetime.now().strftime('%Y년 %m월 %d일')}", BODY_STYLE))
    story.append(Paragraph("보고대상: SK에너지 전략기획팀", BODY_STYLE))
    story.append(Paragraph("보고자: 전략기획팀", BODY_STYLE))
    story.append(Spacer(1, 12))

    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not c.endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0,0), (-1,0), 'KoreanBold'),
            ('FONTNAME', (0,1), (-1,-1), 'KoreanSerif'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'CENTER')
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    if news_data is not None and not news_data.empty:
        story.append(Paragraph("2. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    if insights:
        story.append(PageBreak())
        story.append(Paragraph("3. AI 인사이트", HEADING_STYLE))
        lines = str(insights).splitlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if line.startswith("##"):
                story.append(Paragraph(line.replace("##", "").strip(), HEADING_STYLE))
            else:
                story.append(Paragraph(line, BODY_STYLE))

    if selected_charts:
        try:
            import plotly.io as pio
            story.append(PageBreak())
            story.append(Paragraph("4. 시각화 차트", HEADING_STYLE))
            for fig in selected_charts:
                img_bytes = pio.to_image(fig, format="png", width=700, height=400)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                story.append(RLImage(tmp_path, width=500, height=280))
                story.append(Spacer(1, 16))
                os.unlink(tmp_path)
        except Exception as e:
            story.append(Paragraph(f"차트 삽입 오류: {e}", BODY_STYLE))

    doc.build(story, onFirstPage=_page_no, onLaterPages=_page_no)
    buff.seek(0)
    return buff.getvalue()


def create_excel_report(financial_data=None, news_data=None, insights=None):
    """
    Excel 보고서 생성
    """
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if financial_data is not None and not financial_data.empty:
                clean_financial = financial_data[[col for col in financial_data.columns if not col.endswith('_원시값')]]
                clean_financial.to_excel(writer, sheet_name='재무분석', index=False)
            if news_data is not None and not news_data.empty:
                news_data.to_excel(writer, sheet_name='뉴스분석', index=False)
            if insights:
                insight_df = pd.DataFrame({'구분': ['AI 인사이트'], '내용': [str(insights)]})
                insight_df.to_excel(writer, sheet_name='AI인사이트', index=False)
        output.seek(0)
        return output.getvalue()
    except Exception as e:
        st.error(f"Excel 생성 오류: {e}")
        return None
