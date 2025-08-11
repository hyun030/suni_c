# -*- coding: utf-8 -*-
import io
import os
import tempfile
import pandas as pd
from datetime import datetime

# reportlab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# plotly (optional)
try:
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

def create_enhanced_pdf_report(
    financial_data: pd.DataFrame = None,
    news_data: pd.DataFrame = None,
    insights: str = None,
    selected_charts: list = None,
    show_footer: bool = False,
    report_target: str = "SK이노베이션 경영진",
    report_author: str = "보고자 미기재"
) -> bytes:
    """
    PDF 리포트 생성
    - font_paths: 반드시 본인 PC 절대 경로로 변경 필수
    - financial_data, news_data: pandas DataFrame (필요없는 컬럼 미리 제거)
    - selected_charts: plotly Figure 리스트
    """

    # 폰트 경로 반드시 본인 PC에 맞게 수정하세요
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

    for k, path in font_paths.items():
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(k, path))
                font_paths[k] = k
            except Exception:
                font_paths[k] = fallback_fonts[k]
        else:
            font_paths[k] = fallback_fonts[k]

    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'Title',
        fontName=font_paths["KoreanBold"],
        fontSize=20,
        leading=30,
        spaceAfter=15
    )
    HEADING_STYLE = ParagraphStyle(
        'Heading',
        fontName=font_paths["KoreanBold"],
        fontSize=14,
        leading=23,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'Body',
        fontName=font_paths["KoreanSerif"],
        fontSize=12,
        leading=18,
        spaceAfter=6
    )

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

    # 재무분석 결과
    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), font_paths["KoreanBold"]),
            ('FONTNAME', (0, 1), (-1, -1), font_paths["KoreanSerif"]),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 시각화 차트
    if PLOTLY_AVAILABLE and selected_charts:
        story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
        for idx, fig in enumerate(selected_charts, start=1):
            img_bytes = _fig_to_png_bytes(fig)
            if img_bytes:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                story.append(Paragraph(f"2-{idx}. 추가 차트", BODY_STYLE))
                story.append(RLImage(tmp_path, width=500, height=280))
                story.append(Spacer(1, 16))
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

    # 최신 뉴스
    if news_data is not None and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # AI 인사이트
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
    Excel 리포트 생성
    """
    if financial_data is None or financial_data.empty:
        raise ValueError("financial_data가 없습니다.")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        financial_data.to_excel(writer, index=False, sheet_name='재무데이터')
        writer.save()
    output.seek(0)
    return output.read()


# === 테스트 및 실행 ===
if __name__ == "__main__":
    # 본인 데이터 프레임 예시 (아래는 테스트용)
    data = {
        "항목": ["매출액", "매출원가", "판관비", "영업이익"],
        "2024Q1": [1000, 600, 200, 200],
        "2024Q2": [1100, 650, 210, 240]
    }
    df_fin = pd.DataFrame(data)

    # PDF 생성 테스트
    pdf_bytes = create_enhanced_pdf_report(financial_data=df_fin, report_author="한명찬")
    with open("test_report.pdf", "wb") as f:
        f.write(pdf_bytes)
    print("test_report.pdf 생성 완료")

    # Excel 생성 테스트
    excel_bytes = create_excel_report(financial_data=df_fin)
    with open("test_report.xlsx", "wb") as f:
        f.write(excel_bytes)
    print("test_report.xlsx 생성 완료")
