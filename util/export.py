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

# Plotly 라이브러리 임포트 시도
try:
    import plotly
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
    PDF 보고서 생성 함수 (한글 포함, 절대경로 폰트 등록 + 다양한 기능 완전 포함)
    """

    # 1. 폰트 절대경로 (본인 환경에 맞게 반드시 변경하세요)
    korean_font_path = r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothic.ttf"
    korean_bold_path = r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothicBold.ttf"
    korean_serif_path = r"C:\Users\songo\OneDrive\써니C\예시\nanum-myeongjo\NanumMyeongjo.ttf"

    # 2. 폰트 등록
    for font_name, font_path in [
        ("Korean", korean_font_path),
        ("KoreanBold", korean_bold_path),
        ("KoreanSerif", korean_serif_path),
    ]:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                print(f"[폰트 등록 성공] {font_name} from {font_path}")
            except Exception as e:
                print(f"[폰트 등록 실패] {font_name} from {font_path}: {e}")
        else:
            print(f"[폰트 파일 없음] {font_path}")

    styles = getSampleStyleSheet()

    TITLE_STYLE = ParagraphStyle(
        'Title',
        fontName='KoreanBold' if 'KoreanBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold',
        fontSize=20,
        leading=30,
        spaceAfter=15
    )
    HEADING_STYLE = ParagraphStyle(
        'Heading',
        fontName='KoreanBold' if 'KoreanBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold',
        fontSize=14,
        leading=23,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'Body',
        fontName='KoreanSerif' if 'KoreanSerif' in pdfmetrics.getRegisteredFontNames() else 'Times-Roman',
        fontSize=12,
        leading=18,
        spaceAfter=6
    )

    buff = io.BytesIO()

    def _page_number(canvas, doc):
        canvas.setFont('Helvetica', 9)
        page_num_text = f"- {canvas.getPageNumber()} -"
        canvas.drawCentredString(A4[0] / 2, 18, page_num_text)

    doc = SimpleDocTemplate(
        buff,
        pagesize=A4,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54,
    )

    story = []

    # 표지 및 메타 정보
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 20))

    # 1. 재무분석 결과 표
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), 'KoreanBold'),
            ('FONTNAME', (0, 1), (-1, -1), 'KoreanSerif'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # Plotly Figure → PNG bytes 변환 함수 (kaleido 필요)
    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception:
            return None

    # AI 텍스트 마크다운 특수문자 제거 및 텍스트 라인 분류
    def _clean_ai_text(raw: str) -> list[tuple[str, str]]:
        raw = re.sub(r'[*_#>~]', '', raw)  # 마크다운 특수문자 제거
        blocks = []
        for ln in raw.splitlines():
            ln = ln.strip()
            if not ln:
                continue
            # 숫자 제목 패턴 예) 1.2. 제목
            if re.match(r'^\d+(\.\d+)*\s', ln):
                blocks.append(('title', ln))
            else:
                blocks.append(('body', ln))
        return blocks

    # ASCII 표 문자열 블록 → ReportLab Table 변환
    def _ascii_block_to_table(lines: list[str]):
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

    charts_added = False

    # 2. 시각화 차트 삽입 (Plotly → PNG 변환 및 임시파일 삽입)
    if PLOTLY_AVAILABLE:
        try:
            # 주요 비율 비교 막대그래프 생성 예시
            if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty and '구분' in financial_data.columns:
                ratio_rows = financial_data[financial_data['구분'].astype(str).str.contains('%', na=False)].copy()
                if not ratio_rows.empty:
                    key_order = ['영업이익률(%)', '순이익률(%)', '매출총이익률(%)', '매출원가율(%)', '판관비율(%)']
                    ratio_rows['__order__'] = ratio_rows['구분'].apply(lambda x: key_order.index(x) if x in key_order else 999)
                    ratio_rows = ratio_rows.sort_values('__order__').drop(columns='__order__')

                    melt = []
                    company_cols = [c for c in ratio_rows.columns if c != '구분' and not str(c).endswith('_원시값')]
                    for _, r in ratio_rows.iterrows():
                        for comp in company_cols:
                            val = str(r[comp]).replace('%', '').strip()
                            try:
                                melt.append({'지표': r['구분'], '회사': comp, '수치': float(val)})
                            except:
                                pass

                    if melt:
                        import plotly.express as px
                        bar_df = pd.DataFrame(melt)
                        fig_bar = px.bar(bar_df, x='지표', y='수치', color='회사', barmode='group', title="주요 비율 비교")
                        img_bytes = _fig_to_png_bytes(fig_bar, 900, 450)
                        if img_bytes:
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                                tmp.write(img_bytes)
                                tmp_path = tmp.name
                            story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
                            story.append(Paragraph("2-1. 주요 비율 비교 (막대그래프)", BODY_STYLE))
                            story.append(RLImage(tmp_path, width=500, height=280))
                            story.append(Spacer(1, 16))
                            try:
                                os.unlink(tmp_path)
                            except:
                                pass
                            charts_added = True
                        else:
                            story.append(Paragraph("※ 환경 제약으로 차트 이미지는 제외되었습니다.", BODY_STYLE))
        except Exception as e:
            story.append(Paragraph(f"막대그래프 생성 오류: {e}", BODY_STYLE))

        try:
            # 분기별 추이 꺾은선 차트 (영업이익률, 매출액)
            if quarterly_df is not None and hasattr(quarterly_df, "empty") and not quarterly_df.empty:
                import plotly.graph_objects as go

                # 영업이익률 추이
                if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
                    fig_line = go.Figure()
                    for comp in quarterly_df['회사'].dropna().unique():
                        cdf = quarterly_df[quarterly_df['회사'] == comp]
                        fig_line.add_trace(go.Scatter(x=cdf['분기'], y=cdf['영업이익률'], mode='lines+markers', name=f"{comp}"))
                    fig_line.update_layout(title="분기별 영업이익률 추이", xaxis_title="분기", yaxis_title="영업이익률(%)")
                    img_bytes = _fig_to_png_bytes(fig_line, 900, 450)
                    if img_bytes:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                            tmp.write(img_bytes)
                            tmp_path = tmp.name
                        if not charts_added:
                            story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
                        story.append(Paragraph("2-2. 분기별 영업이익률 추이 (꺾은선)", BODY_STYLE))
                        story.append(RLImage(tmp_path, width=500, height=280))
                        story.append(Spacer(1, 16))
                        try:
                            os.unlink(tmp_path)
                        except:
                            pass
                        charts_added = True
                    else:
                        story.append(Paragraph("※ 환경 제약으로 차트 이미지는 제외되었습니다.", BODY_STYLE))

                # 매출액 추이
                if all(col in quarterly_df.columns for col in ['분기', '회사', '매출액']):
                    fig_rev = go.Figure()
                    for comp in quarterly_df['회사'].dropna().unique():
                        cdf = quarterly_df[quarterly_df['회사'] == comp]
                        fig_rev.add_trace(go.Scatter(x=cdf['분기'], y=cdf['매출액'], mode='lines+markers', name=f"{comp}"))
                    fig_rev.update_layout(title="분기별 매출액 추이", xaxis_title="분기", yaxis_title="매출액(조원)")
                    img_bytes = _fig_to_png_bytes(fig_rev, 900, 450)
                    if img_bytes:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                            tmp.write(img_bytes)
                            tmp_path = tmp.name
                        story.append(Paragraph("2-3. 분기별 매출액 추이 (꺾은선)", BODY_STYLE))
                        story.append(RLImage(tmp_path, width=500, height=280))
                        story.append(Spacer(1, 16))
                        try:
                            os.unlink(tmp_path)
                        except:
                            pass
                        charts_added = True
                    else:
                        story.append(Paragraph("※ 환경 제약으로 차트 이미지는 제외되었습니다.", BODY_STYLE))
        except Exception as e:
            story.append(Paragraph(f"추이 그래프 생성 오류: {e}", BODY_STYLE))

        try:
            # 선택 차트 추가 삽입
            if selected_charts:
                if not charts_added:
                    story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
                for idx, fig in enumerate(selected_charts, start=1):
                    img_bytes = _fig_to_png_bytes(fig, 900, 450)
                    if img_bytes:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                            tmp.write(img_bytes)
                            tmp_path = tmp.name
                        story.append(Paragraph(f"2-{idx+3}. 추가 차트", BODY_STYLE))
                        story.append(RLImage(tmp_path, width=500, height=280))
                        story.append(Spacer(1, 16))
                        try:
                            os.unlink(tmp_path)
                        except:
                            pass
                charts_added = True
        except Exception as e:
            story.append(Paragraph(f"추가 차트 삽입 오류: {e}", BODY_STYLE))

    # 3. 최신 뉴스 하이라이트
    if news_data is not None and hasattr(news_data, "empty") and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # 4. AI 인사이트 (마크다운 특수문자 제거, ASCII 표 변환 포함)
    if insights:
        story.append(PageBreak())
        story.append(Paragraph("4. AI 인사이트", HEADING_STYLE))
        blocks = _clean_ai_text(str(insights))
        ascii_buf = []
        for typ, ln in blocks:
            if '|' in ln:
                ascii_buf.append(ln)
                continue
            if ascii_buf:
                tbl = _ascii_block_to_table(ascii_buf)
                if tbl:
                    story.append(tbl)
                story.append(Spacer(1, 12))
                ascii_buf.clear()
            if typ == 'title':
                story.append(Paragraph(f"<b>{ln}</b>", BODY_STYLE))
            else:
                story.append(Paragraph(ln, BODY_STYLE))
        if ascii_buf:
            tbl = _ascii_block_to_table(ascii_buf)
            if tbl:
                story.append(tbl)

    # 5. 옵션: 푸터 표시
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # PDF 빌드
    doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
    buff.seek(0)
    return buff.getvalue()


# 테스트 실행 예시
if __name__ == "__main__":
    pdf_bytes = create_enhanced_pdf_report()
    with open("enhanced_korean_pdf_report.pdf", "wb") as f:
        f.write(pdf_bytes)
    print("PDF 파일이 생성되었습니다: enhanced_korean_pdf_report.pdf")
