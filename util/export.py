# -*- coding: utf-8 -*-
import io
import os
import tempfile
import pandas as pd
from datetime import datetime

# reportlab import
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Plotly import 및 사용 가능 여부 체크
try:
    import plotly
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False


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


def create_enhanced_pdf_report(
    financial_data=None,
    news_data=None,
    insights: str | None = None,
    selected_charts: list | None = None,
    quarterly_df: pd.DataFrame | None = None,
    show_footer: bool = False,
    report_target: str = "SK이노베이션 경영진",
    report_author: str = "보고자 미기재",
    font_paths: dict | None = None,
):
    if font_paths is None:
        # 폰트 경로 반드시 본인 환경에 맞게 수정 필요
        font_paths = {
            "Korean": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothic.ttf",
            "KoreanBold": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothicBold.ttf",
            "KoreanSerif": r"C:\Users\songo\OneDrive\써니C\예시\nanum-myeongjo\NanumMyeongjo.ttf",
        }

    # 폰트 등록
    for font_name, path in font_paths.items():
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, path))
            except Exception as e:
                print(f"폰트 등록 실패 {font_name}: {e}")
        else:
            print(f"폰트 파일 없음: {path}")

    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'Title',
        fontName='KoreanBold' if 'KoreanBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold',
        fontSize=20,
        leading=30,
        spaceAfter=15,
    )
    HEADING_STYLE = ParagraphStyle(
        'Heading',
        fontName='KoreanBold' if 'KoreanBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold',
        fontSize=14,
        leading=23,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10,
    )
    BODY_STYLE = ParagraphStyle(
        'Body',
        fontName='KoreanSerif' if 'KoreanSerif' in pdfmetrics.getRegisteredFontNames() else 'Times-Roman',
        fontSize=12,
        leading=18,
        spaceAfter=6,
    )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)

    story = []
    # 표지
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 20))

    # 재무분석 테이블
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        cols_to_show = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
        df_disp = financial_data[cols_to_show].copy()
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

    # Plotly 차트 PNG 변환 헬퍼
    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception:
            return None

    charts_added = False
    if PLOTLY_AVAILABLE:
        # 주요 비율 비교 막대 그래프 (예시)
        try:
            if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty and '구분' in financial_data.columns:
                ratio_rows = financial_data[financial_data['구분'].astype(str).str.contains('%', na=False)].copy()
                if not ratio_rows.empty:
                    import plotly.express as px
                    key_order = ['영업이익률(%)', '순이익률(%)', '매출총이익률(%)', '매출원가율(%)', '판관비율(%)']
                    ratio_rows['__order__'] = ratio_rows['구분'].apply(lambda x: key_order.index(x) if x in key_order else 999)
                    ratio_rows = ratio_rows.sort_values('__order__').drop(columns='__order__')

                    melt = []
                    company_cols = [c for c in ratio_rows.columns if c != '구분' and not str(c).endswith('_원시값')]
                    for _, r in ratio_rows.iterrows():
                        for comp in company_cols:
                            val = str(r[comp]).replace('%','').strip()
                            try:
                                melt.append({'지표': r['구분'], '회사': comp, '수치': float(val)})
                            except:
                                pass
                    if melt:
                        bar_df = pd.DataFrame(melt)
                        fig_bar = px.bar(bar_df, x='지표', y='수치', color='회사', barmode='group', title="주요 비율 비교")
                        img_bytes = _fig_to_png_bytes(fig_bar)
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
        except Exception as e:
            story.append(Paragraph(f"막대그래프 생성 오류: {e}", BODY_STYLE))

        # 분기별 꺾은선 그래프 (영업이익률, 매출액)
        try:
            if quarterly_df is not None and hasattr(quarterly_df, "empty") and not quarterly_df.empty:
                import plotly.graph_objects as go

                # 영업이익률
                if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
                    fig_line = go.Figure()
                    for comp in quarterly_df['회사'].dropna().unique():
                        cdf = quarterly_df[quarterly_df['회사'] == comp]
                        fig_line.add_trace(go.Scatter(x=cdf['분기'], y=cdf['영업이익률'], mode='lines+markers', name=comp))
                    fig_line.update_layout(title="분기별 영업이익률 추이", xaxis_title="분기", yaxis_title="영업이익률(%)")
                    img_bytes = _fig_to_png_bytes(fig_line)
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

                # 매출액
                if all(col in quarterly_df.columns for col in ['분기', '회사', '매출액']):
                    fig_rev = go.Figure()
                    for comp in quarterly_df['회사'].dropna().unique():
                        cdf = quarterly_df[quarterly_df['회사'] == comp]
                        fig_rev.add_trace(go.Scatter(x=cdf['분기'], y=cdf['매출액'], mode='lines+markers', name=comp))
                    fig_rev.update_layout(title="분기별 매출액 추이", xaxis_title="분기", yaxis_title="매출액(조원)")
                    img_bytes = _fig_to_png_bytes(fig_rev)
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
        except Exception as e:
            story.append(Paragraph(f"추이 그래프 생성 오류: {e}", BODY_STYLE))

        # 외부 전달된 Plotly 차트들(selected_charts)
        try:
            if selected_charts:
                if not charts_added:
                    story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
                for idx, fig in enumerate(selected_charts, start=1):
                    img_bytes = _fig_to_png_bytes(fig)
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

    # 최신 뉴스
    if news_data is not None and hasattr(news_data, "empty") and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # AI 인사이트 텍스트 + 표 처리
    if insights:
        story.append(PageBreak())
        story.append(Paragraph("4. AI 인사이트", HEADING_STYLE))

        import re
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

        def ascii_to_table(lines):
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

        blocks = clean_ai_text(str(insights))
        ascii_buffer = []
        for typ, line in blocks:
            if '|' in line:
                ascii_buffer.append(line)
                continue
            if ascii_buffer:
                tbl = ascii_to_table(ascii_buffer)
                if tbl:
                    story.append(tbl)
                story.append(Spacer(1, 12))
                ascii_buffer.clear()
            if typ == 'title':
                story.append(Paragraph(f"<b>{line}</b>", BODY_STYLE))
            else:
                story.append(Paragraph(line, BODY_STYLE))
        if ascii_buffer:
            tbl = ascii_to_table(ascii_buffer)
            if tbl:
                story.append(tbl)

    # 푸터 (선택)
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # 페이지 번호 넣기
    def _page_number(canvas, doc):
        canvas.setFont('Helvetica', 9)
        canvas.drawCentredString(A4[0]/2, 20, f"- {canvas.getPageNumber()} -")

    doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
    buffer.seek(0)
    return buffer.getvalue()
