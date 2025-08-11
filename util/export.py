# -*- coding: utf-8 -*-
import io
import os
import sys
import logging
import pandas as pd
from datetime import datetime

# --- 로깅 설정 ---
logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')

# PDF 라이브러리
try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib import colors
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak, Image as RLImage
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    PDF_AVAILABLE = True
except ImportError:
    logging.warning("ReportLab 라이브러리가 설치되지 않아 PDF 생성이 불가합니다.")
    PDF_AVAILABLE = False

# Plotly 라이브러리
try:
    import plotly
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    logging.warning("Plotly 라이브러리가 설치되지 않아 차트 생성이 불가합니다.")
    PLOTLY_AVAILABLE = False

from typing import Optional, List, Union

def create_excel_report(
    financial_data: Optional[pd.DataFrame] = None,
    news_data: Optional[pd.DataFrame] = None,
    insights: Optional[str] = None
) -> bytes:
    """데이터프레임과 텍스트를 Excel 파일로 변환"""
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
    financial_data: Optional[pd.DataFrame] = None,
    news_data: Optional[pd.DataFrame] = None,
    insights: Optional[str] = None,
    selected_charts: Optional[List[plotly.graph_objects.Figure]] = None,
    quarterly_df: Optional[pd.DataFrame] = None,
    show_footer: bool = False,
    report_target: str = "SK이노베이션 경영진",
    report_author: str = "보고자 미기재"
) -> Optional[bytes]:
    """
    • 제목  : KoreanBold 20pt
    • 본문  : KoreanSerif 12pt, 줄간격 170%
    • 표    : ReportLab Table
    • 차트  : Plotly → PNG (kaleido 미설치 시 자동 생략)
    • 섹션  : 1 재무분석 → 2 시각화 → 3 뉴스 → 4 AI 인사이트 (요청 순서)
    """
    if not PDF_AVAILABLE:
        logging.error("PDF 생성에 필요한 ReportLab 라이브러리가 없습니다.")
        return None

    import re
    import tempfile

    def _fig_to_png_bytes(fig, width=900, height=450) -> Optional[bytes]:
        """Plotly 차트를 PNG 바이트로 변환. Kaleido 없으면 None 반환."""
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception as e:
            logging.warning(f"Plotly 차트 PNG 변환 실패: {e}")
            return None

    def _clean_ai_text(raw: str) -> List[tuple[str, str]]:
        """마크다운 기호 제거 후 ('title'|'body', line) 형식으로 반환"""
        raw = re.sub(r'[*_`#>~]', '', raw)  # 굵게/이탤릭/코드/인용/물결 제거
        blocks = []
        for ln in raw.splitlines():
            ln = ln.strip()
            if not ln:
                continue
            if re.match(r'^\d+(\.\d+)*\s', ln):    # 1. / 1.1 처럼 시작하면 제목 라인
                blocks.append(('title', ln))
            else:
                blocks.append(('body', ln))
        return blocks

    def _ascii_block_to_table(lines: List[str]) -> Optional[Table]:
        """파이프(|) 표 → ReportLab Table"""
        header = [c.strip() for c in lines[0].split('|') if c.strip()]
        data = []
        for ln in lines[2:]:  # 구분선(----) 라인 건너뜀
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
            ('FONTNAME', (0, 1), (-1, -1), 'KoreanSerif'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
             [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
        ]))
        return tbl

    # 폰트 등록
    base_dir = ""
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base_dir = os.getcwd()
        logging.info(f"__file__ 변수가 없으므로 현재 작업 디렉토리({base_dir})를 기준으로 폰트 경로 설정")

    font_paths = {
        "Korean":      [os.path.join(base_dir, "fonts", "NanumGothic.ttf")],
        "KoreanBold":  [os.path.join(base_dir, "fonts", "NanumGothicBold.ttf")],
        "KoreanSerif": [os.path.join(base_dir, "fonts", "NanumMyeongjoBold.ttf")]  # 실제 파일명 확인 필요
    }
    for family, paths in font_paths.items():
        font_registered = False
        for p in paths:
            if os.path.exists(p):
                try:
                    pdfmetrics.registerFont(TTFont(family, p))
                    logging.info(f"폰트 등록 성공: {family} ({p})")
                    font_registered = True
                    break
                except Exception as e:
                    logging.warning(f"폰트 등록 실패: {family} ({p}) → {e}")
            else:
                logging.warning(f"폰트 파일 없음: {p}")
        if not font_registered:
            logging.warning(f"{family} 글꼴 등록에 실패했습니다. 기본 폰트로 대체됩니다.")

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

    # 표지 및 메타 정보
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE
    ))
    story.append(Spacer(1, 12))

    # 1. 재무분석 결과
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        # 원시값 컬럼 제거 (존재 시)
        display_cols = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
        if display_cols:
            df_disp = financial_data[display_cols].copy()
            if not df_disp.empty:
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
            else:
                logging.info("재무분석 데이터가 비어있어 표 생성 생략")
        else:
            logging.info("재무분석 데이터에 출력 가능한 컬럼이 없습니다.")

    # 2. 시각화 차트
    charts_added = False
    if PLOTLY_AVAILABLE:
        # 2-1) 주요 비율(%) 비교 막대그래프
        try:
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
                            except Exception:
                                continue

                    if melt:
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
                            charts_added = True
                            try:
                                os.unlink(tmp_path)
                            except Exception as e:
                                logging.warning(f"임시파일 삭제 실패: {e}")
                        else:
                            story.append(Paragraph("※ 환경 제약으로 차트 이미지는 제외되었습니다.", BODY_STYLE))
        except Exception as e:
            logging.error(f"막대그래프 생성 오류: {e}")
            story.append(Paragraph(f"막대그래프 생성 오류: {e}", BODY_STYLE))

        # 2-2, 2-3) 분기별 꺾은선 그래프 (영업이익률, 매출액)
        try:
            if quarterly_df is not None and hasattr(quarterly_df, "empty") and not quarterly_df.empty:
                # 영업이익률 꺾은선
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
                        charts_added = True
                        try:
                            os.unlink(tmp_path)
                        except Exception as e:
                            logging.warning(f"임시파일 삭제 실패: {e}")

                # 매출액 꺾은선
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
                        charts_added = True
                        try:
                            os.unlink(tmp_path)
                        except Exception as e:
                            logging.warning(f"임시파일 삭제 실패: {e}")
                    else:
                        story.append(Paragraph("※ 환경 제약으로 차트 이미지는 제외되었습니다.", BODY_STYLE))
        except Exception as e:
            logging.error(f"추이 그래프 생성 오류: {e}")
            story.append(Paragraph(f"추이 그래프 생성 오류: {e}", BODY_STYLE))

        # 2-4) 외부에서 전달된 Plotly 그림들(selected_charts)
        try:
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
                        except Exception as e:
                            logging.warning(f"임시파일 삭제 실패: {e}")
                charts_added = True
        except Exception as e:
            logging.error(f"추가 차트 삽입 오류: {e}")
            story.append(Paragraph(f"추가 차트 삽입 오류: {e}", BODY_STYLE))

    # 3. 최신 뉴스 하이라이트
    if news_data is not None and hasattr(news_data, "empty") and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # 4. AI 인사이트
    if insights:
        story.append(PageBreak())  # AI 인사이트는 새 페이지에서 시작
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

    # 푸터 (선택)
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # PDF 빌드
    try:
        doc.build(story, onFirstPage=_page_no, onLaterPages=_page_no)
        buff.seek(0)
        return buff.getvalue()
    except Exception as e:
        logging.error(f"PDF 빌드 중 오류 발생: {e}")
        return None
