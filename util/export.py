# -*- coding: utf-8 -*-
"""
필수 라이브러리 설치:
pip install reportlab plotly xlsxwriter
"""

import io
import os
import re
import tempfile
import pandas as pd
from datetime import datetime

# reportlab 임포트
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# plotly 임포트 및 사용 가능 여부 체크
try:
    import plotly.express as px
    import plotly.graph_objects as go
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
    PDF 보고서 생성 (한글 포함, CMD/Streamlit 모두 호환)
    - financial_data: 재무 데이터 (DataFrame)
    - news_data: 뉴스 데이터 (DataFrame, "제목" 컬럼 필수)
    - insights: AI 인사이트 텍스트 (str)
    - selected_charts: plotly Figure 리스트 (추가 차트 삽입용)
    - quarterly_df: 분기별 데이터 (DataFrame, '분기','회사','영업이익률' 등 컬럼 포함 시 차트 생성)
    - show_footer: 하단 문구 출력 여부
    """

    # 1. 폰트 경로 설정 (본인 PC에서 CMD 성공한 절대 경로로 꼭 변경하세요)
    local_font_paths = {
        "Korean": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothic.ttf",
        "KoreanBold": r"C:\Users\songo\OneDrive\써니C\예시\nanum-gothic\NanumGothicBold.ttf",
        "KoreanSerif": r"C:\Users\songo\OneDrive\써니C\예시\nanum-myeongjo\NanumMyeongjo.ttf",
    }
    fallback_fonts = {
        "Korean": "Helvetica",
        "KoreanBold": "Helvetica-Bold",
        "KoreanSerif": "Times-Roman",
    }

    # 2. 폰트 등록
    for font_name, font_path in local_font_paths.items():
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
            except Exception as e:
                print(f"폰트 등록 실패: {font_name} - {e}")
                local_font_paths[font_name] = fallback_fonts[font_name]
        else:
            print(f"폰트 파일 없음: {font_path}")
            local_font_paths[font_name] = fallback_fonts[font_name]

    # 3. 스타일 정의
    styles = getSampleStyleSheet()
    styles['Normal'].fontName = local_font_paths["Korean"]
    TITLE_STYLE = ParagraphStyle(
        'Title',
        fontName=local_font_paths["KoreanBold"],
        fontSize=20,
        leading=30,
        spaceAfter=15
    )
    HEADING_STYLE = ParagraphStyle(
        'Heading',
        fontName=local_font_paths["KoreanBold"],
        fontSize=14,
        leading=23,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'Body',
        fontName=local_font_paths["KoreanSerif"],
        fontSize=12,
        leading=18,
        spaceAfter=6
    )

    # 4. PDF 문서 생성 준비
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)

    story = []

    # 표지
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 20))

    # 5. 재무분석 결과 표
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), local_font_paths["KoreanBold"]),
            ('FONTNAME', (0, 1), (-1, -1), local_font_paths["KoreanSerif"]),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 6. Plotly 차트 이미지 변환 함수
    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception:
            return None

    # 7. 시각화 차트 삽입 (Plotly 가능할 경우)
    charts_added = False
    if PLOTLY_AVAILABLE:
        # 주요 비율 막대그래프 자동 삽입 (financial_data에 % 지표가 있을 때)
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
                            except:
                                pass
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
                            try:
                                os.unlink(tmp_path)
                            except:
                                pass
                            charts_added = True
                        else:
                            story.append(Paragraph("※ 환경 제약으로 차트 이미지는 제외되었습니다.", BODY_STYLE))
        except Exception as e:
            story.append(Paragraph(f"막대그래프 생성 오류: {e}", BODY_STYLE))

        # 분기별 추이 그래프 (영업이익률, 매출액)
        try:
            if quarterly_df is not None and hasattr(quarterly_df, "empty") and not quarterly_df.empty:
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

        # 사용자가 전달한 추가 Plotly 차트 삽입
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
                        except:
                            pass
                charts_added = True
        except Exception as e:
            story.append(Paragraph(f"추가 차트 삽입 오류: {e}", BODY_STYLE))

    # 8. 최신 뉴스 하이라이트
    if news_data is not None and hasattr(news_data, "empty") and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # 9. AI 인사이트 영역
    if insights:
        story.append(PageBreak())
        story.append(Paragraph("4. AI 인사이트", HEADING_STYLE))
        for line in insights.splitlines():
            story.append(Paragraph(line, BODY_STYLE))

    # 10. 푸터 문구
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # PDF 빌드 및 반환
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def create_excel_report(financial_data: pd.DataFrame) -> bytes:
    """
    간단한 Excel 보고서 생성 (financial_data 필수)
    """
    if financial_data is None or financial_data.empty:
        raise ValueError("financial_data가 없습니다.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        financial_data.to_excel(writer, index=False, sheet_name='재무데이터')
        writer.save()
    output.seek(0)
    return output.read()
