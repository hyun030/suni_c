# -*- coding: utf-8 -*-
import io
import os
import tempfile
import pandas as pd
from datetime import datetime
import streamlit as st
import re

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
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False


def get_font_paths():
    """스트림릿 환경에 맞는 폰트 경로를 반환"""
    font_paths = {
        "Korean": "fonts/NanumGothic.ttf",
        "KoreanBold": "fonts/NanumGothicBold.ttf", 
        "KoreanSerif": "fonts/NanumMyeongjo.ttf"
    }
    
    found_fonts = {}
    for font_name, font_path in font_paths.items():
        if os.path.exists(font_path):
            file_size = os.path.getsize(font_path)
            if file_size > 0:
                found_fonts[font_name] = font_path
    
    return found_fonts


def register_fonts_safe():
    """안전하게 폰트를 등록하고 사용 가능한 폰트 이름을 반환"""
    font_paths = get_font_paths()
    registered_fonts = {}
    
    # 기본 폰트 설정
    default_fonts = {
        "Korean": "Helvetica",
        "KoreanBold": "Helvetica-Bold", 
        "KoreanSerif": "Times-Roman"
    }
    
    for font_name, default_font in default_fonts.items():
        if font_name in font_paths:
            try:
                if font_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(font_name, font_paths[font_name]))
                registered_fonts[font_name] = font_name
            except Exception:
                if font_name == "KoreanSerif":
                    try:
                        if "Korean" in font_paths and "Korean" in pdfmetrics.getRegisteredFontNames():
                            registered_fonts[font_name] = "Korean"
                        else:
                            registered_fonts[font_name] = default_font
                    except:
                        registered_fonts[font_name] = default_font
                else:
                    registered_fonts[font_name] = default_font
        else:
            registered_fonts[font_name] = default_font
    
    return registered_fonts


def create_excel_report(financial_data=None, news_data=None, insights=None):
    """Excel 보고서 생성"""
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


def clean_ai_text(raw: str):
    """AI 인사이트 텍스트 정리"""
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


def ascii_to_table(lines, registered_fonts, header_color='#E31E24', row_colors=None):
    """ASCII 표를 reportlab 테이블로 변환"""
    if not lines:
        return None
    
    header = [c.strip() for c in lines[0].split('|') if c.strip()]
    if not header:
        return None
        
    data = []
    for ln in lines[2:]:  # 구분선 건너뛰기
        cols = [c.strip() for c in ln.split('|') if c.strip()]
        if len(cols) == len(header):
            data.append(cols)
    
    if not data:
        return None
    
    if row_colors is None:
        row_colors = [colors.whitesmoke, colors.HexColor('#F7F7F7')]
    
    tbl = Table([header] + data)
    tbl.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(header_color)),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
        ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), row_colors),
    ]))
    return tbl


def fig_to_png_bytes(fig, width=900, height=450):
    """Plotly 차트를 PNG 바이트로 변환 (조용히)"""
    try:
        return fig.to_image(format="png", width=width, height=height)
    except:
        try:
            return fig.to_image(format="png", width=width, height=height, engine="kaleido")
        except:
            try:
                return fig.to_image(format="png", width=width, height=height, engine="orca")
            except:
                return None


def split_dataframe_for_pdf(df, max_rows_per_page=20, max_cols_per_page=8):
    """DataFrame을 PDF에 맞게 페이지별로 분할"""
    chunks = []
    total_rows = len(df)
    total_cols = len(df.columns)
    
    # 행 기준으로 먼저 분할
    for row_start in range(0, total_rows, max_rows_per_page):
        row_end = min(row_start + max_rows_per_page, total_rows)
        row_chunk = df.iloc[row_start:row_end]
        
        # 열 기준으로 분할
        for col_start in range(0, total_cols, max_cols_per_page):
            col_end = min(col_start + max_cols_per_page, total_cols)
            col_names = df.columns[col_start:col_end]
            chunk = row_chunk[col_names]
            
            chunk_info = {
                'data': chunk,
                'row_range': (row_start, row_end-1),
                'col_range': (col_start, col_end-1),
                'is_last_row_chunk': row_end == total_rows,
                'is_last_col_chunk': col_end == total_cols
            }
            chunks.append(chunk_info)
    
    return chunks


def add_chunked_table(story, df, title, registered_fonts, BODY_STYLE, header_color='#F2F2F2'):
    """분할된 테이블을 story에 추가"""
    if df is None or df.empty:
        story.append(Paragraph(f"{title}: 데이터가 없습니다.", BODY_STYLE))
        return
    
    story.append(Paragraph(title, BODY_STYLE))
    story.append(Spacer(1, 8))
    
    chunks = split_dataframe_for_pdf(df)
    
    for i, chunk_info in enumerate(chunks):
        chunk = chunk_info['data']
        
        if len(chunks) > 1:
            row_info = f"행 {chunk_info['row_range'][0]+1}~{chunk_info['row_range'][1]+1}"
            col_info = f"열 {chunk_info['col_range'][0]+1}~{chunk_info['col_range'][1]+1}"
            story.append(Paragraph(f"[{row_info}, {col_info}]", BODY_STYLE))
        
        # 테이블 데이터 준비
        table_data = [chunk.columns.tolist()]
        for _, row in chunk.iterrows():
            table_data.append([str(val) for val in row.values])
        
        # 테이블 생성
        tbl = Table(table_data, repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(header_color)),
            ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
            ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8F8F8')]),
        ]))
        
        story.append(tbl)
        story.append(Spacer(1, 12))
        
        # 다음 청크가 있고 새 페이지가 필요한 경우
        if i < len(chunks) - 1 and (i + 1) % 2 == 0:  # 2개마다 페이지 나누기
            story.append(PageBreak())


def add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """재무분석 결과 섹션 추가"""
    story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
    
    # 1-1. SK에너지 대비 경쟁사 갭차이 분석표
    if financial_data is not None and not financial_data.empty:
        # 원시값 컬럼 제외
        display_cols = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
        df_display = financial_data[display_cols].copy()
        add_chunked_table(story, df_display, "1-1. SK에너지 대비 경쟁사 갭차이 분석", 
                         registered_fonts, BODY_STYLE, '#F2F2F2')
    else:
        story.append(Paragraph("1-1. SK에너지 대비 경쟁사 갭차이 분석: 데이터가 없습니다.", BODY_STYLE))
    
    story.append(Spacer(1, 12))
    
    # 1-2. 분기별 재무지표 상세 데이터
    if quarterly_df is not None and not quarterly_df.empty:
        add_chunked_table(story, quarterly_df, "1-2. 분기별 재무지표 상세 데이터", 
                         registered_fonts, BODY_STYLE, '#E6F3FF')
    else:
        story.append(Paragraph("1-2. 분기별 재무지표 상세 데이터: 데이터가 없습니다.", BODY_STYLE))
    
    story.append(Spacer(1, 18))


def add_charts_section(story, financial_data, quarterly_df, selected_charts, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """시각화 차트 섹션 추가"""
    story.append(Paragraph("2. 시각화 차트 및 분석", HEADING_STYLE))
    
    charts_added = False
    
    if not PLOTLY_AVAILABLE:
        story.append(Paragraph("차트 생성 라이브러리를 사용할 수 없습니다.", BODY_STYLE))
        return charts_added
    
    # 주요 비율 비교 막대 그래프
    try:
        if (financial_data is not None and not financial_data.empty and 
            '구분' in financial_data.columns):
            
            ratio_rows = financial_data[financial_data['구분'].astype(str).str.contains('%', na=False)].copy()
            
            if not ratio_rows.empty:
                # 주요 지표 순서 정렬
                key_order = ['영업이익률(%)', '순이익률(%)', '매출총이익률(%)', '매출원가율(%)', '판관비율(%)']
                ratio_rows['__order__'] = ratio_rows['구분'].apply(lambda x: key_order.index(x) if x in key_order else 999)
                ratio_rows = ratio_rows.sort_values('__order__').drop(columns='__order__')

                # 데이터 변환
                melt = []
                company_cols = [c for c in ratio_rows.columns if c != '구분' and not str(c).endswith('_원시값')]
                
                for _, r in ratio_rows.iterrows():
                    for comp in company_cols:
                        val = str(r[comp]).replace('%','').strip()
                        try:
                            val_float = float(val)
                            melt.append({'지표': r['구분'], '회사': comp, '수치': val_float})
                        except:
                            continue
                
                if melt:
                    bar_df = pd.DataFrame(melt)
                    fig_bar = px.bar(bar_df, x='지표', y='수치', color='회사', barmode='group', 
                                   title="주요 비율 비교")
                    
                    img_bytes = fig_to_png_bytes(fig_bar)
                    if img_bytes:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                            tmp.write(img_bytes)
                            tmp_path = tmp.name
                        story.append(Paragraph("2-1. 주요 비율 비교 (막대그래프)", BODY_STYLE))
                        story.append(RLImage(tmp_path, width=500, height=280))
                        story.append(Spacer(1, 16))
                        try:
                            os.unlink(tmp_path)
                        except:
                            pass
                        charts_added = True
    except:
        pass

    # 분기별 추이 그래프들
    try:
        if quarterly_df is not None and not quarterly_df.empty:
            # 영업이익률 추이
            if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
                fig_line = go.Figure()
                companies = quarterly_df['회사'].dropna().unique()
                
                for comp in companies:
                    cdf = quarterly_df[quarterly_df['회사'] == comp].copy()
                    fig_line.add_trace(go.Scatter(
                        x=cdf['분기'], 
                        y=cdf['영업이익률'], 
                        mode='lines+markers', 
                        name=comp,
                        line=dict(width=3),
                        marker=dict(size=8)
                    ))
                    
                fig_line.update_layout(
                    title="분기별 영업이익률 추이", 
                    xaxis_title="분기", 
                    yaxis_title="영업이익률(%)",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                
                img_bytes = fig_to_png_bytes(fig_line)
                if img_bytes:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    story.append(Paragraph("2-2. 분기별 영업이익률 추이", BODY_STYLE))
                    story.append(RLImage(tmp_path, width=500, height=280))
                    story.append(Spacer(1, 16))
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                    charts_added = True

            # 매출액 추이
            if all(col in quarterly_df.columns for col in ['분기', '회사', '매출액']):
                fig_rev = go.Figure()
                
                for comp in quarterly_df['회사'].dropna().unique():
                    cdf = quarterly_df[quarterly_df['회사'] == comp].copy()
                    fig_rev.add_trace(go.Scatter(
                        x=cdf['분기'], 
                        y=cdf['매출액'], 
                        mode='lines+markers', 
                        name=comp,
                        line=dict(width=3),
                        marker=dict(size=8)
                    ))
                    
                fig_rev.update_layout(
                    title="분기별 매출액 추이", 
                    xaxis_title="분기", 
                    yaxis_title="매출액(조원)",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                
                img_bytes = fig_to_png_bytes(fig_rev)
                if img_bytes:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    story.append(Paragraph("2-3. 분기별 매출액 추이", BODY_STYLE))
                    story.append(RLImage(tmp_path, width=500, height=280))
                    story.append(Spacer(1, 16))
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                    charts_added = True
    except:
        pass

    # 외부에서 전달된 차트들
    try:
        if selected_charts:
            chart_counter = 4 if charts_added else 1
            for idx, fig in enumerate(selected_charts, start=chart_counter):
                img_bytes = fig_to_png_bytes(fig)
                if img_bytes:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    story.append(Paragraph(f"2-{idx}. 추가 차트", BODY_STYLE))
                    story.append(RLImage(tmp_path, width=500, height=280))
                    story.append(Spacer(1, 16))
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
            charts_added = True
    except:
        pass
    
    if not charts_added:
        story.append(Paragraph("생성 가능한 차트가 없습니다.", BODY_STYLE))
        story.append(Spacer(1, 18))
    
    return charts_added


def add_ai_insights_section(story, insights, registered_fonts, BODY_STYLE, header_color='#E31E24'):
    """AI 인사이트 섹션 추가"""
    if not insights:
        story.append(Paragraph("2-AI. 분석 인사이트", BODY_STYLE))
        story.append(Paragraph("AI 인사이트가 제공되지 않았습니다.", BODY_STYLE))
        story.append(Spacer(1, 18))
        return
    
    story.append(Paragraph("2-AI. 분석 인사이트", BODY_STYLE))
    story.append(Spacer(1, 8))

    # AI 인사이트 텍스트 처리
    blocks = clean_ai_text(str(insights))
    ascii_buffer = []
    
    for typ, line in blocks:
        if '|' in line:
            ascii_buffer.append(line)
            continue
        
        if ascii_buffer:
            tbl = ascii_to_table(ascii_buffer, registered_fonts, header_color)
            if tbl:
                story.append(tbl)
            story.append(Spacer(1, 12))
            ascii_buffer.clear()
        
        if typ == 'title':
            story.append(Paragraph(f"<b>{line}</b>", BODY_STYLE))
        else:
            story.append(Paragraph(line, BODY_STYLE))
    
    if ascii_buffer:
        tbl = ascii_to_table(ascii_buffer, registered_fonts, header_color)
        if tbl:
            story.append(tbl)
    
    story.append(Spacer(1, 18))


def add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """뉴스 하이라이트 및 종합 분석 섹션 추가"""
    story.append(Paragraph("3. 뉴스 하이라이트 및 종합 분석", HEADING_STYLE))
    
    if news_data is not None and not news_data.empty:
        story.append(Paragraph("3-1. 최신 뉴스 하이라이트", BODY_STYLE))
        for i, title in enumerate(news_data["제목"].head(10), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 16))
        
        if insights:
            story.append(Paragraph("3-2. AI 종합 분석 및 시사점", BODY_STYLE))
            story.append(Spacer(1, 8))
            
            blocks = clean_ai_text(str(insights))
            ascii_buffer = []
            
            for typ, line in blocks:
                if '|' in line:
                    ascii_buffer.append(line)
                    continue
                
                if ascii_buffer:
                    tbl = ascii_to_table(ascii_buffer, registered_fonts, '#0066CC', 
                                       [colors.whitesmoke, colors.HexColor('#F0F8FF')])
                    if tbl:
                        story.append(tbl)
                    story.append(Spacer(1, 12))
                    ascii_buffer.clear()
                
                if typ == 'title':
                    story.append(Paragraph(f"<b>{line}</b>", BODY_STYLE))
                else:
                    story.append(Paragraph(line, BODY_STYLE))
            
            if ascii_buffer:
                tbl = ascii_to_table(ascii_buffer, registered_fonts, '#0066CC',
                                   [colors.whitesmoke, colors.HexColor('#F0F8FF')])
                if tbl:
                    story.append(tbl)
        else:
            story.append(Paragraph("AI 종합 분석이 제공되지 않았습니다.", BODY_STYLE))
    else:
        story.append(Paragraph("뉴스 데이터가 제공되지 않았습니다.", BODY_STYLE))
        
        if insights:
            story.append(Paragraph("3-1. 종합 분석 및 시사점", BODY_STYLE))
            story.append(Spacer(1, 8))
            
            blocks = clean_ai_text(str(insights))
            ascii_buffer = []
            
            for typ, line in blocks:
                if '|' in line:
                    ascii_buffer.append(line)
                    continue
                
                if ascii_buffer:
                    tbl = ascii_to_table(ascii_buffer, registered_fonts, '#228B22',
                                       [colors.whitesmoke, colors.HexColor('#F0FFF0')])
                    if tbl:
                        story.append(tbl)
                    story.append(Spacer(1, 12))
                    ascii_buffer.clear()
                
                if typ == 'title':
                    story.append(Paragraph(f"<b>{line}</b>", BODY_STYLE))
                else:
                    story.append(Paragraph(line, BODY_STYLE))
            
            if ascii_buffer:
                tbl = ascii_to_table(ascii_buffer, registered_fonts, '#228B22',
                                   [colors.whitesmoke, colors.HexColor('#F0FFF0')])
                if tbl:
                    story.append(tbl)
        else:
            story.append(Paragraph("AI 인사이트도 제공되지 않았습니다.", BODY_STYLE))
    
    story.append(Spacer(1, 18))


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
    """향상된 PDF 보고서 생성"""
    
    # 조용히 폰트 등록
    registered_fonts = register_fonts_safe()
    
    # 스타일 정의
    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'Title',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=20,
        leading=30,
        spaceAfter=15,
        alignment=1,  # 중앙 정렬
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
    add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE)
    
    # 2. 시각화 차트 및 분석
    charts_added = add_charts_section(story, financial_data, quarterly_df, selected_charts, 
                                    registered_fonts, HEADING_STYLE, BODY_STYLE)
    
    # 2-AI. AI 인사이트
    add_ai_insights_section(story, insights, registered_fonts, BODY_STYLE)
    
    # 3. 뉴스 하이라이트 및 종합 분석
    add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE)

    # 푸터 (선택사항)
    if show_footer:
        story.append(Spacer(1, 24))
        footer_text = "※ 본 보고서는 대시보드에서 자동 생성되었습니다."
        story.append(Paragraph(footer_text, BODY_STYLE))

    # 페이지 번호 추가 함수
    def _page_number(canvas, doc):
        canvas.setFont('Helvetica', 9)
        canvas.drawCentredString(A4[0]/2, 20, f"- {canvas.getPageNumber()} -")

    # PDF 문서 생성
    doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
    buffer.seek(0)
    return buffer.getvalue()
