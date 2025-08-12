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
    
    # 파일 존재 여부 및 유효성 확인 후 반환
    found_fonts = {}
    for font_name, font_path in font_paths.items():
        if os.path.exists(font_path):
            file_size = os.path.getsize(font_path)
            if file_size > 0:
                found_fonts[font_name] = font_path
            else:
                st.warning(f"⚠️ 폰트 파일이 비어있음: {font_path} (크기: {file_size})")
        else:
            st.warning(f"⚠️ 폰트 파일을 찾을 수 없음: {font_path}")
    
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
                # 이미 등록된 폰트인지 확인
                if font_name not in pdfmetrics.getRegisteredFontNames():
                    pdfmetrics.registerFont(TTFont(font_name, font_paths[font_name]))
                    st.success(f"✅ {font_name} 폰트 등록 성공: {font_paths[font_name]}")
                else:
                    st.info(f"ℹ️ {font_name} 폰트 이미 등록됨")
                registered_fonts[font_name] = font_name
            except Exception as e:
                st.error(f"❌ {font_name} 폰트 등록 실패: {e}")
                st.info(f"🔄 {font_name} 대신 기본 폰트 사용: {default_font}")
                
                # KoreanSerif가 실패하면 대안으로 NanumGothic 사용 시도
                if font_name == "KoreanSerif":
                    try:
                        if "Korean" in font_paths and "Korean" in pdfmetrics.getRegisteredFontNames():
                            registered_fonts[font_name] = "Korean"
                            st.info(f"✨ KoreanSerif 대신 NanumGothic 사용")
                        else:
                            registered_fonts[font_name] = default_font
                    except:
                        registered_fonts[font_name] = default_font
                else:
                    registered_fonts[font_name] = default_font
        else:
            st.warning(f"⚠️ {font_name} 폰트 파일을 찾을 수 없음. 기본 폰트 사용: {default_font}")
            registered_fonts[font_name] = default_font
    
    st.write("📝 최종 사용될 폰트들:", registered_fonts)
    return registered_fonts


def debug_font_info():
    """폰트 정보를 디버깅하기 위한 함수"""
    st.write("🔍 **폰트 디버깅 정보**")
    st.write(f"현재 작업 디렉토리: {os.getcwd()}")
    
    font_files = ["fonts/NanumGothic.ttf", "fonts/NanumGothicBold.ttf", "fonts/NanumMyeongjo.ttf"]
    for font_file in font_files:
        if os.path.exists(font_file):
            size = os.path.getsize(font_file)
            st.write(f"✅ {font_file} 존재 (크기: {size:,} bytes)")
        else:
            st.write(f"❌ {font_file} 없음")
    
    st.write(f"reportlab 버전: {__import__('reportlab').__version__}")
    st.write("---")


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
    raw = re.sub(r'[*_#>~]', '', raw)  # 마크다운 문자 제거
    blocks = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        # 숫자로 시작하는 제목 판별
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
    """Plotly 차트를 PNG 바이트로 변환"""
    if not PLOTLY_AVAILABLE:
        return None
    try:
        return fig.to_image(format="png", width=width, height=height)
    except Exception as e:
        st.warning(f"차트 이미지 변환 실패: {e}")
        return None


def create_financial_charts(financial_data, quarterly_df):
    """재무 데이터로부터 차트를 생성하는 함수"""
    charts = []
    
    if not PLOTLY_AVAILABLE:
        return charts
    
    # 1. 주요 비율 비교 막대 그래프
    try:
        if financial_data is not None and not financial_data.empty and '구분' in financial_data.columns:
            # 비율 데이터 추출 (% 포함된 행들)
            ratio_rows = financial_data[financial_data['구분'].astype(str).str.contains('%', na=False)].copy()
            
            if not ratio_rows.empty:
                # 주요 지표 순서 정렬
                key_order = ['매출총이익률(%)', '영업이익률(%)', '순이익률(%)', '매출원가율(%)', '판관비율(%)']
                
                # 데이터 변환
                melt_data = []
                company_cols = [c for c in ratio_rows.columns if c not in ['구분'] and not str(c).endswith('_원시값')]
                
                for _, row in ratio_rows.iterrows():
                    metric_name = row['구분']
                    for company in company_cols:
                        value_str = str(row[company]).replace('%', '').strip()
                        try:
                            value = float(value_str)
                            melt_data.append({
                                '지표': metric_name,
                                '회사': company,
                                '수치': value
                            })
                        except (ValueError, TypeError):
                            continue
                
                if melt_data:
                    bar_df = pd.DataFrame(melt_data)
                    # 주요 지표만 필터링
                    bar_df = bar_df[bar_df['지표'].isin(key_order)]
                    
                    fig_bar = px.bar(
                        bar_df, 
                        x='지표', 
                        y='수치', 
                        color='회사', 
                        barmode='group',
                        title="주요 수익성 지표 비교 (%)",
                        labels={'수치': '비율 (%)'}
                    )
                    fig_bar.update_layout(
                        xaxis_tickangle=-45,
                        height=400,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    charts.append(("주요 수익성 지표 비교", fig_bar))
    except Exception as e:
        st.error(f"비율 비교 차트 생성 오류: {e}")
    
    # 2. 절대값 비교 (매출액, 영업이익 등)
    try:
        if financial_data is not None and not financial_data.empty:
            # 절대값 지표들 (조원 단위)
            absolute_metrics = ['매출액(조원)', '영업이익(조원)', '순이익(조원)', '총자산(조원)']
            absolute_rows = financial_data[financial_data['구분'].isin(absolute_metrics)].copy()
            
            if not absolute_rows.empty:
                melt_abs = []
                company_cols = [c for c in absolute_rows.columns if c not in ['구분'] and not str(c).endswith('_원시값')]
                
                for _, row in absolute_rows.iterrows():
                    metric_name = row['구분']
                    for company in company_cols:
                        try:
                            value = float(str(row[company]).replace('조원', '').replace(',', '').strip())
                            melt_abs.append({
                                '지표': metric_name,
                                '회사': company,
                                '금액': value
                            })
                        except (ValueError, TypeError):
                            continue
                
                if melt_abs:
                    abs_df = pd.DataFrame(melt_abs)
                    fig_abs = px.bar(
                        abs_df,
                        x='지표',
                        y='금액',
                        color='회사',
                        barmode='group',
                        title="주요 재무지표 비교 (조원)",
                        labels={'금액': '금액 (조원)'}
                    )
                    fig_abs.update_layout(
                        xaxis_tickangle=-45,
                        height=400,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    charts.append(("주요 재무지표 비교", fig_abs))
    except Exception as e:
        st.error(f"절대값 비교 차트 생성 오류: {e}")
    
    # 3. 분기별 추이 차트들
    try:
        if quarterly_df is not None and not quarterly_df.empty:
            # 영업이익률 추이
            if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
                fig_line = go.Figure()
                colors_map = {'SK에너지': '#E31E24', '경쟁사1': '#1f77b4', '경쟁사2': '#ff7f0e', 
                             '경쟁사3': '#2ca02c', '경쟁사4': '#d62728'}
                
                for company in quarterly_df['회사'].dropna().unique():
                    company_data = quarterly_df[quarterly_df['회사'] == company].copy()
                    company_data = company_data.sort_values('분기')
                    
                    fig_line.add_trace(go.Scatter(
                        x=company_data['분기'],
                        y=company_data['영업이익률'],
                        mode='lines+markers',
                        name=company,
                        line=dict(width=3, color=colors_map.get(company, '#333333')),
                        marker=dict(size=8)
                    ))
                
                fig_line.update_layout(
                    title="분기별 영업이익률 추이",
                    xaxis_title="분기",
                    yaxis_title="영업이익률 (%)",
                    height=400,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                charts.append(("분기별 영업이익률 추이", fig_line))
            
            # 매출액 추이
            if all(col in quarterly_df.columns for col in ['분기', '회사', '매출액']):
                fig_rev = go.Figure()
                
                for company in quarterly_df['회사'].dropna().unique():
                    company_data = quarterly_df[quarterly_df['회사'] == company].copy()
                    company_data = company_data.sort_values('분기')
                    
                    fig_rev.add_trace(go.Scatter(
                        x=company_data['분기'],
                        y=company_data['매출액'],
                        mode='lines+markers',
                        name=company,
                        line=dict(width=3, color=colors_map.get(company, '#333333')),
                        marker=dict(size=8)
                    ))
                
                fig_rev.update_layout(
                    title="분기별 매출액 추이",
                    xaxis_title="분기",
                    yaxis_title="매출액 (조원)",
                    height=400,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                charts.append(("분기별 매출액 추이", fig_rev))
    except Exception as e:
        st.error(f"분기별 추이 차트 생성 오류: {e}")
    
    return charts


def add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """재무분석 결과 섹션 추가 - 개선된 버전"""
    story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
    
    if financial_data is None or (hasattr(financial_data, 'empty') and financial_data.empty):
        story.append(Paragraph("재무 데이터가 제공되지 않았습니다.", BODY_STYLE))
        story.append(Spacer(1, 18))
        return
    
    # 1-1. SK에너지 대비 경쟁사 갭차이 분석표
    story.append(Paragraph("1-1. SK에너지 대비 경쟁사 갭차이 분석", BODY_STYLE))
    story.append(Spacer(1, 8))
    
    # 원시값 컬럼 제외하고 표시용 데이터 준비
    cols_to_show = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
    df_display = financial_data[cols_to_show].copy()
    
    # 테이블 생성 (갭차이 분석)
    max_rows_per_table = 20
    total_rows = len(df_display)
    
    for i in range(0, total_rows, max_rows_per_table):
        end_idx = min(i + max_rows_per_table, total_rows)
        chunk = df_display.iloc[i:end_idx]
        
        table_data = [df_display.columns.tolist()] + chunk.values.tolist()
        tbl = Table(table_data, repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E31E24')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
            ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
        ]))
        story.append(tbl)
        
        if end_idx < total_rows:
            story.append(Spacer(1, 12))
    
    story.append(Spacer(1, 16))
    
    # 1-2. 분기별 재무지표 상세 데이터
    if quarterly_df is not None and not quarterly_df.empty:
        story.append(Paragraph("1-2. 분기별 재무지표 상세 데이터", BODY_STYLE))
        story.append(Spacer(1, 8))
        
        # 분기별 데이터를 회사별로 분리하여 표시
        companies = quarterly_df['회사'].dropna().unique()
        
        for idx, company in enumerate(companies):
            if idx > 0:
                story.append(Spacer(1, 12))
            
            story.append(Paragraph(f"□ {company} 분기별 실적", BODY_STYLE))
            story.append(Spacer(1, 6))
            
            company_data = quarterly_df[quarterly_df['회사'] == company].copy()
            # 분기 순서대로 정렬
            company_data = company_data.sort_values('분기')
            
            # 표시할 컬럼 선택 (회사 컬럼 제외)
            display_cols = [c for c in company_data.columns if c != '회사']
            company_display = company_data[display_cols]
            
            if not company_display.empty:
                table_data = [company_display.columns.tolist()] + company_display.values.tolist()
                tbl = Table(table_data, repeatRows=1)
                tbl.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4472C4')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
                    ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
                    ('FONTSIZE', (0,0), (-1,-1), 7),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor('#F2F2F2'), colors.white]),
                ]))
                story.append(tbl)
    
    story.append(Spacer(1, 18))


def add_charts_section(story, financial_data, quarterly_df, selected_charts, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """시각화 차트 섹션 추가 - 개선된 버전"""
    story.append(Paragraph("2. 시각화 차트 및 분석", HEADING_STYLE))
    
    if not PLOTLY_AVAILABLE:
        story.append(Paragraph("Plotly 라이브러리를 사용할 수 없어 차트를 생성할 수 없습니다.", BODY_STYLE))
        return False
    
    charts_added = False
    chart_counter = 1
    
    # 자동 생성 차트들
    auto_charts = create_financial_charts(financial_data, quarterly_df)
    
    for chart_title, fig in auto_charts:
        try:
            img_bytes = fig_to_png_bytes(fig, width=900, height=450)
            if img_bytes:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                
                story.append(Paragraph(f"2-{chart_counter}. {chart_title}", BODY_STYLE))
                story.append(Spacer(1, 6))
                story.append(RLImage(tmp_path, width=500, height=280))
                story.append(Spacer(1, 16))
                chart_counter += 1
                charts_added = True
                
                try:
                    os.unlink(tmp_path)
                except:
                    pass
        except Exception as e:
            st.error(f"{chart_title} 생성 오류: {e}")
    
    # 외부에서 전달된 추가 차트들
    if selected_charts:
        for idx, fig in enumerate(selected_charts):
            try:
                img_bytes = fig_to_png_bytes(fig, width=900, height=450)
                if img_bytes:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    
                    story.append(Paragraph(f"2-{chart_counter}. 추가 차트 {idx+1}", BODY_STYLE))
                    story.append(Spacer(1, 6))
                    story.append(RLImage(tmp_path, width=500, height=280))
                    story.append(Spacer(1, 16))
                    chart_counter += 1
                    charts_added = True
                    
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
            except Exception as e:
                st.error(f"추가 차트 {idx+1} 생성 오류: {e}")
    
    # 차트가 하나도 없으면 안내 메시지
    if not charts_added:
        story.append(Paragraph("생성 가능한 차트가 없습니다. 재무 데이터 또는 분기별 데이터를 확인해주세요.", BODY_STYLE))
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
        # 표 데이터인 경우 버퍼에 저장
        if '|' in line:
            ascii_buffer.append(line)
            continue
        
        # 버퍼에 표 데이터가 있으면 테이블 생성
        if ascii_buffer:
            tbl = ascii_to_table(ascii_buffer, registered_fonts, header_color)
            if tbl:
                story.append(tbl)
            story.append(Spacer(1, 12))
            ascii_buffer.clear()
        
        # 일반 텍스트 처리
        if typ == 'title':
            story.append(Paragraph(f"<b>{line}</b>", BODY_STYLE))
        else:
            story.append(Paragraph(line, BODY_STYLE))
    
    # 마지막에 남은 표 데이터 처리
    if ascii_buffer:
        tbl = ascii_to_table(ascii_buffer, registered_fonts, header_color)
        if tbl:
            story.append(tbl)
    
    story.append(Spacer(1, 18))


def add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """뉴스 하이라이트 및 종합 분석 섹션 추가"""
    story.append(Paragraph("3. 뉴스 하이라이트 및 종합 분석", HEADING_STYLE))
    
    if news_data is not None and (not hasattr(news_data, 'empty') or not news_data.empty):
        story.append(Paragraph("3-1. 최신 뉴스 하이라이트", BODY_STYLE))
        for i, title in enumerate(news_data["제목"].head(10), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 16))
        
        # AI 종합 분석을 뉴스 섹션에도 추가
        if insights:
            story.append(Paragraph("3-2. AI 종합 분석 및 시사점", BODY_STYLE))
            story.append(Spacer(1, 8))
            
            # AI 인사이트 텍스트 처리 (뉴스 섹션용 - 파란색 헤더)
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
        # 뉴스 데이터가 없는 경우
        story.append(Paragraph("뉴스 데이터가 제공되지 않았습니다.", BODY_STYLE))
        
        # 뉴스 데이터가 없어도 AI 인사이트가 있으면 표시
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
    
    # 스트림릿 환경에서 안전한 폰트 등록
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

    # 1. 재무분석 결과 (개선된 버전)
    add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE)
    
    # 2. 시각화 차트 및 분석 (개선된 버전)
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
