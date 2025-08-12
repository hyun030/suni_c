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
    """Plotly 차트를 PNG 바이트로 변환 - 개선된 버전"""
    if not PLOTLY_AVAILABLE:
        st.error("Plotly 라이브러리를 사용할 수 없습니다.")
        return None
    
    try:
        # 여러 방법으로 이미지 변환 시도
        methods = [
            lambda: fig.to_image(format="png", width=width, height=height, engine="kaleido"),
            lambda: fig.to_image(format="png", width=width, height=height, engine="auto"),
            lambda: fig.to_image(format="png", width=width, height=height)
        ]
        
        for i, method in enumerate(methods):
            try:
                img_bytes = method()
                if img_bytes and len(img_bytes) > 0:
                    st.success(f"차트 이미지 변환 성공 (방법 {i+1})")
                    return img_bytes
            except Exception as method_error:
                st.warning(f"차트 변환 방법 {i+1} 실패: {method_error}")
                continue
        
        # 모든 방법 실패시
        st.error("모든 차트 이미지 변환 방법이 실패했습니다.")
        return None
        
    except Exception as e:
        st.error(f"차트 이미지 변환 전체 실패: {e}")
        return None


def create_financial_charts(financial_data, quarterly_df):
    """재무 데이터로부터 차트를 생성하는 함수 - 디버깅 강화"""
    charts = []
    
    if not PLOTLY_AVAILABLE:
        st.error("Plotly 라이브러리가 사용불가능합니다.")
        return charts
    
    st.info("🔍 차트 생성 시작...")
    
    # 1. 주요 비율 비교 막대 그래프
    try:
        st.info("📊 비율 비교 차트 생성 중...")
        
        if financial_data is not None and not financial_data.empty and '구분' in financial_data.columns:
            st.info(f"재무 데이터 컬럼: {list(financial_data.columns)}")
            st.info(f"재무 데이터 행 수: {len(financial_data)}")
            
            # 비율 데이터 추출 (% 포함된 행들)
            ratio_rows = financial_data[financial_data['구분'].astype(str).str.contains('%', na=False)].copy()
            st.info(f"비율 데이터 행 수: {len(ratio_rows)}")
            
            if not ratio_rows.empty:
                st.info(f"비율 지표들: {list(ratio_rows['구분'])}")
                
                # 주요 지표 순서 정렬
                key_order = ['매출총이익률(%)', '영업이익률(%)', '순이익률(%)', '매출원가율(%)', '판관비율(%)']
                
                # 데이터 변환
                melt_data = []
                company_cols = [c for c in ratio_rows.columns if c not in ['구분'] and not str(c).endswith('_원시값')]
                st.info(f"회사 컬럼들: {company_cols}")
                
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
                            st.info(f"데이터 추가: {metric_name}, {company}, {value}")
                        except (ValueError, TypeError) as ve:
                            st.warning(f"값 변환 실패: {metric_name}, {company}, {value_str} -> {ve}")
                            continue
                
                st.info(f"총 변환된 데이터 포인트: {len(melt_data)}")
                
                if melt_data:
                    bar_df = pd.DataFrame(melt_data)
                    # 주요 지표만 필터링
                    available_metrics = bar_df['지표'].unique()
                    filtered_metrics = [m for m in key_order if m in available_metrics]
                    
                    if filtered_metrics:
                        bar_df_filtered = bar_df[bar_df['지표'].isin(filtered_metrics)]
                        st.info(f"필터링된 지표: {filtered_metrics}")
                        st.info(f"차트용 데이터프레임 모양: {bar_df_filtered.shape}")
                        
                        fig_bar = px.bar(
                            bar_df_filtered, 
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
                        st.success("✅ 비율 비교 차트 생성 완료")
                    else:
                        st.warning("주요 지표가 데이터에 없습니다.")
                else:
                    st.warning("변환된 데이터가 없습니다.")
            else:
                st.warning("비율 데이터가 없습니다.")
        else:
            st.warning("재무 데이터가 없거나 '구분' 컬럼이 없습니다.")
            
    except Exception as e:
        st.error(f"비율 비교 차트 생성 오류: {e}")
        import traceback
        st.error(traceback.format_exc())
    
    # 2. 절대값 비교 (매출액, 영업이익 등)
    try:
        st.info("📊 절대값 비교 차트 생성 중...")
        
        if financial_data is not None and not financial_data.empty:
            # 절대값 지표들 (조원 단위)
            absolute_metrics = ['매출액(조원)', '영업이익(조원)', '순이익(조원)', '총자산(조원)']
            absolute_rows = financial_data[financial_data['구분'].isin(absolute_metrics)].copy()
            
            st.info(f"절대값 데이터 행 수: {len(absolute_rows)}")
            
            if not absolute_rows.empty:
                melt_abs = []
                company_cols = [c for c in absolute_rows.columns if c not in ['구분'] and not str(c).endswith('_원시값')]
                
                for _, row in absolute_rows.iterrows():
                    metric_name = row['구분']
                    for company in company_cols:
                        try:
                            value_str = str(row[company]).replace('조원', '').replace(',', '').strip()
                            value = float(value_str)
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
                    st.success("✅ 절대값 비교 차트 생성 완료")
                    
    except Exception as e:
        st.error(f"절대값 비교 차트 생성 오류: {e}")
    
    # 3. 분기별 추이 차트들
    try:
        st.info("📊 분기별 추이 차트 생성 중...")
        
        if quarterly_df is not None and not quarterly_df.empty:
            st.info(f"분기별 데이터 컬럼: {list(quarterly_df.columns)}")
            st.info(f"분기별 데이터 모양: {quarterly_df.shape}")
            
            colors_map = {'SK에너지': '#E31E24', '경쟁사1': '#1f77b4', '경쟁사2': '#ff7f0e', 
                         '경쟁사3': '#2ca02c', '경쟁사4': '#d62728'}
            
            # 영업이익률 추이
            if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
                fig_line = go.Figure()
                
                companies = quarterly_df['회사'].dropna().unique()
                st.info(f"분기별 데이터 회사들: {list(companies)}")
                
                for company in companies:
                    company_data = quarterly_df[quarterly_df['회사'] == company].copy()
                    company_data = company_data.sort_values('분기')
                    
                    st.info(f"{company} 데이터 포인트: {len(company_data)}")
                    
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
                st.success("✅ 영업이익률 추이 차트 생성 완료")
            
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
                st.success("✅ 매출액 추이 차트 생성 완료")
                
    except Exception as e:
        st.error(f"분기별 추이 차트 생성 오류: {e}")
    
    st.info(f"총 {len(charts)}개 차트 생성 완료")
    return charts 오류: {e}")
    
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


def create_adaptive_table(data, registered_fonts, header_color='#E31E24', max_col_width=80):
    """화면 크기에 맞춰 자동으로 조정되는 테이블 생성"""
    if data is None or data.empty:
        return None
    
    # 컬럼 수가 많으면 세로형으로 변환
    if len(data.columns) > 6:
        # 가로가 긴 테이블을 세로형으로 변환
        melted_data = []
        for _, row in data.iterrows():
            for col in data.columns:
                melted_data.append([row.get('구분', ''), col, str(row[col])])
        
        table_data = [['지표', '회사', '값']] + melted_data
        col_widths = [120, 80, 100]
    else:
        # 일반 테이블
        table_data = [data.columns.tolist()] + data.values.tolist()
        # 컬럼 너비 자동 조정
        col_widths = [min(max_col_width, max(len(str(col))*6 + 20, 60)) for col in data.columns]
    
    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor(header_color)),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), registered_fonts.get('KoreanBold', 'Helvetica-Bold')),
        ('FONTNAME', (0,1), (-1,-1), registered_fonts.get('Korean', 'Helvetica')),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
    ]))
    
    return tbl


def add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """재무분석 결과 섹션 추가 - 개선된 버전"""
    story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
    
    # 1-1. 분기별 재무지표 상세 데이터 (순서 변경)
    if quarterly_df is not None and not quarterly_df.empty:
        story.append(Paragraph("1-1. 분기별 재무지표 상세 데이터", BODY_STYLE))
        story.append(Spacer(1, 8))
        
        # 분기별 데이터를 회사별로 분리하여 표시
        companies = quarterly_df['회사'].dropna().unique()
        
        for idx, company in enumerate(companies):
            if idx > 0:
                story.append(Spacer(1, 12))
            
            story.append(Paragraph(f"□ {company} 분기별 실적", BODY_STYLE))
            story.append(Spacer(1, 6))
            
            company_data = quarterly_df[quarterly_df['회사'] == company].copy()
            company_data = company_data.sort_values('분기')
            
            # 표시할 컬럼 선택 (회사 컬럼 제외)
            display_cols = [c for c in company_data.columns if c != '회사']
            company_display = company_data[display_cols]
            
            if not company_display.empty:
                # 적응형 테이블 생성
                tbl = create_adaptive_table(company_display, registered_fonts, '#4472C4')
                if tbl:
                    story.append(tbl)
        
        story.append(Spacer(1, 16))
    else:
        story.append(Paragraph("1-1. 분기별 재무지표 상세 데이터", BODY_STYLE))
        story.append(Paragraph("분기별 재무 데이터가 제공되지 않았습니다.", BODY_STYLE))
        story.append(Spacer(1, 16))
    
    # 1-2. SK에너지 대비 경쟁사 갭차이 분석 (순서 변경)
    if financial_data is not None and (not hasattr(financial_data, 'empty') or not financial_data.empty):
        story.append(Paragraph("1-2. SK에너지 대비 경쟁사 갭차이 분석", BODY_STYLE))
        story.append(Spacer(1, 8))
        
        # 원시값 컬럼 제외하고 표시용 데이터 준비
        cols_to_show = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
        df_display = financial_data[cols_to_show].copy()
        
        # 적응형 테이블 생성
        tbl = create_adaptive_table(df_display, registered_fonts, '#E31E24')
        if tbl:
            story.append(tbl)
    else:
        story.append(Paragraph("1-2. SK에너지 대비 경쟁사 갭차이 분석", BODY_STYLE))
        story.append(Paragraph("갭차이 분석 데이터가 제공되지 않았습니다.", BODY_STYLE))
    
    story.append(Spacer(1, 18))


def add_charts_section(story, financial_data, quarterly_df, selected_charts, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """시각화 차트 섹션 추가 - 강화된 디버깅 버전"""
    story.append(Paragraph("2. 시각화 차트 및 분석", HEADING_STYLE))
    
    if not PLOTLY_AVAILABLE:
        story.append(Paragraph("Plotly 라이브러리를 사용할 수 없어 차트를 생성할 수 없습니다.", BODY_STYLE))
        st.error("❌ Plotly 라이브러리 없음")
        return False
    
    st.info("🎯 차트 섹션 생성 시작...")
    charts_added = False
    chart_counter = 1
    
    # 자동 생성 차트들
    try:
        st.info("🔄 자동 차트 생성 중...")
        auto_charts = create_financial_charts(financial_data, quarterly_df)
        st.info(f"생성된 자동 차트 수: {len(auto_charts)}")
        
        for chart_title, fig in auto_charts:
            try:
                st.info(f"📊 {chart_title} 이미지 변환 중...")
                img_bytes = fig_to_png_bytes(fig, width=900, height=450)
                
                if img_bytes and len(img_bytes) > 0:
                    st.info(f"✅ {chart_title} 이미지 변환 성공 (크기: {len(img_bytes)} bytes)")
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    
                    # PDF에 차트 추가
                    story.append(Paragraph(f"2-{chart_counter}. {chart_title}", BODY_STYLE))
                    story.append(Spacer(1, 6))
                    
                    # 이미지 크기 조정하여 PDF에 맞춤
                    try:
                        story.append(RLImage(tmp_path, width=480, height=270))  # 크기 조정
                        charts_added = True
                        chart_counter += 1
                        st.success(f"✅ {chart_title} PDF에 추가 완료")
                    except Exception as img_error:
                        st.error(f"❌ {chart_title} PDF 이미지 추가 실패: {img_error}")
                    
                    story.append(Spacer(1, 16))
                    
                    # 임시 파일 정리
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                        
                else:
                    st.error(f"❌ {chart_title} 이미지 변환 실패 - 빈 데이터")
                    
            except Exception as chart_error:
                st.error(f"❌ {chart_title} 처리 오류: {chart_error}")
                import traceback
                st.error(traceback.format_exc())
                
    except Exception as auto_error:
        st.error(f"❌ 자동 차트 생성 전체 오류: {auto_error}")
    
    # 외부에서 전달된 추가 차트들
    if selected_charts:
        st.info(f"📊 외부 차트 {len(selected_charts)}개 처리 중...")
        
        for idx, fig in enumerate(selected_charts):
            try:
                chart_name = f"추가 차트 {idx+1}"
                st.info(f"🔄 {chart_name} 처리 중...")
                
                img_bytes = fig_to_png_bytes(fig, width=900, height=450)
                if img_bytes and len(img_bytes) > 0:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    
                    story.append(Paragraph(f"2-{chart_counter}. {chart_name}", BODY_STYLE))
                    story.append(Spacer(1, 6))
                    story.append(RLImage(tmp_path, width=480, height=270))
                    story.append(Spacer(1, 16))
                    chart_counter += 1
                    charts_added = True
                    
                    st.success(f"✅ {chart_name} 추가 완료")
                    
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                else:
                    st.error(f"❌ {chart_name} 이미지 변환 실패")
                    
            except Exception as ext_chart_error:
                st.error(f"❌ 추가 차트 {idx+1} 처리 오류: {ext_chart_error}")
    
    # 차트가 하나도 없으면 안내 메시지
    if not charts_added:
        st.warning("⚠️ 차트가 하나도 생성되지 않았습니다.")
        story.append(Paragraph("생성 가능한 차트가 없습니다. 재무 데이터 구조를 확인해주세요.", BODY_STYLE))
        
        # 디버깅 정보 추가
        if financial_data is not None:
            story.append(Paragraph(f"재무 데이터 정보: {len(financial_data)}행, 컬럼: {list(financial_data.columns)[:5]}...", BODY_STYLE))
        if quarterly_df is not None:
            story.append(Paragraph(f"분기별 데이터 정보: {len(quarterly_df)}행, 컬럼: {list(quarterly_df.columns)[:5]}...", BODY_STYLE))
            
        story.append(Spacer(1, 18))
    else:
        st.success(f"✅ 총 {chart_counter-1}개 차트가 PDF에 추가되었습니다.")
    
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
    insights=None,
    selected_charts=None,
    quarterly_df=None,
    show_footer=False,
    report_target="SK이노베이션 경영진",
    report_author="보고자 미기재",
    font_paths=None,
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
