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


def clean_ai_text(raw):
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


def fig_to_png_bytes(fig, width=900, height=450):
    """Plotly 차트를 PNG 바이트로 변환 - 개선된 버전"""
    if not PLOTLY_AVAILABLE:
        st.error("Plotly 라이브러리를 사용할 수 없습니다.")
        return None
    
    try:
        # Kaleido 사용 (가장 안정적)
        img_bytes = fig.to_image(format="png", width=width, height=height, engine="kaleido")
        if img_bytes and len(img_bytes) > 0:
            st.success(f"차트 이미지 변환 성공 (Kaleido)")
            return img_bytes
    except Exception as e:
        st.warning(f"Kaleido 변환 실패: {e}")
        
    try:
        # 기본 엔진 사용
        img_bytes = fig.to_image(format="png", width=width, height=height)
        if img_bytes and len(img_bytes) > 0:
            st.success(f"차트 이미지 변환 성공 (기본 엔진)")
            return img_bytes
    except Exception as e:
        st.error(f"차트 이미지 변환 실패: {e}")
        return None


def create_key_value_display(data, registered_fonts, title="데이터", color='#E31E24'):
    """테이블 대신 키-값 형식으로 데이터 표시"""
    if data is None or data.empty:
        return []
    
    elements = []
    
    # 제목 스타일
    title_style = ParagraphStyle(
        'KeyValueTitle',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=11,
        textColor=colors.HexColor(color),
        spaceAfter=8,
        leftIndent=0
    )
    
    # 항목 스타일
    item_style = ParagraphStyle(
        'KeyValueItem',
        fontName=registered_fonts.get('Korean', 'Helvetica'),
        fontSize=10,
        leading=14,
        leftIndent=20,
        spaceAfter=4
    )
    
    elements.append(Paragraph(f"■ {title}", title_style))
    
    # 회사별로 데이터 정리
    if '구분' in data.columns:
        companies = [col for col in data.columns if col not in ['구분'] and not str(col).endswith('_원시값')]
        
        for company in companies:
            elements.append(Paragraph(f"□ {company}", item_style))
            
            for _, row in data.iterrows():
                metric = row['구분']
                value = row[company]
                if pd.notna(value) and str(value).strip():
                    elements.append(Paragraph(f"  • {metric}: {value}", item_style))
            
            elements.append(Spacer(1, 6))
    else:
        # 일반적인 키-값 표시
        for col in data.columns:
            if not str(col).endswith('_원시값'):
                for _, row in data.iterrows():
                    value = row[col]
                    if pd.notna(value) and str(value).strip():
                        elements.append(Paragraph(f"  • {col}: {value}", item_style))
    
    return elements


def create_bullet_summary(data, registered_fonts, title="요약", max_items=8):
    """데이터를 불릿 포인트 형태로 요약 표시"""
    if data is None or data.empty:
        return []
    
    elements = []
    
    # 제목 스타일
    title_style = ParagraphStyle(
        'BulletTitle',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=12,
        textColor=colors.HexColor('#2E86AB'),
        spaceAfter=10,
    )
    
    # 불릿 스타일
    bullet_style = ParagraphStyle(
        'BulletItem',
        fontName=registered_fonts.get('Korean', 'Helvetica'),
        fontSize=10,
        leading=16,
        leftIndent=15,
        bulletIndent=5,
        bulletFontName=registered_fonts.get('Korean', 'Helvetica'),
        spaceAfter=3
    )
    
    elements.append(Paragraph(f"📊 {title}", title_style))
    
    count = 0
    if '구분' in data.columns:
        companies = [col for col in data.columns if col not in ['구분'] and not str(col).endswith('_원시값')]
        
        for _, row in data.iterrows():
            if count >= max_items:
                break
                
            metric = row['구분']
            # 주요 지표만 선택
            if any(keyword in metric for keyword in ['이익률', '매출액', '순이익', '총자산', '부채비율']):
                values = []
                for company in companies:
                    value = row[company]
                    if pd.notna(value) and str(value).strip():
                        values.append(f"{company}: {value}")
                
                if values:
                    bullet_text = f"• {metric} → {', '.join(values)}"
                    elements.append(Paragraph(bullet_text, bullet_style))
                    count += 1
    
    if count == 0:
        elements.append(Paragraph("• 표시할 주요 지표가 없습니다.", bullet_style))
    
    return elements


def create_enhanced_financial_charts(financial_data, quarterly_df):
    """재무 데이터로부터 차트를 생성하는 함수 - 안정성 개선"""
    charts = []
    
    if not PLOTLY_AVAILABLE:
        st.error("Plotly 라이브러리가 사용불가능합니다.")
        return charts
    
    st.info("🔍 차트 생성 시작...")
    
    # 차트 기본 설정
    chart_config = {
        'displayModeBar': False,
        'responsive': True
    }
    
    # 색상 팔레트
    colors_palette = ['#E31E24', '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
    
    # 1. 주요 비율 비교 막대 그래프 (안정성 개선)
    try:
        st.info("📊 수익성 지표 차트 생성 중...")
        
        if financial_data is not None and not financial_data.empty and '구분' in financial_data.columns:
            # 비율 데이터 추출
            ratio_keywords = ['이익률', '%']
            ratio_rows = financial_data[
                financial_data['구분'].astype(str).str.contains('|'.join(ratio_keywords), na=False)
            ].copy()
            
            if not ratio_rows.empty:
                # 주요 지표 우선 선택
                priority_metrics = ['매출총이익률(%)', '영업이익률(%)', '순이익률(%)']
                available_metrics = ratio_rows['구분'].tolist()
                selected_metrics = [m for m in priority_metrics if m in available_metrics]
                
                if not selected_metrics:
                    selected_metrics = available_metrics[:5]  # 상위 5개
                
                filtered_data = ratio_rows[ratio_rows['구분'].isin(selected_metrics)]
                
                # 데이터 변환
                company_cols = [c for c in filtered_data.columns 
                              if c not in ['구분'] and not str(c).endswith('_원시값')]
                
                if company_cols:
                    melt_data = []
                    for _, row in filtered_data.iterrows():
                        metric = row['구분']
                        for company in company_cols:
                            value_str = str(row[company]).replace('%', '').replace(',', '').strip()
                            try:
                                value = float(value_str)
                                melt_data.append({
                                    '지표': metric,
                                    '회사': company,
                                    '수치': value
                                })
                            except (ValueError, TypeError):
                                continue
                    
                    if melt_data:
                        df_chart = pd.DataFrame(melt_data)
                        fig_bar = px.bar(
                            df_chart,
                            x='지표',
                            y='수치',
                            color='회사',
                            barmode='group',
                            title="주요 수익성 지표 비교",
                            labels={'수치': '비율 (%)'},
                            color_discrete_sequence=colors_palette
                        )
                        
                        fig_bar.update_layout(
                            xaxis_tickangle=-30,
                            height=400,
                            font=dict(size=10),
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1
                            ),
                            plot_bgcolor='white',
                            paper_bgcolor='white'
                        )
                        
                        charts.append(("주요 수익성 지표 비교", fig_bar))
                        st.success("✅ 수익성 지표 차트 생성 완료")
                        
    except Exception as e:
        st.error(f"수익성 지표 차트 생성 오류: {e}")
    
    # 2. 매출액/자산 규모 비교 (개선)
    try:
        st.info("📊 규모 지표 차트 생성 중...")
        
        if financial_data is not None and not financial_data.empty:
            size_keywords = ['매출액', '총자산', '자본']
            size_rows = financial_data[
                financial_data['구분'].astype(str).str.contains('|'.join(size_keywords), na=False)
            ].copy()
            
            if not size_rows.empty:
                company_cols = [c for c in size_rows.columns 
                              if c not in ['구분'] and not str(c).endswith('_원시값')]
                
                if company_cols:
                    melt_data = []
                    for _, row in size_rows.iterrows():
                        metric = row['구분']
                        for company in company_cols:
                            value_str = str(row[company]).replace('조원', '').replace(',', '').strip()
                            try:
                                value = float(value_str)
                                melt_data.append({
                                    '지표': metric,
                                    '회사': company,
                                    '금액': value
                                })
                            except (ValueError, TypeError):
                                continue
                    
                    if melt_data:
                        df_size = pd.DataFrame(melt_data)
                        fig_size = px.bar(
                            df_size,
                            x='지표',
                            y='금액',
                            color='회사',
                            barmode='group',
                            title="주요 규모 지표 비교 (조원)",
                            labels={'금액': '금액 (조원)'},
                            color_discrete_sequence=colors_palette
                        )
                        
                        fig_size.update_layout(
                            xaxis_tickangle=-30,
                            height=400,
                            font=dict(size=10),
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1
                            ),
                            plot_bgcolor='white',
                            paper_bgcolor='white'
                        )
                        
                        charts.append(("주요 규모 지표 비교", fig_size))
                        st.success("✅ 규모 지표 차트 생성 완료")
                        
    except Exception as e:
        st.error(f"규모 지표 차트 생성 오류: {e}")
    
    # 3. 분기별 추이 차트 (개선)
    try:
        st.info("📊 분기별 추이 차트 생성 중...")
        
        if quarterly_df is not None and not quarterly_df.empty:
            required_cols = ['분기', '회사']
            if all(col in quarterly_df.columns for col in required_cols):
                companies = quarterly_df['회사'].dropna().unique()
                
                # 영업이익률 추이
                if '영업이익률' in quarterly_df.columns:
                    fig_trend = go.Figure()
                    
                    for i, company in enumerate(companies):
                        company_data = quarterly_df[quarterly_df['회사'] == company].copy()
                        company_data = company_data.sort_values('분기')
                        
                        fig_trend.add_trace(go.Scatter(
                            x=company_data['분기'],
                            y=company_data['영업이익률'],
                            mode='lines+markers',
                            name=company,
                            line=dict(width=3, color=colors_palette[i % len(colors_palette)]),
                            marker=dict(size=8)
                        ))
                    
                    fig_trend.update_layout(
                        title="분기별 영업이익률 추이",
                        xaxis_title="분기",
                        yaxis_title="영업이익률 (%)",
                        height=400,
                        font=dict(size=10),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1
                        ),
                        plot_bgcolor='white',
                        paper_bgcolor='white',
                        xaxis=dict(showgrid=True, gridwidth=1, gridcolor='LightGray'),
                        yaxis=dict(showgrid=True, gridwidth=1, gridcolor='LightGray')
                    )
                    
                    charts.append(("분기별 영업이익률 추이", fig_trend))
                    st.success("✅ 분기별 추이 차트 생성 완료")
                    
    except Exception as e:
        st.error(f"분기별 추이 차트 생성 오류: {e}")
    
    st.info(f"총 {len(charts)}개 차트 생성 완료")
    return charts


def add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """재무분석 결과 섹션 추가 - 표 대신 다양한 형식 사용"""
    story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
    
    # 1-1. 핵심 재무지표 요약 (불릿 포인트 형식)
    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1-1. 핵심 재무지표 요약", BODY_STYLE))
        story.append(Spacer(1, 8))
        
        # 불릿 포인트 형태로 요약 표시
        bullet_elements = create_bullet_summary(financial_data, registered_fonts, "주요 지표")
        for element in bullet_elements:
            story.append(element)
        
        story.append(Spacer(1, 16))
    
    # 1-2. 상세 재무 데이터 (키-값 형식)
    if financial_data is not None and not financial_data.empty:
        story.append(Paragraph("1-2. SK에너지 대비 경쟁사 갭차이 분석", BODY_STYLE))
        story.append(Spacer(1, 8))
        
        # 키-값 형식으로 표시
        kv_elements = create_key_value_display(financial_data, registered_fonts, "재무지표 비교", '#E31E24')
        for element in kv_elements:
            story.append(element)
        
        story.append(Spacer(1, 16))
    
    # 1-3. 분기별 성과 추이
    if quarterly_df is not None and not quarterly_df.empty:
        story.append(Paragraph("1-3. 분기별 성과 추이", BODY_STYLE))
        story.append(Spacer(1, 8))
        
        # 분기별 데이터를 요약 형식으로 표시
        companies = quarterly_df['회사'].dropna().unique()
        
        for company in companies:
            company_data = quarterly_df[quarterly_df['회사'] == company].copy()
            if not company_data.empty:
                kv_elements = create_key_value_display(
                    company_data.drop('회사', axis=1), 
                    registered_fonts, 
                    f"{company} 분기별 실적", 
                    '#4472C4'
                )
                for element in kv_elements:
                    story.append(element)
                story.append(Spacer(1, 8))
        
        story.append(Spacer(1, 16))
    
    story.append(Spacer(1, 18))


def add_charts_section(story, financial_data, quarterly_df, selected_charts, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """시각화 차트 섹션 추가 - 안정성 개선"""
    story.append(Paragraph("2. 시각화 차트 및 분석", HEADING_STYLE))
    
    if not PLOTLY_AVAILABLE:
        story.append(Paragraph("시각화 라이브러리를 사용할 수 없어 차트를 생성할 수 없습니다.", BODY_STYLE))
        story.append(Paragraph("대신 데이터 요약을 제공합니다.", BODY_STYLE))
        
        # 차트 대신 데이터 요약 제공
        if financial_data is not None and not financial_data.empty:
            summary_elements = create_bullet_summary(financial_data, registered_fonts, "데이터 하이라이트")
            for element in summary_elements:
                story.append(element)
        
        story.append(Spacer(1, 18))
        return False
    
    st.info("🎯 차트 섹션 생성 시작...")
    charts_added = False
    chart_counter = 1
    
    # 자동 생성 차트들 (안정성 개선)
    try:
        auto_charts = create_enhanced_financial_charts(financial_data, quarterly_df)
        st.info(f"생성된 자동 차트 수: {len(auto_charts)}")
        
        for chart_title, fig in auto_charts:
            try:
                st.info(f"📊 {chart_title} 이미지 변환 중...")
                img_bytes = fig_to_png_bytes(fig, width=800, height=400)  # 크기 최적화
                
                if img_bytes and len(img_bytes) > 0:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    
                    # PDF에 차트 추가
                    story.append(Paragraph(f"2-{chart_counter}. {chart_title}", BODY_STYLE))
                    story.append(Spacer(1, 6))
                    
                    try:
                        # 이미지 크기를 페이지에 맞게 조정
                        story.append(RLImage(tmp_path, width=480, height=240))
                        charts_added = True
                        chart_counter += 1
                        st.success(f"✅ {chart_title} PDF에 추가 완료")
                    except Exception as img_error:
                        st.error(f"❌ {chart_title} PDF 이미지 추가 실패: {img_error}")
                        # 차트 대신 텍스트 설명 추가
                        story.append(Paragraph(f"[차트: {chart_title} - 이미지 로드 실패]", BODY_STYLE))
                    
                    story.append(Spacer(1, 12))
                    
                    # 임시 파일 정리
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                        
                else:
                    st.error(f"❌ {chart_title} 이미지 변환 실패")
                    # 차트 대신 데이터 요약 추가
                    story.append(Paragraph(f"2-{chart_counter}. {chart_title} (데이터 요약)", BODY_STYLE))
                    story.append(Paragraph("차트 생성에 실패했지만 관련 데이터는 1장에서 확인할 수 있습니다.", BODY_STYLE))
                    story.append(Spacer(1, 12))
                    chart_counter += 1
                    
            except Exception as chart_error:
                st.error(f"❌ {chart_title} 처리 오류: {chart_error}")
                
    except Exception as auto_error:
        st.error(f"❌ 자동 차트 생성 전체 오류: {auto_error}")
    
    # 외부에서 전달된 추가 차트들
    if selected_charts:
        st.info(f"📊 외부 차트 {len(selected_charts)}개 처리 중...")
        
        for idx, fig in enumerate(selected_charts):
            try:
                chart_name = f"추가 차트 {idx+1}"
                img_bytes = fig_to_png_bytes(fig, width=800, height=400)
                
                if img_bytes and len(img_bytes) > 0:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    
                    story.append(Paragraph(f"2-{chart_counter}. {chart_name}", BODY_STYLE))
                    story.append(Spacer(1, 6))
                    story.append(RLImage(tmp_path, width=480, height=240))
                    story.append(Spacer(1, 12))
                    chart_counter += 1
                    charts_added = True
                    
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                else:
                    st.error(f"❌ {chart_name} 이미지 변환 실패")
                    
            except Exception as ext_chart_error:
                st.error(f"❌ 추가 차트 {idx+1} 처리 오류: {ext_chart_error}")
    
    # 차트 섹션 마무리
    if not charts_added:
        st.warning("⚠️ 차트가 생성되지 않았습니다.")
        story.append(Paragraph("시각화 차트 생성에 실패했습니다. 아래 데이터 요약을 참고하세요.", BODY_STYLE))
        
        # 차트 대신 핵심 인사이트 제공
        if financial_data is not None and not financial_data.empty:
            insight_elements = create_bullet_summary(
                financial_data, registered_fonts, 
                "주요 재무지표 인사이트", max_items=6
            )
            for element in insight_elements:
                story.append(element)
                
        story.append(Spacer(1, 18))
    else:
        st.success(f"✅ 총 {chart_counter-1}개 차트가 PDF에 추가되었습니다.")
    
    return charts_added


def create_narrative_summary(data, registered_fonts, title="분석 요약"):
    """데이터를 문장형 요약으로 표시"""
    if data is None or data.empty:
        return []
    
    elements = []
    
    # 제목 스타일
    title_style = ParagraphStyle(
        'NarrativeTitle',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=11,
        textColor=colors.HexColor('#2E8B57'),
        spaceAfter=8,
    )
    
    # 본문 스타일
    narrative_style = ParagraphStyle(
        'NarrativeBody',
        fontName=registered_fonts.get('Korean', 'Helvetica'),
        fontSize=10,
        leading=16,
        spaceAfter=8,
        leftIndent=10,
        rightIndent=10
    )
    
    elements.append(Paragraph(f"📝 {title}", title_style))
    
    if '구분' in data.columns:
        companies = [col for col in data.columns if col not in ['구분'] and not str(col).endswith('_원시값')]
        
        # 주요 지표별 분석
        key_metrics = []
        for _, row in data.iterrows():
            metric = row['구분']
            if any(keyword in metric for keyword in ['이익률', '매출액', '순이익']):
                values = {}
                for company in companies:
                    value = row[company]
                    if pd.notna(value) and str(value).strip():
                        values[company] = str(value)
                
                if values:
                    if len(values) > 1:
                        # 비교 분석
                        companies_list = list(values.keys())
                        if 'SK에너지' in companies_list:
                            sk_value = values.get('SK에너지', '')
                            others = [(k, v) for k, v in values.items() if k != 'SK에너지']
                            if others:
                                other_desc = ', '.join([f"{k} {v}" for k, v in others[:2]])
                                narrative = f"{metric}에서 SK에너지는 {sk_value}를 기록했으며, 경쟁사는 {other_desc} 수준입니다."
                            else:
                                narrative = f"{metric}에서 SK에너지는 {sk_value}를 달성했습니다."
                        else:
                            # SK에너지가 없는 경우
                            top_companies = list(values.items())[:3]
                            desc = ', '.join([f"{k} {v}" for k, v in top_companies])
                            narrative = f"{metric} 현황: {desc}"
                        
                        key_metrics.append(narrative)
        
        # 요약문 생성
        if key_metrics:
            summary_text = " ".join(key_metrics[:4])  # 최대 4개 지표
            elements.append(Paragraph(summary_text, narrative_style))
        else:
            elements.append(Paragraph("분석 가능한 핵심 지표가 제한적입니다.", narrative_style))
    
    return elements


def add_ai_insights_section(story, insights, registered_fonts, BODY_STYLE, header_color='#E31E24'):
    """AI 인사이트 섹션 추가 - 텍스트 중심으로 개선"""
    if not insights:
        story.append(Paragraph("2-AI. 분석 인사이트", BODY_STYLE))
        story.append(Paragraph("AI 인사이트가 제공되지 않았습니다.", BODY_STYLE))
        story.append(Spacer(1, 18))
        return
    
    story.append(Paragraph("2-AI. 분석 인사이트", BODY_STYLE))
    story.append(Spacer(1, 8))

    # AI 인사이트를 섹션별로 나누어 처리
    insights_text = str(insights)
    sections = insights_text.split('\n\n')  # 빈 줄로 섹션 구분
    
    # 인사이트 스타일
    insight_style = ParagraphStyle(
        'InsightBody',
        fontName=registered_fonts.get('Korean', 'Helvetica'),
        fontSize=10,
        leading=16,
        spaceAfter=8,
        leftIndent=15,
        rightIndent=15,
        backColor=colors.HexColor('#FFF8F0')  # 연한 배경색
    )
    
    insight_header_style = ParagraphStyle(
        'InsightHeader',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=11,
        leading=16,
        spaceAfter=6,
        textColor=colors.HexColor(header_color)
    )
    
    section_count = 1
    for section in sections:
        section = section.strip()
        if not section:
            continue
            
        lines = section.split('\n')
        
        # 테이블 데이터 확인
        if any('|' in line for line in lines):
            # 테이블이 있는 경우 - 텍스트로 변환
            story.append(Paragraph(f"□ 분석 {section_count}", insight_header_style))
            
            table_data = []
            for line in lines:
                if '|' in line and line.strip():
                    cols = [col.strip() for col in line.split('|') if col.strip()]
                    if cols:
                        table_data.append(cols)
            
            # 테이블을 텍스트로 변환
            if len(table_data) > 1:  # 헤더 + 데이터
                headers = table_data[0]
                for i, row in enumerate(table_data[2:], 1):  # 구분선 건너뛰기
                    if len(row) == len(headers):
                        row_text = []
                        for j, cell in enumerate(row):
                            if j < len(headers):
                                row_text.append(f"{headers[j]}: {cell}")
                        
                        if row_text:
                            story.append(Paragraph(f"• {', '.join(row_text)}", insight_style))
            
        else:
            # 일반 텍스트 처리
            if lines:
                first_line = lines[0]
                # 제목인지 확인
                if (first_line.strip() and 
                    (first_line.startswith(('1.', '2.', '3.', '4.', '5.')) or 
                     len(first_line) < 50)):
                    story.append(Paragraph(f"□ {first_line}", insight_header_style))
                    remaining_lines = lines[1:]
                else:
                    story.append(Paragraph(f"□ 분석 {section_count}", insight_header_style))
                    remaining_lines = lines
                
                # 나머지 내용
                for line in remaining_lines:
                    line = line.strip()
                    if line:
                        # 마크다운 문자 정리
                        clean_line = re.sub(r'[*_#>`~]', '', line)
                        story.append(Paragraph(f"  {clean_line}", insight_style))
        
        story.append(Spacer(1, 10))
        section_count += 1
    
    story.append(Spacer(1, 18))


def add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """뉴스 하이라이트 및 종합 분석 섹션 추가 - 가독성 개선"""
    story.append(Paragraph("3. 뉴스 하이라이트 및 종합 분석", HEADING_STYLE))
    
    # 뉴스 항목 스타일
    news_style = ParagraphStyle(
        'NewsItem',
        fontName=registered_fonts.get('Korean', 'Helvetica'),
        fontSize=10,
        leading=16,
        spaceAfter=6,
        leftIndent=20,
        bulletIndent=10
    )
    
    news_header_style = ParagraphStyle(
        'NewsHeader',
        fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
        fontSize=11,
        textColor=colors.HexColor('#1E3A8A'),
        spaceAfter=8,
    )
    
    if news_data is not None and (not hasattr(news_data, 'empty') or not news_data.empty):
        story.append(Paragraph("3-1. 최신 뉴스 하이라이트", news_header_style))
        
        # 뉴스를 카테고리별로 분류해서 표시 (더 체계적으로)
        news_items = news_data["제목"].head(8).tolist()  # 상위 8개
        
        for i, title in enumerate(news_items, 1):
            # 뉴스 제목 길이 조정
            if len(title) > 80:
                title = title[:77] + "..."
            story.append(Paragraph(f"• {title}", news_style))
        
        story.append(Spacer(1, 16))
        
        # 종합 분석
        if insights:
            story.append(Paragraph("3-2. 종합 분석 및 시사점", news_header_style))
            story.append(Spacer(1, 8))
            
            # 인사이트를 요약 형태로 재구성
            summary_style = ParagraphStyle(
                'SummaryStyle',
                fontName=registered_fonts.get('Korean', 'Helvetica'),
                fontSize=10,
                leading=18,
                spaceAfter=10,
                leftIndent=10,
                rightIndent=10,
                backColor=colors.HexColor('#F0F9FF'),
                borderColor=colors.HexColor('#0EA5E9'),
                borderWidth=1,
                borderPadding=8
            )
            
            # AI 인사이트를 간단한 요약으로 변환
            insights_lines = str(insights).split('\n')
            key_points = []
            
            for line in insights_lines:
                line = line.strip()
                if line and not line.startswith('|') and len(line) > 10:
                    # 마크다운 제거 및 정리
                    clean_line = re.sub(r'[*_#>`~]', '', line)
                    if any(keyword in clean_line for keyword in ['분석', '전망', '시사점', '결론', '요약']):
                        key_points.append(clean_line)
                        if len(key_points) >= 3:  # 최대 3개 포인트
                            break
            
            if key_points:
                summary_text = " ".join(key_points[:2])  # 상위 2개 포인트
                story.append(Paragraph(summary_text, summary_style))
            else:
                story.append(Paragraph("상세한 AI 분석 결과는 2-AI 섹션을 참고하시기 바랍니다.", summary_style))
        else:
            story.append(Paragraph("AI 종합 분석이 제공되지 않았습니다.", BODY_STYLE))
    else:
        # 뉴스 데이터가 없는 경우
        story.append(Paragraph("뉴스 데이터가 제공되지 않았습니다.", BODY_STYLE))
        
        if insights:
            story.append(Paragraph("3-1. 일반 분석 및 시사점", news_header_style))
            story.append(Spacer(1, 8))
            
            # 간단한 텍스트 요약
            brief_style = ParagraphStyle(
                'BriefStyle',
                fontName=registered_fonts.get('Korean', 'Helvetica'),
                fontSize=10,
                leading=16,
                spaceAfter=8,
                leftIndent=15
            )
            
            story.append(Paragraph("제공된 데이터를 바탕으로 한 주요 분석 결과입니다.", brief_style))
    
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
    """향상된 PDF 보고서 생성 - 안정성과 가독성 개선"""
    
    try:
        # 스트림릿 환경에서 안전한 폰트 등록
        registered_fonts = register_fonts_safe()
        
        # 스타일 정의 - 한국어 최적화
        styles = getSampleStyleSheet()
        
        TITLE_STYLE = ParagraphStyle(
            'Title',
            fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
            fontSize=18,
            leading=28,
            spaceAfter=20,
            alignment=1,  # 중앙 정렬
            textColor=colors.HexColor('#1E3A8A')
        )
        
        HEADING_STYLE = ParagraphStyle(
            'Heading',
            fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
            fontSize=13,
            leading=20,
            textColor=colors.HexColor('#E31E24'),
            spaceBefore=18,
            spaceAfter=12,
            borderWidth=1,
            borderColor=colors.HexColor('#E31E24'),
            borderPadding=8,
            backColor=colors.HexColor('#FFF8F8')
        )
        
        BODY_STYLE = ParagraphStyle(
            'Body',
            fontName=registered_fonts.get('Korean', 'Helvetica'),
            fontSize=10,
            leading=16,
            spaceAfter=8,
        )

        # PDF 문서 생성
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=A4, 
            leftMargin=50, 
            rightMargin=50, 
            topMargin=50, 
            bottomMargin=60
        )

        story = []
        
        # 표지 섹션
        story.append(Paragraph("SK에너지 경쟁력 분석 보고서", TITLE_STYLE))
        story.append(Spacer(1, 15))
        story.append(Paragraph("손익개선을 위한 종합 비교분석", TITLE_STYLE))
        story.append(Spacer(1, 30))
        
        # 보고서 정보 박스
        info_style = ParagraphStyle(
            'InfoBox',
            fontName=registered_fonts.get('Korean', 'Helvetica'),
            fontSize=11,
            leading=18,
            alignment=1,
            backColor=colors.HexColor('#F8F9FA'),
            borderColor=colors.HexColor('#6C757D'),
            borderWidth=1,
            borderPadding=15
        )
        
        report_info = f"""
        <b>보고일자:</b> {datetime.now().strftime('%Y년 %m월 %d일')}<br/><br/>
        <b>보고대상:</b> {report_target}<br/><br/>
        <b>보고자:</b> {report_author}
        """
        story.append(Paragraph(report_info, info_style))
        story.append(Spacer(1, 40))

        # 1. 재무분석 결과 (표 대신 다양한 형식)
        add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE)
        
        # 2. 시각화 차트 및 분석 (안정성 개선)
        charts_added = add_charts_section(story, financial_data, quarterly_df, selected_charts, 
                                        registered_fonts, HEADING_STYLE, BODY_STYLE)
        
        # 2-AI. AI 인사이트 (텍스트 중심)
        add_ai_insights_section(story, insights, registered_fonts, BODY_STYLE)
        
        # 3. 뉴스 하이라이트 및 종합 분석 (가독성 개선)
        add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE)

        # 결론 및 권고사항 (추가)
        story.append(Paragraph("4. 결론 및 권고사항", HEADING_STYLE))
        
        conclusion_style = ParagraphStyle(
            'Conclusion',
            fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
            fontSize=11,
            leading=18,
            spaceAfter=10,
            backColor=colors.HexColor('#FFF9C4'),
            borderColor=colors.HexColor('#F59E0B'),
            borderWidth=1,
            borderPadding=12
        )
        
        story.append(Paragraph(
            "본 보고서는 SK에너지의 경쟁력 분석을 통해 손익개선 방안을 제시하고 있습니다. "
            "정기적인 모니터링과 벤치마킹을 통해 지속적인 성과 개선을 추진하시기 바랍니다.",
            conclusion_style
        ))

        # 푸터 (선택사항)
        if show_footer:
            story.append(Spacer(1, 30))
            footer_style = ParagraphStyle(
                'Footer',
                fontName=registered_fonts.get('Korean', 'Helvetica'),
                fontSize=9,
                alignment=1,
                textColor=colors.HexColor('#6C757D')
            )
            footer_text = "※ 본 보고서는 대시보드에서 자동 생성되었습니다. | " + datetime.now().strftime('%Y-%m-%d %H:%M')
            story.append(Paragraph(footer_text, footer_style))

        # 페이지 번호 및 헤더 추가
        def add_page_decorations(canvas, doc):
            # 페이지 번호
            canvas.setFont('Helvetica', 9)
            canvas.drawCentredString(A4[0]/2, 25, f"- {canvas.getPageNumber()} -")
            
            # 헤더 라인
            canvas.setStrokeColor(colors.HexColor('#E31E24'))
            canvas.setLineWidth(2)
            canvas.line(50, A4[1]-35, A4[0]-50, A4[1]-35)

        # PDF 문서 빌드
        doc.build(story, onFirstPage=add_page_decorations, onLaterPages=add_page_decorations)
        buffer.seek(0)
        
        st.success("✅ PDF 보고서 생성 완료!")
        return buffer.getvalue()
        
    except Exception as e:
        st.error(f"❌ PDF 보고서 생성 실패: {e}")
        import traceback
        st.error(traceback.format_exc())
        return None
