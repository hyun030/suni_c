
# -*- coding: utf-8 -*-
"""
완전한 PDF/Excel 보고서 생성 모듈 (kaleido 포함)
설치 필요: pip install kaleido plotly reportlab pandas openpyxl
"""

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

# OpenAI GPT 연동을 위한 import (선택사항)
try:
    import openai
    GPT_AVAILABLE = True
except ImportError:
    GPT_AVAILABLE = False
    print("⚠️ OpenAI 패키지가 없습니다. GPT 기능을 사용하려면 'pip install openai'를 실행하세요.")

# Plotly import 및 kaleido 체크
try:
    import plotly.express as px
    import plotly.graph_objects as go
    import plotly.io as pio
    
    # kaleido 패키지 체크
    try:
        import kaleido
        PLOTLY_AVAILABLE = True
        print("✅ Plotly 및 kaleido 라이브러리 로드 성공")
    except ImportError:
        PLOTLY_AVAILABLE = False
        print("❌ kaleido가 설치되지 않았습니다. 다음 명령어로 설치하세요:")
        print("   pip install kaleido")
        
except ImportError as e:
    PLOTLY_AVAILABLE = False
    print(f"⚠️ Plotly 라이브러리 로드 실패: {e}")
    print("다음 명령어로 설치하세요: pip install plotly kaleido")


def get_company_color(company, companies):
    """회사별 색상 반환 (간단한 색상 매핑)"""
    color_map = {
        'SK에너지': '#E31E24',  # SK 빨강
        'S-Oil': '#0066CC',     # 파랑  
        'GS칼텍스': '#00AA44',  # 초록
        'HD현대오일뱅크': '#FF6600',  # 주황
        '롯데케미칼': '#9900CC'  # 보라
    }
    
    # 기본 색상들
    default_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
    
    for comp in companies:
        if company in comp:
            return color_map.get(comp, default_colors[hash(comp) % len(default_colors)])
    
    return color_map.get(company, default_colors[hash(company) % len(default_colors)])


def create_sk_bar_chart(chart_df: pd.DataFrame):
    """SK에너지 강조 막대 차트"""
    if not PLOTLY_AVAILABLE or chart_df.empty: 
        print("⚠️ Plotly/kaleido가 없거나 데이터가 비어있어 차트를 생성할 수 없습니다.")
        return None
    
    companies = chart_df['회사'].unique()
    color_map = {comp: get_company_color(comp, companies) for comp in companies}
    
    fig = px.bar(
        chart_df, x='구분', y='수치', color='회사',
        title="💼 SK에너지 vs 경쟁사 수익성 지표 비교",
        text='수치', color_discrete_map=color_map, barmode='group', height=450
    )
    fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig.update_layout(
        yaxis_title="수치(%)", xaxis_title="재무 지표", legend_title="회사",
        font=dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif")
    )
    return fig


def create_sk_radar_chart(chart_df):
    """SK에너지 중심 레이더 차트 (지표별 Min-Max 정규화 적용)"""
    if chart_df.empty or not PLOTLY_AVAILABLE:
        print("⚠️ Plotly/kaleido가 없거나 데이터가 비어있어 레이더 차트를 생성할 수 없습니다.")
        return None
    
    companies = chart_df['회사'].unique() if '회사' in chart_df.columns else []
    metrics = chart_df['구분'].unique() if '구분' in chart_df.columns else []
    
    # 지표별 최소, 최대값 계산
    min_max = {}
    for metric in metrics:
        values = chart_df.loc[chart_df['구분'] == metric, '수치']
        min_val = values.min()
        max_val = values.max()
        # 최소 최대값이 같으면 max_val = min_val + 1로 설정(0 나누기 방지)
        if min_val == max_val:
            max_val = min_val + 1
        min_max[metric] = (min_val, max_val)
    
    fig = go.Figure()
    
    for i, company in enumerate(companies):
        company_data = chart_df[chart_df['회사'] == company] if '회사' in chart_df.columns else chart_df
        normalized_values = []
        for metric in metrics:
            raw_value = company_data.loc[company_data['구분'] == metric, '수치'].values
            if len(raw_value) == 0:
                norm_value = 0
            else:
                val = raw_value[0]
                min_val, max_val = min_max[metric]
                norm_value = (val - min_val) / (max_val - min_val)
            normalized_values.append(norm_value)
        
        # 닫힌 도형을 위해 첫 값 반복
        normalized_values.append(normalized_values[0])
        theta_labels = list(metrics) + [metrics[0]] if len(metrics) > 0 else ['지표1']
        
        # 색상
        color = get_company_color(company, companies)
        
        # SK에너지 스타일 강조
        if 'SK' in company:
            line_width = 5
            marker_size = 12
            name_style = f"**{company}**"
        else:
            line_width = 3
            marker_size = 8
            name_style = company
        
        fig.add_trace(go.Scatterpolar(
            r=normalized_values,
            theta=theta_labels,
            fill='toself',
            name=name_style,
            line=dict(width=line_width, color=color),
            marker=dict(size=marker_size, color=color)
        ))
    
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 1],  # 정규화 했으니 0~1 범위
                tickmode='linear',
                tick0=0,
                dtick=0.2,
                tickfont=dict(size=14)
            ),
            angularaxis=dict(
                tickfont=dict(size=16)
            )
        ),
        title="🎯 SK에너지 vs 경쟁사 수익성 지표 비교 (정규화)",
        height=600,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=14)
        ),
        title_font_size=20,
        font=dict(size=14)
    )
    
    return fig


def create_quarterly_trend_chart(quarterly_df: pd.DataFrame):
    """분기별 추이 혼합 차트"""
    if not PLOTLY_AVAILABLE or quarterly_df.empty: 
        print("⚠️ Plotly/kaleido가 없거나 데이터가 비어있어 분기별 차트를 생성할 수 없습니다.")
        return None

    fig = go.Figure()
    companies = quarterly_df['회사'].unique()

    for company in companies:
        company_data = quarterly_df[quarterly_df['회사'] == company]
        color = get_company_color(company, companies)
        
        # 매출액 (Bar)
        if '매출액(조원)' in company_data.columns:
            fig.add_trace(go.Bar(
                x=company_data['분기'], y=company_data['매출액(조원)'], name=f"{company} 매출액(조)",
                marker_color=color
            ))
    
    fig.update_layout(
        barmode='group', title="📈 분기별 매출액 추이",
        xaxis_title="분기", yaxis_title="매출액 (조원)",
        font=dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif")
    )
    return fig


def create_gap_trend_chart(quarterly_df: pd.DataFrame):
    """분기별 갭 추이 차트"""
    if not PLOTLY_AVAILABLE or quarterly_df.empty: 
        print("⚠️ Plotly/kaleido가 없거나 데이터가 비어있어 갭 추이 차트를 생성할 수 없습니다.")
        return None

    fig = go.Figure()
    companies = quarterly_df['회사'].unique()

    for company in companies:
        company_data = quarterly_df[quarterly_df['회사'] == company]
        color = get_company_color(company, companies)
        
        # 영업이익률 (Line)
        if '영업이익률(%)' in company_data.columns:
            fig.add_trace(go.Scatter(
                x=company_data['분기'], y=company_data['영업이익률(%)'], 
                name=f"{company} 영업이익률(%)",
                mode='lines+markers', line=dict(color=color, width=3),
                marker=dict(size=8)
            ))
    
    fig.update_layout(
        title="📊 분기별 영업이익률 갭 추이",
        xaxis_title="분기", yaxis_title="영업이익률 (%)",
        font=dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif"),
        height=450
    )
    return fig


def create_gap_analysis(financial_df: pd.DataFrame, raw_cols: list):
    """SK에너지 대비 경쟁사 갭차이 분석"""
    if financial_df.empty or not raw_cols:
        return pd.DataFrame()
    
    # SK에너지 컬럼 찾기
    sk_col = None
    for col in raw_cols:
        if 'SK에너지' in col:
            sk_col = col
            break
    
    if not sk_col:
        return pd.DataFrame()
    
    gap_analysis = []
    
    for _, row in financial_df.iterrows():
        indicator = row['구분']
        sk_value = row.get(sk_col, 0)
        
        if sk_value == 0:
            continue
            
        gap_data = {'지표': indicator, 'SK에너지': sk_value}
        
        for col in raw_cols:
            if col != sk_col:
                company_name = col.replace('_원시값', '')
                company_value = row.get(col, 0)
                
                # 갭차이 계산 (SK에너지 대비)
                if sk_value != 0:
                    gap_percentage = ((company_value - sk_value) / abs(sk_value)) * 100
                    gap_amount = company_value - sk_value
                else:
                    gap_percentage = 0
                    gap_amount = company_value
                
                gap_data[f'{company_name}_갭(%)'] = round(gap_percentage, 2)
                gap_data[f'{company_name}_갭(금액)'] = gap_amount
                gap_data[f'{company_name}_원본값'] = company_value
        
        gap_analysis.append(gap_data)
    
    return pd.DataFrame(gap_analysis)


def create_gap_chart(gap_analysis_df: pd.DataFrame):
    """갭차이 시각화 차트"""
    if not PLOTLY_AVAILABLE or gap_analysis_df.empty:
        print("⚠️ Plotly/kaleido가 없거나 데이터가 비어있어 갭 차트를 생성할 수 없습니다.")
        return None
    
    # 갭% 컬럼만 추출
    gap_cols = [col for col in gap_analysis_df.columns if col.endswith('_갭(%)')]
    if not gap_cols:
        return None
    
    # 데이터 준비
    chart_data = []
    for _, row in gap_analysis_df.iterrows():
        indicator = row['지표']
        for col in gap_cols:
            company = col.replace('_갭(%)', '')
            gap_value = row[col]
            chart_data.append({
                '지표': indicator,
                '회사': company,
                '갭(%)': gap_value
            })
    
    chart_df = pd.DataFrame(chart_data)
    
    # 색상 매핑
    companies = chart_df['회사'].unique()
    color_map = {comp: get_company_color(comp, companies) for comp in companies}
    
    fig = px.bar(
        chart_df, x='지표', y='갭(%)', color='회사',
        title="📊 SK에너지 대비 경쟁사 갭차이 분석",
        text='갭(%)', color_discrete_map=color_map, barmode='group', height=500
    )
    
    fig.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
    fig.update_layout(
        yaxis_title="갭차이 (%)", xaxis_title="재무 지표", legend_title="회사",
        font=dict(family="Malgun Gothic, Apple SD Gothic Neo, sans-serif"),
        # 0선 추가
        shapes=[dict(
            type='line', x0=-0.5, x1=len(chart_df['지표'].unique())-0.5, y0=0, y1=0,
            line=dict(color='red', width=2, dash='dash')
        )],
        annotations=[dict(
            x=0.5, y=0, xref='paper', yref='y',
            text='SK에너지 기준선', showarrow=False,
            font=dict(color='red', size=12)
        )]
    )
    
    return fig


def save_chart_as_image(fig, filename_prefix="chart"):
    """Plotly 차트를 이미지 파일로 저장 (kaleido 사용)"""
    try:
        if not PLOTLY_AVAILABLE:
            print("❌ Plotly/kaleido가 설치되지 않아 차트를 저장할 수 없습니다.")
            print("다음 명령어로 설치하세요: pip install plotly kaleido")
            return None
            
        # 임시 파일 생성
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png', prefix=f'{filename_prefix}_')
        temp_path = temp_file.name
        temp_file.close()
        
        print(f"🔄 차트 저장 시도: {type(fig)} -> {temp_path}")
        
        # Plotly 차트를 고해상도 PNG로 저장 (kaleido 사용)
        if hasattr(fig, 'write_image'):
            try:
                fig.write_image(
                    temp_path, 
                    format='png',
                    width=1000,    # 고해상도
                    height=600, 
                    scale=2,       # 2배 확대로 선명도 증가
                    engine='kaleido'  # kaleido 엔진 명시적 지정
                )
                print(f"✅ Plotly 차트 저장 성공 (kaleido 사용)")
                
                # 파일이 실제로 생성되었는지 확인
                if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
                    print(f"✅ 차트 이미지 저장: {temp_path} ({os.path.getsize(temp_path)} bytes)")
                    return temp_path
                else:
                    print(f"❌ 차트 파일이 비어있거나 생성되지 않음")
                    return None
                    
            except Exception as e:
                print(f"⚠️ kaleido를 사용한 차트 저장 실패: {e}")
                
                # 대안: plotly.io 사용
                try:
                    import plotly.io as pio
                    img_bytes = pio.to_image(fig, format='png', width=1000, height=600, scale=2)
                    with open(temp_path, 'wb') as f:
                        f.write(img_bytes)
                    print(f"✅ plotly.io 대안 방법 성공")
                    return temp_path
                except Exception as e2:
                    print(f"⚠️ plotly.io 대안 방법도 실패: {e2}")
                    return None
        else:
            print(f"❌ 지원하지 않는 차트 타입: {type(fig)}")
            return None
            
    except Exception as e:
        print(f"❌ 차트 이미지 저장 실패: {e}")
        return None


def capture_streamlit_charts(chart_objects):
    """Streamlit 차트 객체들을 이미지 파일로 저장하고 경로 리스트 반환"""
    chart_paths = []
    
    if not chart_objects:
        print("⚠️ 차트 객체가 없습니다")
        return chart_paths
    
    if not PLOTLY_AVAILABLE:
        print("❌ Plotly/kaleido가 설치되지 않아 차트 변환을 할 수 없습니다.")
        print("다음 명령어로 설치하세요: pip install plotly kaleido")
        return chart_paths
    
    print(f"🔄 {len(chart_objects)}개 차트 처리 시작...")
    
    for i, chart in enumerate(chart_objects):
        if chart is not None:
            print(f"🔄 차트 {i+1} 처리 중: {type(chart)}")
            chart_path = save_chart_as_image(chart, f"chart_{i+1}")
            if chart_path:
                chart_paths.append(chart_path)
                print(f"✅ 차트 {i+1} 성공")
            else:
                print(f"❌ 차트 {i+1} 실패")
        else:
            print(f"⚠️ 차트 {i+1}이 None입니다")
    
    print(f"✅ 총 {len(chart_paths)}개 차트 이미지 생성 완료")
    return chart_paths


def get_font_paths():
    """스트림릿 환경에 맞는 폰트 경로를 반환"""
    font_paths = {
        "Korean": "fonts/NanumGothic.ttf",
        "KoreanBold": "fonts/NanumGothicBold.ttf", 
        "KoreanSerif": "fonts/NanumMyeongjo.ttf"
    }
    
    found_fonts = {}
    for font_name, font_path in font_paths.items():
        try:
            if os.path.exists(font_path):
                file_size = os.path.getsize(font_path)
                if file_size > 0:
                    found_fonts[font_name] = font_path
                    print(f"✅ 폰트 발견: {font_name} -> {font_path}")
                else:
                    print(f"⚠️ 폰트 파일이 비어있음: {font_path}")
            else:
                print(f"⚠️ 폰트 파일 없음: {font_path}")
        except Exception as e:
            print(f"❌ 폰트 체크 오류 ({font_name}): {e}")
    
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
                    print(f"✅ 폰트 등록 성공: {font_name}")
                registered_fonts[font_name] = font_name
            except Exception as e:
                print(f"⚠️ 폰트 등록 실패 ({font_name}): {e}")
                registered_fonts[font_name] = default_font
        else:
            registered_fonts[font_name] = default_font
            print(f"🔄 기본 폰트 사용: {font_name} -> {default_font}")
    
    return registered_fonts


def generate_strategic_recommendations(insights, financial_data=None, gpt_api_key=None):
    """AI 인사이트를 바탕으로 GPT가 SK에너지 전략 제안을 생성"""
    try:
        if not insights or not GPT_AVAILABLE or not gpt_api_key:
            return "GPT 연동이 불가능하여 전략 제안을 생성할 수 없습니다."
        
        # OpenAI API 키 설정
        openai.api_key = gpt_api_key
        
        # 재무 데이터 요약 생성
        financial_summary = ""
        if financial_data is not None and not financial_data.empty:
            financial_summary = f"""
            
현재 재무 상황:
- 분석 대상: {', '.join([col for col in financial_data.columns if col != '구분' and not str(col).endswith('_원시값')])}
- 주요 지표 개수: {len(financial_data)}개
"""
        
        # GPT 프롬프트 구성
        prompt = f"""
당신은 SK에너지의 경영 컨설턴트입니다. 다음 AI 분석 인사이트를 바탕으로 SK에너지가 취해야 할 구체적인 전략과 실행 방안을 제안해주세요.

## AI 분석 인사이트:
{insights}

{financial_summary}

## 요청사항:
1. 위 분석 결과를 바탕으로 SK에너지의 현재 상황을 진단해주세요
2. 경쟁사 대비 개선이 필요한 영역을 식별해주세요  
3. 단기(6개월), 중기(1-2년), 장기(3-5년) 전략을 구체적으로 제안해주세요
4. 각 전략의 기대효과와 실행 시 주의사항을 포함해주세요

## 답변 형식:
### 1. 현황 진단
[현재 상황 분석]

### 2. 개선 영역
[우선순위별 개선 포인트]

### 3. 전략 로드맵
#### 단기 전략 (6개월)
- [구체적 실행 방안]

#### 중기 전략 (1-2년)  
- [구체적 실행 방안]

#### 장기 전략 (3-5년)
- [구체적 실행 방안]

### 4. 기대효과 및 리스크
[각 전략의 예상 성과와 주의점]

한국어로 답변해주시고, 실무진이 바로 실행할 수 있을 정도로 구체적이고 실용적인 제안을 해주세요.
"""

        print("🔄 GPT에 전략 제안 요청 중...")
        
        response = openai.ChatCompletion.create(
            model="gpt-4",  # 또는 "gpt-3.5-turbo"
            messages=[
                {"role": "system", "content": "당신은 에너지 업계 전문 경영 컨설턴트입니다."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=2000,
            temperature=0.7
        )
        
        recommendations = response.choices[0].message.content.strip()
        print("✅ GPT 전략 제안 생성 완료")
        return recommendations
        
    except Exception as e:
        print(f"❌ GPT 전략 제안 생성 실패: {e}")
        return f"전략 제안 생성 중 오류가 발생했습니다: {str(e)}"


def clean_ai_text(raw):
    """AI 인사이트 텍스트 정리"""
    try:
        if not raw or pd.isna(raw):
            return []
        
        raw_str = str(raw).strip()
        if not raw_str:
            return []
            
        raw_str = re.sub(r'[*_#>~]', '', raw_str)
        blocks = []
        
        for line in raw_str.splitlines():
            line = line.strip()
            if not line:
                continue
            if re.match(r'^\d+(\.\d+)*\s', line) or line.startswith('###'):
                blocks.append(('title', line))
            else:
                blocks.append(('body', line))
        
        return blocks
    except Exception as e:
        print(f"❌ AI 텍스트 정리 오류: {e}")
        return []


def ascii_to_table(lines, registered_fonts, header_color='#E31E24', row_colors=None):
    """ASCII 표를 reportlab 테이블로 변환"""
    try:
        if not lines or len(lines) < 3:
            return None
        
        header = [c.strip() for c in lines[0].split('|') if c.strip()]
        if not header:
            return None
            
        data = []
        for ln in lines[2:]:
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
    except Exception as e:
        print(f"❌ ASCII 테이블 변환 오류: {e}")
        return None


def split_dataframe_for_pdf(df, max_rows_per_page=20, max_cols_per_page=8):
    """DataFrame을 PDF에 맞게 페이지별로 분할"""
    try:
        if df is None or df.empty:
            return []
            
        chunks = []
        total_rows = len(df)
        total_cols = len(df.columns)
        
        for row_start in range(0, total_rows, max_rows_per_page):
            row_end = min(row_start + max_rows_per_page, total_rows)
            row_chunk = df.iloc[row_start:row_end]
            
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
    except Exception as e:
        print(f"❌ DataFrame 분할 오류: {e}")
        return []


def safe_str_convert(value):
    """안전하게 값을 문자열로 변환"""
    try:
        if pd.isna(value):
            return ""
        return str(value)
    except:
        return ""


def add_chunked_table(story, df, title, registered_fonts, BODY_STYLE, header_color='#F2F2F2'):
    """분할된 테이블을 story에 추가"""
    try:
        if df is None or df.empty:
            story.append(Paragraph(f"{title}: 데이터가 없습니다.", BODY_STYLE))
            return
        
        print(f"🔄 테이블 추가 중: {title}")
        story.append(Paragraph(title, BODY_STYLE))
        story.append(Spacer(1, 8))
        
        chunks = split_dataframe_for_pdf(df)
        
        for i, chunk_info in enumerate(chunks):
            chunk = chunk_info['data']
            
            if len(chunks) > 1:
                row_info = f"행 {chunk_info['row_range'][0]+1}~{chunk_info['row_range'][1]+1}"
                col_info = f"열 {chunk_info['col_range'][0]+1}~{chunk_info['col_range'][1]+1}"
                story.append(Paragraph(f"[{row_info}, {col_info}]", BODY_STYLE))
            
            table_data = [chunk.columns.tolist()]
            for _, row in chunk.iterrows():
                table_data.append([safe_str_convert(val) for val in row.values])
            
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
            
            if i < len(chunks) - 1 and (i + 1) % 2 == 0:
                story.append(PageBreak())
        
        print(f"✅ 테이블 추가 완료: {title}")
    except Exception as e:
        print(f"❌ 테이블 추가 오류 ({title}): {e}")
        story.append(Paragraph(f"{title}: 테이블 생성 중 오류가 발생했습니다.", BODY_STYLE))


def add_financial_data_section(story, financial_data, quarterly_df, chart_images, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """재무분석 결과 섹션 추가 (표 + 차트 이미지)"""
    try:
        print("🔄 재무분석 섹션 추가 중...")
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        
        # 1-1. 분기별 재무지표 상세 데이터
        if quarterly_df is not None and not quarterly_df.empty:
            add_chunked_table(story, quarterly_df, "1-1. 분기별 재무지표 상세 데이터", 
                             registered_fonts, BODY_STYLE, '#E6F3FF')
        else:
            story.append(Paragraph("1-1. 분기별 재무지표 상세 데이터: 데이터가 없습니다.", BODY_STYLE))
        
        story.append(Spacer(1, 12))
        
        # 1-2. SK에너지 대비 경쟁사 갭차이 분석표
        if financial_data is not None and not financial_data.empty:
            display_cols = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
            df_display = financial_data[display_cols].copy()
            add_chunked_table(story, df_display, "1-2. SK에너지 대비 경쟁사 갭차이 분석", 
                             registered_fonts, BODY_STYLE, '#F2F2F2')
        else:
            story.append(Paragraph("1-2. SK에너지 대비 경쟁사 갭차이 분석: 데이터가 없습니다.", BODY_STYLE))
        
        # 1-3. 차트 이미지들 추가
        if chart_images and len(chart_images) > 0:
            story.append(Spacer(1, 12))
            story.append(Paragraph("1-3. 시각화 차트", BODY_STYLE))
            story.append(Spacer(1, 8))
            
            for i, chart_path in enumerate(chart_images, 1):
                if chart_path and os.path.exists(chart_path):
                    try:
                        story.append(Paragraph(f"차트 {i}", BODY_STYLE))
                        story.append(RLImage(chart_path, width=500, height=300))
                        story.append(Spacer(1, 16))
                        print(f"✅ 차트 {i} 추가 완료")
                    except Exception as e:
                        print(f"⚠️ 차트 {i} 추가 실패: {e}")
                        story.append(Paragraph(f"차트 {i}: 이미지 로드 실패", BODY_STYLE))
                else:
                    print(f"⚠️ 차트 파일이 없음: {chart_path}")
        
        story.append(Spacer(1, 18))
        print("✅ 재무분석 섹션 추가 완료")
    except Exception as e:
        print(f"❌ 재무분석 섹션 추가 오류: {e}")


def add_ai_insights_section(story, insights, registered_fonts, BODY_STYLE, header_color='#E31E24'):
    """AI 인사이트 섹션 추가"""
    try:
        print("🔄 AI 인사이트 섹션 추가 중...")
        
        if not insights:
            story.append(Paragraph("AI 인사이트가 제공되지 않았습니다.", BODY_STYLE))
            story.append(Spacer(1, 18))
            return
        
        story.append(Spacer(1, 8))
        blocks = clean_ai_text(insights)
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
        print("✅ AI 인사이트 섹션 추가 완료")
    except Exception as e:
        print(f"❌ AI 인사이트 섹션 추가 오류: {e}")


def add_strategic_recommendations_section(story, recommendations, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """전략 제안 섹션 추가"""
    try:
        print("🔄 전략 제안 섹션 추가 중...")
        
        if not recommendations or "생성할 수 없습니다" in recommendations:
            story.append(Paragraph("GPT 기반 전략 제안을 생성할 수 없습니다.", BODY_STYLE))
            story.append(Spacer(1, 18))
            return
        
        story.append(Spacer(1, 8))
        
        # 전략 제안 텍스트 처리
        blocks = clean_ai_text(recommendations)
        
        for typ, line in blocks:
            if typ == 'title':
                story.append(Paragraph(f"<b>{line}</b>", BODY_STYLE))
            else:
                story.append(Paragraph(line, BODY_STYLE))
        
        story.append(Spacer(1, 18))
        print("✅ 전략 제안 섹션 추가 완료")
    except Exception as e:
        print(f"❌ 전략 제안 섹션 추가 오류: {e}")


def add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """뉴스 하이라이트 및 종합 분석 섹션 내용 추가 (헤딩 제외)"""
    try:
        print("🔄 뉴스 섹션 내용 추가 중...")
        
        if news_data is not None and not news_data.empty:
            story.append(Paragraph("4-1. 최신 뉴스 하이라이트", BODY_STYLE))
            for i, title in enumerate(news_data["제목"].head(10), 1):
                story.append(Paragraph(f"{i}. {safe_str_convert(title)}", BODY_STYLE))
            story.append(Spacer(1, 16))
            print(f"✅ 뉴스 하이라이트 {len(news_data)}건 추가")
        else:
            story.append(Paragraph("뉴스 데이터가 제공되지 않았습니다.", BODY_STYLE))
            print("⚠️ 뉴스 데이터 없음")
            
        story.append(Spacer(1, 18))
        print("✅ 뉴스 섹션 내용 추가 완료")
    except Exception as e:
        print(f"❌ 뉴스 섹션 내용 추가 오류: {e}")


def create_excel_report(financial_data=None, news_data=None, insights=None):
    """Excel 보고서 생성"""
    try:
        print("🔄 Excel 보고서 생성 시작...")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 재무분석 시트
            if financial_data is not None and not financial_data.empty:
                financial_data.to_excel(writer, sheet_name='재무분석', index=False)
            else:
                pd.DataFrame({'메모': ['재무 데이터가 없습니다.']}).to_excel(writer, sheet_name='재무분석', index=False)
            
            # 뉴스분석 시트
            if news_data is not None and not news_data.empty:
                news_data.to_excel(writer, sheet_name='뉴스분석', index=False)
            else:
                pd.DataFrame({'메모': ['뉴스 데이터가 없습니다.']}).to_excel(writer, sheet_name='뉴스분석', index=False)
            
            # AI인사이트 시트
            if insights:
                insight_lines = str(insights).split('\n')
                insight_df = pd.DataFrame({'AI 인사이트': insight_lines})
                insight_df.to_excel(writer, sheet_name='AI인사이트', index=False)
            else:
                pd.DataFrame({'메모': ['AI 인사이트가 없습니다.']}).to_excel(writer, sheet_name='AI인사이트', index=False)
        
        output.seek(0)
        print("✅ Excel 보고서 생성 완료!")
        return output.getvalue()
        
    except Exception as e:
        print(f"❌ Excel 보고서 생성 오류: {e}")
        # 최소한의 에러 Excel 생성
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            error_df = pd.DataFrame({
                '오류': [f"Excel 생성 중 오류 발생: {str(e)}"],
                '해결방법': ['시스템 관리자에게 문의해주세요.']
            })
            error_df.to_excel(writer, sheet_name='오류정보', index=False)
        output.seek(0)
        return output.getvalue()


def create_enhanced_pdf_report(
    financial_data=None,
    news_data=None,
    insights=None,
    selected_charts=None,  # 기존 매개변수명 유지 (하위 호환성)
    quarterly_df=None,
    show_footer=False,
    report_target="SK이노베이션 경영진",
    report_author="보고자 미기재",
    gpt_api_key=None,  # GPT API 키
    chart_images=None,  # Streamlit 차트 이미지 경로들
    font_paths=None,
):
    """향상된 PDF 보고서 생성 (GPT 전략 제안 포함, kaleido 사용)"""
    
    try:
        print("🔄 PDF 보고서 생성 시작...")
        
        # kaleido 체크
        if not PLOTLY_AVAILABLE:
            print("❌ Plotly/kaleido가 설치되지 않았습니다.")
            print("다음 명령어로 설치하세요: pip install plotly kaleido")
        
        # 하위 호환성: selected_charts를 chart_images로 변환
        if selected_charts and not chart_images:
            print("🔄 selected_charts를 chart_images로 변환 중...")
            if isinstance(selected_charts, list) and len(selected_charts) > 0:
                first_item = selected_charts[0]
                if isinstance(first_item, str):
                    # 이미 이미지 경로들인 경우
                    chart_images = selected_charts
                else:
                    # Plotly 차트 객체들인 경우 이미지로 변환
                    chart_images = capture_streamlit_charts(selected_charts)
            else:
                chart_images = []
        
        # chart_images가 없으면 빈 리스트로 설정
        if not chart_images:
            chart_images = []
        
        # 폰트 등록
        registered_fonts = register_fonts_safe()
        
        # 스타일 정의
        TITLE_STYLE = ParagraphStyle(
            'Title',
            fontName=registered_fonts.get('KoreanBold', 'Helvetica-Bold'),
            fontSize=20,
            leading=30,
            spaceAfter=15,
            alignment=1,
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
        <b>보고대상:</b> {safe_str_convert(report_target)}<br/>
        <b>보고자:</b> {safe_str_convert(report_author)}
        """
        story.append(Paragraph(report_info, BODY_STYLE))
        story.append(Spacer(1, 30))

        # 1. 재무분석 결과 (표 + 차트 이미지)
        add_financial_data_section(story, financial_data, quarterly_df, chart_images, 
                                   registered_fonts, HEADING_STYLE, BODY_STYLE)
        
        # 2. AI 인사이트
        story.append(Paragraph("2. AI 분석 인사이트", HEADING_STYLE))
        add_ai_insights_section(story, insights, registered_fonts, BODY_STYLE)
        
        # 3. GPT 기반 전략 제안 (AI 인사이트가 있을 때만)
        if insights:
            print("🔄 GPT 전략 제안 생성 중...")
            strategic_recommendations = generate_strategic_recommendations(
                insights, financial_data, gpt_api_key
            )
            story.append(Paragraph("3. SK에너지 전략 제안", HEADING_STYLE))
            add_strategic_recommendations_section(story, strategic_recommendations, 
                                                registered_fonts, HEADING_STYLE, BODY_STYLE)
        else:
            print("⚠️ AI 인사이트가 없어서 GPT 전략 제안을 생성하지 않습니다")
        
        # 4. 뉴스 하이라이트 및 종합 분석
        story.append(Paragraph("4. 뉴스 하이라이트 및 종합 분석", HEADING_STYLE))
        add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE)

        # 푸터 (선택사항)
        if show_footer:
            story.append(Spacer(1, 24))
            footer_text = "※ 본 보고서는 대시보드에서 자동 생성되었습니다."
            story.append(Paragraph(footer_text, BODY_STYLE))

        # 페이지 번호 추가 함수
        def _page_number(canvas, doc):
            try:
                canvas.setFont('Helvetica', 9)
                canvas.drawCentredString(A4[0]/2, 20, f"- {canvas.getPageNumber()} -")
            except Exception as e:
                print(f"⚠️ 페이지 번호 추가 실패: {e}")

        # PDF 문서 생성
        print("🔄 PDF 문서 빌드 중...")
        doc.build(story, onFirstPage=_page_number, onLaterPages=_page_number)
        buffer.seek(0)
        
        # 차트 이미지 파일들 정리
        if chart_images:
            for chart_path in chart_images:
                try:
                    if chart_path and os.path.exists(chart_path):
                        os.unlink(chart_path)
                        print(f"✅ 임시 차트 파일 삭제: {chart_path}")
                except Exception as e:
                    print(f"⚠️ 임시 파일 삭제 실패: {e}")
        
        print("✅ PDF 보고서 생성 완료!")
        return buffer.getvalue()
        
    except Exception as e:
        print(f"❌ PDF 보고서 생성 중 오류 발생: {e}")
        import traceback
        print(f"상세 오류: {traceback.format_exc()}")
        
        # 최소한의 에러 PDF 생성 시도
        try:
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            
            error_story = [
                Paragraph("보고서 생성 오류", getSampleStyleSheet()['Title']),
                Spacer(1, 20),
                Paragraph(f"오류 내용: {str(e)}", getSampleStyleSheet()['Normal']),
                Spacer(1, 12),
                Paragraph("시스템 관리자에게 문의해주세요.", getSampleStyleSheet()['Normal'])
            ]
            
            doc.build(error_story)
            buffer.seek(0)
            return buffer.getvalue()
        except Exception as e2:
            print(f"❌ 에러 보고서 생성도 실패: {e2}")
            raise e


def generate_report_with_gpt_insights(
    financial_data=None,
    news_data=None,
    insights=None,
    streamlit_charts=None,  # Streamlit에서 생성한 차트 객체들
    quarterly_df=None,
    gpt_api_key=None,
    **kwargs
):
    """
    Streamlit 차트와 GPT 인사이트를 포함한 완전한 보고서 생성 (kaleido 사용)
    
    사용 예시:
    pdf_bytes = generate_report_with_gpt_insights(
        financial_data=df,
        insights=ai_insights,
        streamlit_charts=[fig1, fig2, fig3],  # Streamlit에서 st.plotly_chart()로 보여준 차트들
        gpt_api_key=openai_api_key
    )
    """
    try:
        print("🔄 완전한 보고서 생성 시작...")
        
        if not PLOTLY_AVAILABLE:
            print("❌ Plotly/kaleido가 설치되지 않았습니다.")
            print("차트 없이 텍스트 기반 보고서를 생성합니다.")
        
        # Streamlit 차트들을 이미지로 변환
        chart_images = []
        if streamlit_charts and PLOTLY_AVAILABLE:
            print(f"🔄 {len(streamlit_charts)}개 차트를 이미지로 변환 중...")
            chart_images = capture_streamlit_charts(streamlit_charts)
            print(f"✅ {len(chart_images)}개 차트 이미지 생성 완료")
        elif streamlit_charts and not PLOTLY_AVAILABLE:
            print("⚠️ kaleido가 없어서 차트를 변환할 수 없습니다.")
        
        # PDF 보고서 생성
        pdf_bytes = create_enhanced_pdf_report(
            financial_data=financial_data,
            news_data=news_data,
            insights=insights,
            chart_images=chart_images,
            quarterly_df=quarterly_df,
            gpt_api_key=gpt_api_key,
            **kwargs
        )
        
        return pdf_bytes
        
    except Exception as e:
        print(f"❌ 완전한 보고서 생성 실패: {e}")
        raise e


# 설치 체크 및 안내 함수
def check_dependencies():
    """필요한 패키지들이 설치되어 있는지 체크"""
    missing_packages = []
    
    try:
        import plotly
        print("✅ plotly 설치됨")
    except ImportError:
        missing_packages.append("plotly")
    
    try:
        import kaleido
        print("✅ kaleido 설치됨")
    except ImportError:
        missing_packages.append("kaleido")
    
    try:
        import reportlab
        print("✅ reportlab 설치됨")
    except ImportError:
        missing_packages.append("reportlab")
    
    try:
        import pandas
        print("✅ pandas 설치됨")
    except ImportError:
        missing_packages.append("pandas")
    
    try:
        import openpyxl
        print("✅ openpyxl 설치됨")
    except ImportError:
        missing_packages.append("openpyxl")
    
    if missing_packages:
        print(f"❌ 다음 패키지들을 설치해주세요:")
        for pkg in missing_packages:
            print(f"   pip install {pkg}")
        return False
    else:
        print("✅ 모든 필수 패키지가 설치되어 있습니다!")
        return True


# 사용 예시
if __name__ == "__main__":
    print("📦 SK에너지 보고서 생성 모듈")
    print("=" * 50)
    check_dependencies()
    
    if PLOTLY_AVAILABLE:
        print("🎯 차트 기능 사용 가능!")
    else:
        print("⚠️ 차트 기능을 사용하려면 다음을 실행하세요:")
        print("   pip install plotly kaleido")
