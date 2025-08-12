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
    print("✅ Plotly 라이브러리 로드 성공")
except ImportError as e:
    PLOTLY_AVAILABLE = False
    print(f"⚠️ Plotly 라이브러리 로드 실패: {e}")


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
            print(f"🔄 기본 폰트 사용: {font_name} -> {default_font}")
    
    return registered_fonts


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
            if re.match(r'^\d+(\.\d+)*\s', line):
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
        if not lines or len(lines) < 3:  # 헤더 + 구분선 + 최소 1개 데이터
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
    except Exception as e:
        print(f"❌ ASCII 테이블 변환 오류: {e}")
        return None


def fig_to_png_bytes(fig, width=900, height=450):
    """Plotly 차트를 PNG 바이트로 변환"""
    try:
        if not PLOTLY_AVAILABLE:
            print("❌ Plotly 라이브러리가 없어서 차트 변환 불가능")
            return None
            
        if fig is None:
            print("❌ 차트 객체가 None입니다")
            return None
            
        print(f"🔄 차트를 PNG로 변환 중... (크기: {width}x{height})")
        
        # 첫 번째 시도: 기본 방법
        try:
            img_bytes = fig.to_image(format="png", width=width, height=height)
            print("✅ 기본 엔진으로 차트 변환 성공")
            return img_bytes
        except Exception as e1:
            print(f"⚠️ 기본 엔진 실패: {e1}")
            
            # 두 번째 시도: kaleido 엔진
            try:
                img_bytes = fig.to_image(format="png", width=width, height=height, engine="kaleido")
                print("✅ kaleido 엔진으로 차트 변환 성공")
                return img_bytes
            except Exception as e2:
                print(f"⚠️ kaleido 엔진 실패: {e2}")
                
                # 세 번째 시도: orca 엔진
                try:
                    img_bytes = fig.to_image(format="png", width=width, height=height, engine="orca")
                    print("✅ orca 엔진으로 차트 변환 성공")
                    return img_bytes
                except Exception as e3:
                    print(f"❌ 모든 차트 변환 엔진 실패")
                    print(f"   - 기본: {e1}")
                    print(f"   - kaleido: {e2}")
                    print(f"   - orca: {e3}")
                    return None
                    
    except Exception as e:
        print(f"❌ 차트 변환 중 예상치 못한 오류: {e}")
        return None


def split_dataframe_for_pdf(df, max_rows_per_page=20, max_cols_per_page=8):
    """DataFrame을 PDF에 맞게 페이지별로 분할"""
    try:
        if df is None or df.empty:
            return []
            
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
        
        print(f"✅ DataFrame 분할 완료: {len(chunks)}개 청크")
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
            print(f"⚠️ 테이블 데이터 없음: {title}")
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
            
            # 테이블 데이터 준비
            table_data = [chunk.columns.tolist()]
            for _, row in chunk.iterrows():
                table_data.append([safe_str_convert(val) for val in row.values])
            
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
        
        print(f"✅ 테이블 추가 완료: {title}")
    except Exception as e:
        print(f"❌ 테이블 추가 오류 ({title}): {e}")
        story.append(Paragraph(f"{title}: 테이블 생성 중 오류가 발생했습니다.", BODY_STYLE))


def add_financial_data_section(story, financial_data, quarterly_df, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """재무분석 결과 섹션 추가"""
    try:
        print("🔄 재무분석 섹션 추가 중...")
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        
        # 1-1. 분기별 재무지표 상세 데이터 (순서 변경)
        if quarterly_df is not None and not quarterly_df.empty:
            add_chunked_table(story, quarterly_df, "1-1. 분기별 재무지표 상세 데이터", 
                             registered_fonts, BODY_STYLE, '#E6F3FF')
            print("✅ 1-1. 분기별 재무지표 데이터 추가 완료")
        else:
            story.append(Paragraph("1-1. 분기별 재무지표 상세 데이터: 데이터가 없습니다.", BODY_STYLE))
            print("⚠️ 1-1. 분기별 재무지표 데이터 없음")
        
        story.append(Spacer(1, 12))
        
        # 1-2. SK에너지 대비 경쟁사 갭차이 분석표 (순서 변경)
        if financial_data is not None and not financial_data.empty:
            # 원시값 컬럼 제외
            display_cols = [c for c in financial_data.columns if not str(c).endswith('_원시값')]
            df_display = financial_data[display_cols].copy()
            add_chunked_table(story, df_display, "1-2. SK에너지 대비 경쟁사 갭차이 분석", 
                             registered_fonts, BODY_STYLE, '#F2F2F2')
            print("✅ 1-2. SK에너지 대비 경쟁사 갭차이 분석 추가 완료")
        else:
            story.append(Paragraph("1-2. SK에너지 대비 경쟁사 갭차이 분석: 데이터가 없습니다.", BODY_STYLE))
            print("⚠️ 1-2. SK에너지 대비 경쟁사 갭차이 분석 데이터 없음")
        
        story.append(Spacer(1, 18))
        print("✅ 재무분석 섹션 추가 완료")
    except Exception as e:
        print(f"❌ 재무분석 섹션 추가 오류: {e}")


def add_charts_section(story, financial_data, quarterly_df, selected_charts, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """시각화 차트 섹션 추가"""
    try:
        print("🔄 차트 섹션 추가 중...")
        story.append(Paragraph("2. 시각화 차트 및 분석", HEADING_STYLE))
        
        charts_added = False
        
        if not PLOTLY_AVAILABLE:
            story.append(Paragraph("차트 생성 라이브러리를 사용할 수 없습니다.", BODY_STYLE))
            print("⚠️ Plotly 사용 불가능으로 차트 생성 스킵")
            return charts_added
        
        # 주요 비율 비교 막대 그래프
        try:
            if (financial_data is not None and not financial_data.empty and 
                '구분' in financial_data.columns):
                
                print("🔄 비율 비교 차트 생성 중...")
                print(f"   - financial_data shape: {financial_data.shape}")
                print(f"   - financial_data columns: {list(financial_data.columns)}")
                
                ratio_rows = financial_data[financial_data['구분'].astype(str).str.contains('%', na=False)].copy()
                print(f"   - 비율 행 개수: {len(ratio_rows)}")
                
                if not ratio_rows.empty:
                    # 주요 지표 순서 정렬
                    key_order = ['영업이익률(%)', '순이익률(%)', '매출총이익률(%)', '매출원가율(%)', '판관비율(%)']
                    ratio_rows['__order__'] = ratio_rows['구분'].apply(lambda x: key_order.index(x) if x in key_order else 999)
                    ratio_rows = ratio_rows.sort_values('__order__').drop(columns='__order__')
                    print(f"   - 정렬된 비율 지표: {list(ratio_rows['구분'])}")

                    # 데이터 변환
                    melt = []
                    company_cols = [c for c in ratio_rows.columns if c != '구분' and not str(c).endswith('_원시값')]
                    print(f"   - 회사 컬럼: {company_cols}")
                    
                    for _, r in ratio_rows.iterrows():
                        for comp in company_cols:
                            val = str(r[comp]).replace('%','').strip()
                            try:
                                val_float = float(val)
                                melt.append({'지표': r['구분'], '회사': comp, '수치': val_float})
                            except:
                                print(f"   - 숫자 변환 실패: {comp}={val}")
                                continue
                    
                    print(f"   - 변환된 데이터 행 수: {len(melt)}")
                    
                    if melt:
                        bar_df = pd.DataFrame(melt)
                        print(f"   - 차트 데이터프레임 생성: {bar_df.shape}")
                        
                        fig_bar = px.bar(bar_df, x='지표', y='수치', color='회사', barmode='group', 
                                       title="주요 비율 비교")
                        print("   - Plotly 차트 객체 생성 완료")
                        
                        img_bytes = fig_to_png_bytes(fig_bar)
                        if img_bytes:
                            print(f"   - 차트 이미지 변환 성공: {len(img_bytes)} bytes")
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                                tmp.write(img_bytes)
                                tmp_path = tmp.name
                                print(f"   - 임시 파일 생성: {tmp_path}")
                            
                            story.append(Paragraph("2-1. 주요 비율 비교 (막대그래프)", BODY_STYLE))
                            story.append(RLImage(tmp_path, width=500, height=280))
                            story.append(Spacer(1, 16))
                            print("   - 차트를 PDF에 추가 완료")
                            
                            try:
                                os.unlink(tmp_path)
                                print("   - 임시 파일 정리 완료")
                            except:
                                print("   - 임시 파일 정리 실패 (무시)")
                                pass
                                
                            charts_added = True
                            print("✅ 비율 비교 차트 생성 완료")
                        else:
                            print("❌ 차트 이미지 변환 실패")
                    else:
                        print("⚠️ 변환 가능한 데이터가 없음")
                else:
                    print("⚠️ 비율 데이터 행이 없음")
            else:
                print("⚠️ 재무 데이터가 없거나 '구분' 컬럼이 없음")
                if financial_data is not None:
                    print(f"   - 데이터 존재: True, 컬럼: {list(financial_data.columns)}")
                else:
                    print("   - 데이터 존재: False")
        except Exception as e:
            print(f"❌ 비율 비교 차트 생성 실패: {e}")
            import traceback
            print(f"상세 오류: {traceback.format_exc()}")

        # 분기별 추이 그래프들
        try:
            if quarterly_df is not None and not quarterly_df.empty:
                # 영업이익률 추이
                if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
                    print("🔄 영업이익률 추이 차트 생성 중...")
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
                        print("✅ 영업이익률 추이 차트 생성 완료")

                # 매출액 추이
                if all(col in quarterly_df.columns for col in ['분기', '회사', '매출액']):
                    print("🔄 매출액 추이 차트 생성 중...")
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
                        print("✅ 매출액 추이 차트 생성 완료")
        except Exception as e:
            print(f"⚠️ 분기별 추이 차트 생성 실패: {e}")

        # 외부에서 전달된 차트들
        try:
            if selected_charts:
                print(f"🔄 추가 차트 {len(selected_charts)}개 처리 중...")
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
                print("✅ 추가 차트 처리 완료")
        except Exception as e:
            print(f"⚠️ 추가 차트 처리 실패: {e}")
        
        if not charts_added:
            story.append(Paragraph("생성 가능한 차트가 없습니다.", BODY_STYLE))
            story.append(Spacer(1, 18))
            print("⚠️ 생성된 차트 없음")
        
        print("✅ 차트 섹션 추가 완료")
        return charts_added
    except Exception as e:
        print(f"❌ 차트 섹션 추가 오류: {e}")
        return False


def add_ai_insights_section(story, insights, registered_fonts, BODY_STYLE, header_color='#E31E24'):
    """AI 인사이트 섹션 추가"""
    try:
        print("🔄 AI 인사이트 섹션 추가 중...")
        
        if not insights:
            story.append(Paragraph("2-AI. 분석 인사이트", BODY_STYLE))
            story.append(Paragraph("AI 인사이트가 제공되지 않았습니다.", BODY_STYLE))
            story.append(Spacer(1, 18))
            print("⚠️ AI 인사이트 없음")
            return
        
        story.append(Paragraph("2-AI. 분석 인사이트", BODY_STYLE))
        story.append(Spacer(1, 8))

        # AI 인사이트 텍스트 처리
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


def add_news_section(story, news_data, insights, registered_fonts, HEADING_STYLE, BODY_STYLE):
    """뉴스 하이라이트 및 종합 분석 섹션 추가"""
    try:
        print("🔄 뉴스 섹션 추가 중...")
        story.append(Paragraph("3. 뉴스 하이라이트 및 종합 분석", HEADING_STYLE))
        
        if news_data is not None and not news_data.empty:
            story.append(Paragraph("3-1. 최신 뉴스 하이라이트", BODY_STYLE))
            for i, title in enumerate(news_data["제목"].head(10), 1):
                story.append(Paragraph(f"{i}. {safe_str_convert(title)}", BODY_STYLE))
            story.append(Spacer(1, 16))
            print(f"✅ 뉴스 하이라이트 {len(news_data)}건 추가")
            
            if insights:
                story.append(Paragraph("3-2. AI 종합 분석 및 시사점", BODY_STYLE))
                story.append(Spacer(1, 8))
                
                blocks = clean_ai_text(insights)
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
                print("✅ AI 종합 분석 추가 완료")
            else:
                story.append(Paragraph("AI 종합 분석이 제공되지 않았습니다.", BODY_STYLE))
                print("⚠️ AI 종합 분석 없음")
        else:
            story.append(Paragraph("뉴스 데이터가 제공되지 않았습니다.", BODY_STYLE))
            print("⚠️ 뉴스 데이터 없음")
            
            if insights:
                story.append(Paragraph("3-1. 종합 분석 및 시사점", BODY_STYLE))
                story.append(Spacer(1, 8))
                
                blocks = clean_ai_text(insights)
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
                print("✅ 종합 분석 추가 완료")
            else:
                story.append(Paragraph("AI 인사이트도 제공되지 않았습니다.", BODY_STYLE))
                print("⚠️ AI 인사이트도 없음")
        
        story.append(Spacer(1, 18))
        print("✅ 뉴스 섹션 추가 완료")
    except Exception as e:
        print(f"❌ 뉴스 섹션 추가 오류: {e}")


def create_excel_report(financial_data=None, news_data=None, insights=None):
    """Excel 보고서 생성"""
    try:
        print("🔄 Excel 보고서 생성 시작...")
        print(f"   - 재무데이터: {'있음' if financial_data is not None and not financial_data.empty else '없음'}")
        print(f"   - 뉴스데이터: {'있음' if news_data is not None and not news_data.empty else '없음'}")
        print(f"   - AI인사이트: {'있음' if insights else '없음'}")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 재무분석 시트
            if financial_data is not None and not financial_data.empty:
                print("✅ 재무분석 시트 추가")
                financial_data.to_excel(writer, sheet_name='재무분석', index=False)
            else:
                # 빈 시트라도 생성
                pd.DataFrame({'메모': ['재무 데이터가 없습니다.']}).to_excel(writer, sheet_name='재무분석', index=False)
                print("⚠️ 재무분석 시트 - 데이터 없음")
            
            # 뉴스분석 시트
            if news_data is not None and not news_data.empty:
                print("✅ 뉴스분석 시트 추가")
                news_data.to_excel(writer, sheet_name='뉴스분석', index=False)
            else:
                pd.DataFrame({'메모': ['뉴스 데이터가 없습니다.']}).to_excel(writer, sheet_name='뉴스분석', index=False)
                print("⚠️ 뉴스분석 시트 - 데이터 없음")
            
            # AI인사이트 시트
            if insights:
                print("✅ AI인사이트 시트 추가")
                # 인사이트를 적절히 포맷팅
                insight_lines = str(insights).split('\n')
                insight_df = pd.DataFrame({'AI 인사이트': insight_lines})
                insight_df.to_excel(writer, sheet_name='AI인사이트', index=False)
            else:
                pd.DataFrame({'메모': ['AI 인사이트가 없습니다.']}).to_excel(writer, sheet_name='AI인사이트', index=False)
                print("⚠️ AI인사이트 시트 - 데이터 없음")
        
        output.seek(0)
        print("✅ Excel 보고서 생성 완료!")
        return output.getvalue()
        
    except Exception as e:
        print(f"❌ Excel 보고서 생성 오류: {e}")
        import traceback
        print(f"상세 오류: {traceback.format_exc()}")
        
        # 최소한의 에러 Excel 생성
        try:
            print("🔄 최소한의 에러 Excel 생성 시도...")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                error_df = pd.DataFrame({
                    '오류': [f"Excel 생성 중 오류 발생: {str(e)}"],
                    '해결방법': ['시스템 관리자에게 문의해주세요.']
                })
                error_df.to_excel(writer, sheet_name='오류정보', index=False)
            output.seek(0)
            print("✅ 에러 Excel 생성 완료")
            return output.getvalue()
        except Exception as e2:
            print(f"❌ 에러 Excel 생성도 실패: {e2}")
            raise e


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
    
    try:
        print("🔄 PDF 보고서 생성 시작...")
        print(f"   - 재무데이터: {'있음' if financial_data is not None and not financial_data.empty else '없음'}")
        print(f"   - 뉴스데이터: {'있음' if news_data is not None and not news_data.empty else '없음'}")
        print(f"   - AI인사이트: {'있음' if insights else '없음'}")
        print(f"   - 추가차트: {len(selected_charts) if selected_charts else 0}개")
        print(f"   - 분기별데이터: {'있음' if quarterly_df is not None and not quarterly_df.empty else '없음'}")
        
        # 조용히 폰트 등록
        registered_fonts = register_fonts_safe()
        print(f"✅ 폰트 등록 완료: {registered_fonts}")
        
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
        <b>보고대상:</b> {safe_str_convert(report_target)}<br/>
        <b>보고자:</b> {safe_str_convert(report_author)}
        """
        story.append(Paragraph(report_info, BODY_STYLE))
        story.append(Spacer(1, 30))
        print("✅ 보고서 표지 생성 완료")

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
            print("✅ 푸터 추가 완료")

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
        
        print("✅ PDF 보고서 생성 완료!")
        return buffer.getvalue()
        
    except Exception as e:
        print(f"❌ PDF 보고서 생성 중 오류 발생: {e}")
        import traceback
        print(f"상세 오류: {traceback.format_exc()}")
        
        # 최소한의 에러 PDF 생성 시도
        try:
            print("🔄 최소한의 에러 보고서 생성 시도...")
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
            print("✅ 에러 보고서 생성 완료")
            return buffer.getvalue()
        except Exception as e2:
            print(f"❌ 에러 보고서 생성도 실패: {e2}")
            raise e  # 원본 에러를 다시 발생시킴
