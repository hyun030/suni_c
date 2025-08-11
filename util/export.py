    # -*- coding: utf-8 -*-
import io
import os
import re
import tempfile
import pandas as pd
from datetime import datetime

# 외부 라이브러리
# 설치: pip install reportlab plotly xlsxwriter
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image as RLImage
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

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
    PDF 보고서 생성 (한글 포함)
    ※ font_paths 내 경로는 본인 PC 절대 경로로 수정 필수
    """

    if not PDF_AVAILABLE:
        raise ImportError("reportlab 라이브러리가 설치되어 있지 않습니다.")

    # 1. 폰트 경로 설정 (본인 환경에 맞게 절대경로 수정하세요)
    font_paths = {
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
    for fam, path in font_paths.items():
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(fam, path))
                # 등록 성공 시 이름 그대로 유지
            except Exception as e:
                print(f"[폰트 등록 실패] {fam} - {e}")
                font_paths[fam] = fallback_fonts[fam]
        else:
            print(f"[폰트 파일 없음] {path}")
            font_paths[fam] = fallback_fonts[fam]

    # 3. 스타일 정의
    styles = getSampleStyleSheet()
    TITLE_STYLE = ParagraphStyle(
        'TITLE',
        fontName=font_paths["KoreanBold"] if font_paths["KoreanBold"] in pdfmetrics.getRegisteredFontNames() else fallback_fonts["KoreanBold"],
        fontSize=20,
        leading=34,
        spaceAfter=18
    )
    HEADING_STYLE = ParagraphStyle(
        'HEADING',
        fontName=font_paths["KoreanBold"] if font_paths["KoreanBold"] in pdfmetrics.getRegisteredFontNames() else fallback_fonts["KoreanBold"],
        fontSize=14,
        leading=23.8,
        textColor=colors.HexColor('#E31E24'),
        spaceBefore=16,
        spaceAfter=10
    )
    BODY_STYLE = ParagraphStyle(
        'BODY',
        fontName=font_paths["KoreanSerif"] if font_paths["KoreanSerif"] in pdfmetrics.getRegisteredFontNames() else fallback_fonts["KoreanSerif"],
        fontSize=12,
        leading=20.4,
        spaceAfter=6
    )

    # 내부 함수: AI 텍스트 정리
    def _clean_ai_text(raw: str) -> list[tuple[str, str]]:
        raw = re.sub(r'[*_#>~`]', '', raw)
        blocks = []
        for ln in raw.splitlines():
            ln = ln.strip()
            if not ln:
                continue
            if re.match(r'^\d+(\.\d+)*\s', ln):
                blocks.append(('title', ln))
            else:
                blocks.append(('body', ln))
        return blocks

    # 내부 함수: ASCII 테이블 변환
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
            ('FONTNAME', (0, 0), (-1, 0), font_paths["KoreanBold"]),
            ('FONTNAME', (0, 1), (-1, -1), font_paths["KoreanSerif"]),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1),
             [colors.whitesmoke, colors.HexColor('#F7F7F7')]),
        ]))
        return tbl

    # 내부 함수: plotly figure → PNG bytes 변환
    def _fig_to_png_bytes(fig, width=900, height=450):
        try:
            return fig.to_image(format="png", width=width, height=height)
        except Exception:
            return None

    # 4. PDF 문서 생성 준비
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=54, rightMargin=54, topMargin=54, bottomMargin=54)

    story = []

    # 표지
    story.append(Paragraph("손익개선을 위한 SK에너지 및 경쟁사 비교 분석 보고서", TITLE_STYLE))
    story.append(Paragraph(
        f"보고일자: {datetime.now().strftime('%Y년 %m월 %d일')}    보고대상: {report_target}    보고자: {report_author}",
        BODY_STYLE))
    story.append(Spacer(1, 12))

    # 5. 재무분석 결과 표 출력
    if financial_data is not None and hasattr(financial_data, "empty") and not financial_data.empty:
        story.append(Paragraph("1. 재무분석 결과", HEADING_STYLE))
        df_disp = financial_data[[c for c in financial_data.columns if not str(c).endswith('_원시값')]].copy()
        tbl = Table([df_disp.columns.tolist()] + df_disp.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F2F2F2')),
            ('FONTNAME', (0, 0), (-1, 0), font_paths["KoreanBold"]),
            ('FONTNAME', (0, 1), (-1, -1), font_paths["KoreanSerif"]),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 18))

    # 6. plotly 차트 삽입
    charts_added = False
    if PLOTLY_AVAILABLE:
        # selected_charts 우선 렌더링
        if selected_charts:
            story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
            for idx, fig in enumerate(selected_charts, start=1):
                img_bytes = _fig_to_png_bytes(fig)
                if img_bytes:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    story.append(Paragraph(f"2-{idx}. 추가 차트", BODY_STYLE))
                    story.append(RLImage(tmp_path, width=500, height=280))
                    story.append(Spacer(1, 16))
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass
            charts_added = True

        # quarterly_df 시각화 (분기별 영업이익률, 매출액 등)
        if quarterly_df is not None and hasattr(quarterly_df, "empty") and not quarterly_df.empty:
            # 영업이익률 추이 꺾은선
            if all(col in quarterly_df.columns for col in ['분기', '회사', '영업이익률']):
                fig_line = go.Figure()
                for comp in quarterly_df['회사'].dropna().unique():
                    cdf = quarterly_df[quarterly_df['회사'] == comp]
                    fig_line.add_trace(go.Scatter(x=cdf['분기'], y=cdf['영업이익률'], mode='lines+markers', name=f"{comp}"))
                fig_line.update_layout(title="분기별 영업이익률 추이", xaxis_title="분기", yaxis_title="영업이익률(%)")
                img_bytes = _fig_to_png_bytes(fig_line)
                if img_bytes:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    if not charts_added:
                        story.append(Paragraph("2. 시각화 차트", HEADING_STYLE))
                    story.append(Paragraph("2-1. 분기별 영업이익률 추이 (꺾은선)", BODY_STYLE))
                    story.append(RLImage(tmp_path, width=500, height=280))
                    story.append(Spacer(1, 16))
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                    charts_added = True

            # 매출액 추이 꺾은선
            if all(col in quarterly_df.columns for col in ['분기', '회사', '매출액']):
                fig_rev = go.Figure()
                for comp in quarterly_df['회사'].dropna().unique():
                    cdf = quarterly_df[quarterly_df['회사'] == comp]
                    fig_rev.add_trace(go.Scatter(x=cdf['분기'], y=cdf['매출액'], mode='lines+markers', name=f"{comp}"))
                fig_rev.update_layout(title="분기별 매출액 추이", xaxis_title="분기", yaxis_title="매출액(조원)")
                img_bytes = _fig_to_png_bytes(fig_rev)
                if img_bytes:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                        tmp.write(img_bytes)
                        tmp_path = tmp.name
                    story.append(Paragraph("2-2. 분기별 매출액 추이 (꺾은선)", BODY_STYLE))
                    story.append(RLImage(tmp_path, width=500, height=280))
                    story.append(Spacer(1, 16))
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass
                    charts_added = True

    # 7. 최신 뉴스 하이라이트
    if news_data is not None and hasattr(news_data, "empty") and not news_data.empty:
        story.append(Paragraph("3. 최신 뉴스 하이라이트", HEADING_STYLE))
        for i, title in enumerate(news_data["제목"].head(5), 1):
            story.append(Paragraph(f"{i}. {title}", BODY_STYLE))
        story.append(Spacer(1, 12))

    # 8. AI 인사이트 영역
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

    # 9. 푸터
    if show_footer:
        story.append(Spacer(1, 24))
        story.append(Paragraph("※ 본 보고서는 대시보드에서 자동 생성되었습니다.", BODY_STYLE))

    # PDF 빌드
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
