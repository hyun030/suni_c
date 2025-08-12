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
