import io
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

def build_praise_pptx(slides_data, font_name="프리젠테이션 Bold", font_size_pt=81.0):
    """
    슬라이드 리스트를 바탕으로 16:9 와이드스크린, 검정 배경, 양 끝 꽉 찬 텍스트 박스의
    표준 PPTX 프레젠테이션을 생성합니다. (의도치 않은 줄바꿈 원천 차단)
    """
    prs = Presentation()
    
    # 16:9 와이드스크린 (20.0 x 11.25 inches)
    prs.slide_width = Inches(20.0)
    prs.slide_height = Inches(11.25)
    
    blank_layout = prs.slide_layouts[6]
    
    for slide_info in slides_data:
        slide = prs.slides.add_slide(blank_layout)
        
        # 1. 완전 검은색 배경 (#000000)
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(0, 0, 0)
        
        text_content = slide_info.get("text", "").strip()
        
        # 빈 슬라이드(암전)인 경우 텍스트 박스 생략
        if slide_info.get("type") == "blank" or not text_content:
            continue
            
        # 2. 텍스트 박스: 가로 양 끝(0 ~ 20 inch) 100% 꽉 채움 + 좌우 여백 0
        left = Inches(0.0)
        width = Inches(20.0)
        
        lines = [l.strip() for l in text_content.split('\n') if l.strip()]
        
        # 세로 중앙 균형 배치
        if len(lines) == 1:
            top = Inches(3.6)
            height = Inches(4.0)
        elif len(lines) == 2:
            top = Inches(2.6)
            height = Inches(6.0)
        else:
            top = Inches(1.8)
            height = Inches(7.5)
            
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.0)
        tf.margin_right = Inches(0.0)
        tf.margin_top = Inches(0.0)
        tf.margin_bottom = Inches(0.0)
        
        for idx, line_str in enumerate(lines):
            if idx == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
                p.space_before = Pt(20) # 줄 간격 20pt
                
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = line_str
            run.font.name = font_name
            run.font.size = Pt(font_size_pt)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255) # 흰색
            
    return prs

def get_pptx_bytes(prs):
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output
