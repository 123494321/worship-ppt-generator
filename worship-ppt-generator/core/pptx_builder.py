import io
import os
import copy
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRESET_DIR = os.path.join(BASE_DIR, "assets", "preset_ppts")

def get_preset_ppts():
    """
    assets/preset_ppts 디렉토리에 등록된 교회 표준 PPT 파일 목록을 반환합니다.
    """
    if not os.path.exists(PRESET_DIR):
        os.makedirs(PRESET_DIR, exist_ok=True)
    files = [f for f in os.listdir(PRESET_DIR) if f.endswith('.pptx')]
    return sorted(files)

def clean_font_family(name):
    cleaned = name.replace(" (Canva 전용)", "").replace(" (윈도우 기본/호환 100%)", "").replace(" (프리텐다드)", "").strip()
    if "프리젠테이션" in cleaned:
        return "프리젠테이션 Bold"
    if cleaned.endswith(" Bold"):
        cleaned = cleaned[:-5].strip()
    return cleaned

def clean_proto_shape(grp_element):
    """
    그룹 셰이프 내부의 불필요한 투명 더미 이미지(Freeform/blip r:embed)를 제거하여
    파워포인트 오픈 시 깨짐이나 빨간색 x 표시를 원천 차단합니다.
    """
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    for sp in list(grp_element.findall('p:sp', ns)):
        cNvPr = sp.find('p:nvSpPr/p:cNvPr', ns)
        if cNvPr is not None and 'Freeform' in cNvPr.get('name', ''):
            grp_element.remove(sp)
    return grp_element

def extract_prototypes_from_prs(prs):
    """
    템플릿 프레젠테이션의 슬라이드를 순회하며 [제목], [1줄 가사], [2줄 가사]
    원본 그룹 셰이프 및 텍스트 박스 좌표/서식을 프로토타입으로 추출합니다.
    """
    proto_title = None
    proto_1line = None
    proto_2line = None

    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }

    for slide in prs.slides:
        if len(slide.shapes) == 0:
            continue
        shp = slide.shapes[0]
        text = ""
        for sp in shp._element.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}sp'):
            txBody = sp.find('p:txBody', ns)
            if txBody is not None:
                p_nodes = txBody.findall('a:p', ns)
                text = '\n'.join([
                    ''.join([r.find('a:t', ns).text for r in p.findall('a:r', ns) if r.find('a:t', ns) is not None])
                    for p in p_nodes
                ])
                break
        lines = [l.strip() for l in text.split('\n') if l.strip()]

        if len(lines) == 1:
            if any(k in lines[0] for k in ['LOGOS', '청년', '모임', '기도', '[']):
                if proto_title is None:
                    proto_title = clean_proto_shape(copy.deepcopy(shp._element))
            else:
                if proto_1line is None:
                    proto_1line = clean_proto_shape(copy.deepcopy(shp._element))
        elif len(lines) >= 2:
            if proto_2line is None:
                proto_2line = clean_proto_shape(copy.deepcopy(shp._element))

        if (proto_title is not None) and (proto_1line is not None) and (proto_2line is not None):
            break

    return proto_title, proto_1line, proto_2line

def build_praise_pptx(slides_data, font_name=None, font_size_pt=None, template_source=None):
    """
    사용자가 지정한 기준 베이스 템플릿(또는 표준 프리셋)의
    원본 좌표, 여백, 배율, 폰트 종류, 글자 크기, 폰트 임베딩을 100% 온전히 딥클론하여
    가사 줄 수에 맞게 슬라이드를 증감/생성하는 스마트 클론 PPTX 빌더입니다.
    """
    if template_source is not None:
        prs = Presentation(template_source)
    else:
        preset_files = get_preset_ppts()
        if preset_files:
            default_preset = os.path.join(PRESET_DIR, preset_files[0])
            prs = Presentation(default_preset)
        else:
            prs = Presentation()
            prs.slide_width = Inches(20.0)
            prs.slide_height = Inches(11.25)

    proto_title, proto_1line, proto_2line = extract_prototypes_from_prs(prs)

    # 템플릿 슬라이드 초기화 (내용만 비움)
    sldIdLst = prs.slides._sldIdLst
    for i in range(len(sldIdLst) - 1, -1, -1):
        prs.part.drop_rel(sldIdLst[i].rId)
        del sldIdLst[i]

    blank_layout = prs.slide_layouts[6]
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }

    for slide_info in slides_data:
        slide = prs.slides.add_slide(blank_layout)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        text_content = slide_info.get("text", "").strip()
        # 암전 슬라이드인 경우 빈 슬라이드로 유지
        if slide_info.get("type") == "blank" or not text_content:
            continue

        lines = [l.strip() for l in text_content.split('\n') if l.strip()]

        # 프로토타입 매칭
        if slide_info.get("type") == "title" or (len(lines) == 1 and any(k in text_content for k in ['LOGOS', '모임', '기도'])):
            proto = proto_title if proto_title is not None else (proto_1line if proto_1line is not None else proto_2line)
        elif len(lines) == 1:
            proto = proto_1line if proto_1line is not None else (proto_title if proto_title is not None else proto_2line)
        else:
            proto = proto_2line if proto_2line is not None else (proto_1line if proto_1line is not None else proto_title)

        if proto is not None:
            cloned_elem = copy.deepcopy(proto)
            slide.shapes._spTree.append(cloned_elem)

            for sp in cloned_elem.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}sp'):
                cNvSpPr = sp.find('p:nvSpPr/p:cNvSpPr', ns)
                if cNvSpPr is not None and cNvSpPr.get('txBox') == 'true':
                    txBody = sp.find('p:txBody', ns)
                    if txBody is not None:
                        p_nodes = txBody.findall('a:p', ns)
                        for idx, line_str in enumerate(lines):
                            if idx < len(p_nodes):
                                p_curr = p_nodes[idx]
                            else:
                                p_curr = copy.deepcopy(p_nodes[-1])
                                txBody.append(p_curr)

                            runs = p_curr.findall('a:r', ns)
                            if runs:
                                t = runs[0].find('a:t', ns)
                                if t is not None:
                                    t.text = line_str
                                for extra in runs[1:]:
                                    p_curr.remove(extra)
                                rPr = runs[0].find('a:rPr', ns)
                                if rPr is not None:
                                    if font_size_pt is not None:
                                        xml_sz = str(int(font_size_pt / 0.75 * 100))
                                        rPr.set('sz', xml_sz)
                                    if font_name is not None:
                                        family = clean_font_family(font_name)
                                        for tag in ['latin', 'ea', 'cs', 'sym']:
                                            el = rPr.find(f'a:{tag}', ns)
                                            if el is not None:
                                                el.set('typeface', family)

                        # 여분의 문단 제거
                        if len(p_nodes) > len(lines):
                            for extra_p in p_nodes[len(lines):]:
                                txBody.remove(extra_p)
                    break
        else:
            # 폴백: 표준 텍스트 박스
            txBox = slide.shapes.add_textbox(Inches(0.0), Inches(0.0), prs.slide_width, prs.slide_height)
            tf = txBox.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            fallback_family = clean_font_family(font_name) if font_name else "Pretendard"
            fallback_pt = font_size_pt if font_size_pt is not None else 60.0
            for idx, line_str in enumerate(lines):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                run.text = line_str
                run.font.name = fallback_family
                run.font.size = Pt(fallback_pt)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

    return prs

def get_pptx_bytes(prs):
    """
    Presentation 객체를 바이트 스트림으로 반환합니다.
    """
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output
