import re
from core.routine_parser import parse_routine_tokens, get_base_part_key, get_major_part_key, normalize_string

DEFAULT_CONTI_TEMPLATE = ""

def parse_user_conti(text):
    if not text or not text.strip():
        return {
            "date_str": "미입력",
            "title": "미입력",
            "prayer_person": "미입력",
            "songs": []
        }
        
    result = {
        "date_str": "미입력",
        "title": "미입력",
        "prayer_person": "미입력",
        "songs": []
    }
    
    lines = text.split('\n')
    current_song = None
    current_part = None
    part_buffer_lines = []
    
    def flush_part_buffer():
        nonlocal part_buffer_lines
        if current_song and current_part:
            raw_part_text = '\n'.join(part_buffer_lines).strip()
            if raw_part_text:
                chunks = re.split(r'\n\s*\n+', raw_part_text)
                slides = []
                for chk in chunks:
                    c = chk.strip()
                    if c:
                        slides.append(c)
                norm_part = normalize_string(current_part).upper()
                if norm_part not in current_song["parts"]:
                    current_song["parts"][norm_part] = []
                current_song["parts"][norm_part].extend(slides)
        part_buffer_lines = []

    for line in lines:
        stripped = line.strip()
        
        if stripped.startswith('# '):
            header = stripped[2:].strip()
            date_match = re.search(r'(\d{4}[.\-]\d{1,2}[.\-]\d{1,2})', header)
            if date_match:
                result["date_str"] = date_match.group(1).replace('-', '.')
                clean_title = header.replace(date_match.group(0), '').strip()
                if clean_title:
                    result["title"] = clean_title
            else:
                result["title"] = header
            continue
            
        if stripped.startswith('기도:') or stripped.startswith('대표기도:'):
            prayer_val = re.sub(r'^(기도|대표기도)\s*:\s*', '', stripped).strip()
            if prayer_val:
                result["prayer_person"] = prayer_val
            continue
            
        if stripped.startswith('##'):
            flush_part_buffer()
            raw_title = stripped.lstrip('#').strip()
            clean_title = re.sub(r'^\d+\s*[.\-)]\s*', '', raw_title).strip()
            current_song = {
                "title": clean_title if clean_title else raw_title,
                "routine_raw": "",
                "routine_items": [],
                "parts": {}
            }
            result["songs"].append(current_song)
            current_part = None
            continue
            
        if stripped.startswith('루틴:') or stripped.startswith('루틴 :'):
            if current_song:
                routine_val = re.sub(r'^루틴\s*:\s*', '', stripped).strip()
                current_song["routine_raw"] = routine_val
                current_song["routine_items"] = parse_routine_tokens(routine_val)
            continue
            
        part_match = re.match(r'^\[([A-Za-z0-9가-힣\'`’]+)\]$', stripped)
        if part_match:
            flush_part_buffer()
            current_part = normalize_string(part_match.group(1)).strip().upper()
            continue
            
        if current_song and current_part:
            part_buffer_lines.append(line)
            
    flush_part_buffer()
    return result

def generate_plan_text_from_conti(parsed_data, include_title_blank=True, include_end_blank=True):
    songs = parsed_data.get("songs", [])
    if not songs:
        return ""
        
    plan_blocks = []
    page_num = 1
    
    date_str = parsed_data.get("date_str", "미입력")
    title_str = parsed_data.get("title", "청년 모임")
    
    def add_slide(title_desc, content):
        nonlocal page_num
        clean_content = content.strip()
        if clean_content == "[BLANK]" or not clean_content:
            # 연속 중복 암전 방지
            if plan_blocks and plan_blocks[-1].split('\n')[-1].strip() == "[BLANK]":
                return
            plan_blocks.append(f"[{page_num}페이지] ({title_desc})\n[BLANK]")
            page_num += 1
        else:
            plan_blocks.append(f"[{page_num}페이지] ({title_desc})\n{clean_content}")
            page_num += 1

    # 1. 표지
    add_slide("표지", f"[{date_str}] {title_str}")
    # 2. 찬양 시작 전 암전
    add_slide("찬양 시작 전 암전", "[BLANK]")
    
    for s_idx, song in enumerate(songs):
        song_title = song.get("title", f"곡 {s_idx+1}")
        items = song.get("routine_items", [])
        parts = song.get("parts", {})
        
        # 곡 제목 슬라이드
        add_slide(f"{song_title} 제목", f"[ {song_title} ]")
        
        # 전주 암전
        if include_title_blank:
            add_slide("전주 암전", "[BLANK]")
            
        for (token, repeat_count) in items:
            base_key = get_base_part_key(token)
            major_key = get_major_part_key(token)
            
            if base_key in ['Intro', 'Outro']:
                continue
            elif base_key in ['Interlude', 'Prayer']:
                add_slide(f"{token} 암전", "[BLANK]")
                continue
            elif base_key.startswith('CUSTOM:'):
                custom_lyric = base_key[7:].strip()
                for _ in range(repeat_count):
                    add_slide(f"{song_title} - {token}", custom_lyric)
                continue
                
            # 파트 가사 찾기
            part_slides = None
            token_upper = normalize_string(token).upper()
            base_upper = base_key.upper()
            major_upper = major_key.upper()
            
            if token_upper in parts and parts[token_upper]:
                part_slides = parts[token_upper]
            elif base_upper in parts and parts[base_upper]:
                part_slides = parts[base_upper]
            elif major_upper in parts and parts[major_upper]:
                part_slides = parts[major_upper]
            elif base_key == 'Tag':
                # Tag fallback: Look for [TAG], [TAG1], or last slide of [C]
                if 'TAG' in parts and parts['TAG']:
                    part_slides = parts['TAG']
                elif 'C' in parts and parts['C']:
                    part_slides = [parts['C'][-1]] # 후렴 마지막 슬라이드
                elif 'C1' in parts and parts['C1']:
                    part_slides = [parts['C1'][-1]]
                    
            if part_slides:
                for _ in range(repeat_count):
                    for s_text in part_slides:
                        add_slide(f"{song_title} - {token}", s_text)
            else:
                # 가사 없음 방어
                if re.search(r'\d+', token) and ('(' in token or ')' in token):
                    add_slide(f"{token} 암전", "[BLANK]")
                else:
                    for _ in range(repeat_count):
                        add_slide(f"{song_title} - {token} 가사 없음", f"[{token} 가사를 입력하세요]")
                
        if include_end_blank:
            add_slide("곡 종료 암전", "[BLANK]")
            
    # 찬양 종료 암전
    add_slide("찬양 종료 암전", "[BLANK]")
    
    # 대표기도 슬라이드
    prayer_person = parsed_data.get("prayer_person", "청년")
    add_slide("대표기도", f"대표기도 : {prayer_person}")
    
    return '\n\n'.join(plan_blocks)

def parse_plan_text_to_slides(plan_text):
    if not plan_text or not plan_text.strip():
        return []
        
    blocks = re.split(r'\n\s*\n+', plan_text.strip())
    slides = []
    
    for block in blocks:
        b = block.strip()
        if not b:
            continue
            
        lines = b.split('\n')
        header_line = lines[0].strip()
        
        if re.match(r'^\[\d+페이지\]', header_line):
            content_lines = lines[1:]
        else:
            content_lines = lines
            
        content_text = '\n'.join([c.strip() for c in content_lines if c.strip()]).strip()
        
        if content_text == "[BLANK]" or not content_text:
            slides.append({
                "type": "blank",
                "text": ""
            })
        else:
            slides.append({
                "type": "text",
                "text": content_text
            })
            
    return slides
