import streamlit as st
import io
import os
import importlib
import core.routine_parser
import core.text_engine
import core.pptx_builder

importlib.reload(core.routine_parser)
importlib.reload(core.text_engine)
importlib.reload(core.pptx_builder)

from core.text_engine import parse_user_conti, generate_plan_text_from_conti, parse_plan_text_to_slides
from core.pptx_builder import build_praise_pptx, get_pptx_bytes, get_preset_ppts, PRESET_DIR

st.set_page_config(
    page_title="LOGOS 찬양 PPT 제작 스튜디오",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (In-app styling)
st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    * {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    .main-title-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        padding: 18px 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-title {
        font-size: 1.65rem;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    .main-subtitle {
        font-size: 0.95rem;
        opacity: 0.9;
    }
    
    /* 1. Hide Deploy Button */
    [data-testid="stAppDeployButton"],
    .stAppDeployButton,
    [data-testid="stDeployButton"] {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* 2. Hide Version Copy Button */
    .stMenuVersionCopyButton {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* 3. Hide Page Bottom Footer */
    footer {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* 4. Tab Bar Styling (Scoped to main container to prevent leaking into sidebar) */
    section.main div[data-testid="stRadio"] > div[role="radiogroup"],
    [data-testid="stMain"] div[data-testid="stRadio"] > div[role="radiogroup"] {
        display: flex;
        flex-direction: row;
        gap: 8px;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 0px;
        margin-bottom: 20px;
    }
    section.main div[data-testid="stRadio"] label,
    [data-testid="stMain"] div[data-testid="stRadio"] label {
        padding: 8px 18px;
        border-radius: 6px 6px 0 0;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        cursor: pointer;
        background: transparent;
        border-bottom: 3px solid transparent;
        margin-bottom: -2px;
        transition: all 0.15s ease-in-out;
    }
    section.main div[data-testid="stRadio"] label > div:first-child,
    [data-testid="stMain"] div[data-testid="stRadio"] label > div:first-child {
        display: none !important;
    }
    section.main div[data-testid="stRadio"] label p,
    [data-testid="stMain"] div[data-testid="stRadio"] label p {
        color: #64748b !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
    }
    section.main div[data-testid="stRadio"] label:hover,
    [data-testid="stMain"] div[data-testid="stRadio"] label:hover {
        background: rgba(37, 99, 235, 0.04);
    }
    section.main div[data-testid="stRadio"] label:has(input:checked),
    [data-testid="stMain"] div[data-testid="stRadio"] label:has(input:checked) {
        border-bottom: 3px solid #2563eb !important;
        background: rgba(37, 99, 235, 0.05);
    }
    section.main div[data-testid="stRadio"] label:has(input:checked) p,
    [data-testid="stMain"] div[data-testid="stRadio"] label:has(input:checked) p {
        color: #2563eb !important;
        font-weight: 800 !important;
    }
</style>
""", unsafe_allow_html=True)

# State initialization
if "conti_text" not in st.session_state:
    st.session_state.conti_text = ""
if "plan_text" not in st.session_state:
    st.session_state.plan_text = ""
if "selected_tab" not in st.session_state:
    st.session_state.selected_tab = "1. 콘티 작성"

# Header
st.markdown("""
<div class="main-title-container">
    <div>
        <div class="main-title">🎵 LOGOS 찬양 PPT 제작 스튜디오</div>
        <div class="main-subtitle">1. 콘티 작성 ➔ 2. 기획안 검토 ➔ 3. PPT 다운로드</div>
    </div>
    <div>
        <span style="background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 16px; font-size: 0.85rem; font-weight: 700; letter-spacing: 0.5px;">
            v1.3.2
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar: Global Settings & File Handlers
with st.sidebar:
    st.header("📁 기준 PPT 설정 [필수]")
    st.caption("텍스트 위치, 배율, 폰트 임베딩의 기준이 되는 원본 PPT입니다.")
    
    template_mode = st.radio(
        "기준 PPT 지정 방식",
        ["교회 표준 프리셋 선택", "내 PC에서 직접 업로드"],
        index=0
    )
    
    active_template_source = None
    active_template_name = None
    
    if template_mode == "교회 표준 프리셋 선택":
        preset_files = get_preset_ppts()
        if preset_files:
            selected_preset = st.radio(
                "표준 프리셋 선택",
                preset_files,
                index=0,
                help="프로젝트에 등록된 교회 표준 찬양 PPT 목록입니다."
            )
            active_template_source = os.path.join(PRESET_DIR, selected_preset)
            active_template_name = selected_preset
            st.caption(f"📌 적용 중: **{selected_preset}**")
        else:
            st.error("등록된 표준 프리셋 파일이 없습니다. 직접 업로드를 선택해 주세요.")
    else:
        uploaded_template = st.file_uploader(
            "기준 PPT 파일 (.pptx)",
            type=["pptx"],
            help="사용할 기준 PPTX 파일을 업로드하세요. 해당 파일의 텍스트 박스 위치, 여백, 폰트 임베딩을 그대로 복제합니다."
        )
        if uploaded_template is not None:
            active_template_source = uploaded_template
            active_template_name = uploaded_template.name
            st.caption(f"📤 적용 중: **{uploaded_template.name}**")
        else:
            st.warning("⚠️ 기준이 될 .pptx 파일을 업로드해야 PPT 생성이 가능합니다.")

    st.divider()
    st.header("📁 콘티 파일 업로드")
    uploaded_file = st.file_uploader(
        "작성한 콘티 파일 (.txt)",
        type=["txt"],
        help="콘티 텍스트 파일을 업로드하면 에디터에 내용이 자동으로 입력됩니다."
    )
    if uploaded_file is not None:
        file_content = uploaded_file.read().decode("utf-8")
        if file_content != st.session_state.conti_text:
            st.session_state.conti_text = file_content
            st.session_state.conti_editor = file_content
            st.rerun()

    st.divider()
    st.header("💡 슬라이드 암전 규칙")
    title_blank = st.checkbox("곡 제목 뒤 암전(전주/준비) 자동 삽입", value=True)
    end_blank = st.checkbox("곡 종료 시 암전(곡 전환) 자동 삽입", value=True)

# Stateful Tabs (100% Cross-origin Cloud Safe)
tab_options = ["1. 콘티 작성", "2. PPT 기획안"]

if "switch_to_tab" in st.session_state:
    st.session_state.selected_tab = st.session_state.switch_to_tab
    del st.session_state.switch_to_tab

if "selected_tab" not in st.session_state or st.session_state.selected_tab not in tab_options:
    st.session_state.selected_tab = "1. 콘티 작성"

active_tab = st.radio(
    "메뉴 탭 선택",
    tab_options,
    key="selected_tab",
    horizontal=True,
    label_visibility="collapsed"
)

# ==========================================
# TAB 1: CONTI WRITING
# ==========================================
if active_tab == "1. 콘티 작성":
    st.subheader("1단계: 콘티 서식 작성")
    st.caption("고정 정보(# 모임명, 기도자), 곡 제목(##), 루틴 문자열, 파트([V], [C] 등)를 작성하세요. 슬라이드는 **빈 줄(더블 엔터)**로 나뉩니다.")
    
    col_edit, col_guide = st.columns([8, 4])
    
    with col_edit:
        user_conti = st.text_area(
            "콘티 본문",
            value=st.session_state.conti_text,
            height=520,
            key="conti_editor",
            placeholder="# 2026.09.06 LOGOS 청년 모임\n기도: ㅇㅇㅇ 청년\n\n## 1. 찬양 제목\n루틴: Intro - V - C...\n\n[V]\n가사 1행\n가사 2행\n\n[C]\n후렴 1행\n후렴 2행",
            label_visibility="collapsed"
        )
        st.session_state.conti_text = user_conti
            
        if st.button("PPT 기획안 생성", type="primary", use_container_width=True):
            if not st.session_state.conti_text.strip():
                st.warning("⚠️ 콘티 내용을 입력하거나 .txt 파일을 업로드한 후 버튼을 눌러주세요!")
            else:
                parsed = parse_user_conti(st.session_state.conti_text)
                new_plan = generate_plan_text_from_conti(
                    parsed,
                    include_title_blank=title_blank,
                    include_end_blank=end_blank
                )
                st.session_state.plan_text = new_plan
                st.session_state.plan_editor = new_plan
                st.session_state.switch_to_tab = "2. PPT 기획안"
                st.rerun()
            
    with col_guide:
        st.markdown("### 작성 규칙 안내")
        st.markdown("""
        - **표지 정보**: `# 2026.09.06 LOGOS 청년 모임`
        - **대표기도자**: `기도: ㅇㅇㅇ 청년`
        - **찬양곡 시작**: `## 1. 새 힘 얻으리`
        - **루틴**: `루틴: Intro(8) - V - V - P - C...` (악보에서 복사)
        - **파트 선언**: `[V]`, `[V1]`, `[V2]`, `[P]`, `[C]`, `[B]` 등
        - **슬라이드 분할**: **빈 줄 하나(더블 엔터)**를 넣으면 다음 슬라이드로 넘어갑니다.
        """)
        
        st.divider()
        parsed_preview = parse_user_conti(st.session_state.get("conti_editor", st.session_state.conti_text))
        st.markdown(f"- **인식된 날짜**: `{parsed_preview.get('date_str')}`")
        st.markdown(f"- **인식된 모임명**: `{parsed_preview.get('title')}`")
        st.markdown(f"- **인식된 대표기도자**: `{parsed_preview.get('prayer_person')}`")
        st.markdown(f"- **인식된 곡 수**: `{len(parsed_preview.get('songs'))}곡`")
        for idx, s in enumerate(parsed_preview.get('songs', [])):
            st.markdown(f"  - **곡 {idx+1}**: {s.get('title')} ({len(s.get('parts', {}))}개 파트)")

# ==========================================
# TAB 2: PPT PLAN TEXT (HUMAN GATE)
# ==========================================
elif active_tab == "2. PPT 기획안":
    st.subheader("2단계: PPT 기획안 검토 및 수정")
    st.caption("생성된 슬라이드 텍스트를 검토하고 오탈자나 [BLANK] 암전 위치를 자유롭게 수정하세요.")
    
    current_plan = st.session_state.get("plan_editor", st.session_state.get("plan_text", ""))
    current_slides = parse_plan_text_to_slides(current_plan)
    total_slide_count = len(current_slides)
    
    c_stat1, c_stat2 = st.columns([3, 9])
    with c_stat1:
        st.metric("총 슬라이드 수", f"{total_slide_count} 장")
    with c_stat2:
        template_display = active_template_name if active_template_name else "미지정 (사이드바 설정 필요)"
        st.metric("적용 기준 PPT", template_display)
        
    st.divider()
    
    plan_user_input = st.text_area(
        "기획안 텍스트",
        value=st.session_state.plan_text,
        height=520,
        key="plan_editor",
        placeholder="1단계에서 콘티를 작성하고 [PPT 기획안 생성] 버튼을 누르면 이곳에 전체 슬라이드 기획안이 생성됩니다.",
        label_visibility="collapsed"
    )
    if plan_user_input != st.session_state.plan_text:
        st.session_state.plan_text = plan_user_input
        current_slides = parse_plan_text_to_slides(plan_user_input)
        
    st.divider()
    
    is_template_ready = (active_template_source is not None)
    
    if not is_template_ready:
        st.warning("⚠️ **[필수] 기준 PPT가 설정되지 않았습니다.**\n\n좌측 사이드바의 **「📁 기준 PPT 설정」**에서 [교회 표준 프리셋]을 선택하거나 [직접 파일 업로드]를 완료해야 PPT를 생성할 수 있습니다.")
    else:
        st.info(f"🎯 **적용 중인 기준 PPT**: `{active_template_name}`")
        
    if st.button("PPT 파일 생성", type="primary", use_container_width=True, disabled=not is_template_ready):
        if not current_slides:
            st.warning("⚠️ 슬라이드 기획안 내용이 비어있습니다. 1단계에서 콘티를 작성 후 생성해 주세요.")
        else:
            with st.spinner("기준 PPT 서식 및 폰트 임베딩을 1:1 딥클론하여 PPTX 빌드 중..."):
                prs = build_praise_pptx(current_slides, template_source=active_template_source)
                pptx_bytes = get_pptx_bytes(prs)
                
                parsed = parse_user_conti(st.session_state.get("conti_editor", st.session_state.conti_text))
                raw_date = parsed.get("date_str", "20260906")
                clean_date = raw_date.replace(".", "").replace(" ", "").replace("-", "") if raw_date != "미입력" else "찬양콘티"
                filename = f"{clean_date} 찬양 가사.pptx"
                
                st.download_button(
                    label="PPT 파일 다운로드",
                    data=pptx_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True
                )
                st.success(f"{len(current_slides)}장의 찬양 PPT 파일이 준비되었습니다. 위 버튼을 눌러 다운로드하세요.")

# Global Style & DOM Injector (Localization + Custom Footer)
st.components.v1.html("""
<script>
    try {
        const pDoc = window.parent.document;
        
        // 1. Inject Head CSS
        let s = pDoc.getElementById('custom-header-footer-style');
        if (!s) {
            s = pDoc.createElement('style');
            s.id = 'custom-header-footer-style';
            pDoc.head.appendChild(s);
        }
        s.innerHTML = `
            [data-testid="stAppDeployButton"],
            .stAppDeployButton,
            [data-testid="stDeployButton"] {
                display: none !important;
                visibility: hidden !important;
            }
            .stMenuVersionCopyButton {
                display: none !important;
                visibility: hidden !important;
            }
            footer {
                display: none !important;
                visibility: hidden !important;
            }
        `;

        // 2. Real DOM Text Replacement & Settings Menu Localization
        const updateMenuLocalization = () => {
            // A. Remove Copy Button & Update Footer
            const copyBtns = pDoc.querySelectorAll('.stMenuVersionCopyButton');
            copyBtns.forEach(btn => {
                btn.style.display = 'none';
                const parent = btn.parentElement;
                if (parent) {
                    const textEl = parent.firstElementChild;
                    if (textEl && textEl !== btn) {
                        if (textEl.textContent !== 'Made by @loose_lab v1.3.2') {
                            textEl.textContent = 'Made by @loose_lab v1.3.2';
                            textEl.style.fontSize = '0.82rem';
                            textEl.style.color = '#808495';
                            textEl.style.userSelect = 'text';
                            textEl.style.webkitUserSelect = 'text';
                            textEl.style.cursor = 'text';
                        }
                    }
                }
            });

            // B. Hide Deploy Button
            const deployBtn = pDoc.querySelector('[data-testid="stAppDeployButton"], .stAppDeployButton, [data-testid="stDeployButton"]');
            if (deployBtn) {
                deployBtn.style.display = 'none';
            }

            // C. Localize Settings Menu Items (Korean)
            const popover = pDoc.querySelector('div[data-baseweb="popover"], div[data-testid="stMainMenuPopover"]');
            if (popover) {
                const walker = pDoc.createTreeWalker(popover, NodeFilter.SHOW_TEXT, null, false);
                let node;
                while (node = walker.nextNode()) {
                    const val = node.nodeValue.trim();
                    if (val === 'System') {
                        node.nodeValue = '시스템';
                    } else if (val === 'Light') {
                        node.nodeValue = '라이트';
                    } else if (val === 'Dark') {
                        node.nodeValue = '다크';
                    } else if (val.startsWith('Rerun')) {
                        node.nodeValue = node.nodeValue.replace(/Rerun/g, '다시 실행');
                    } else if (val === 'Auto rerun') {
                        node.nodeValue = '자동 다시 실행';
                    } else if (val.startsWith('Clear cache')) {
                        node.nodeValue = node.nodeValue.replace(/Clear cache/g, '캐시 삭제');
                    } else if (val === 'Print') {
                        node.nodeValue = '인쇄';
                    } else if (val === 'Record screen') {
                        node.nodeValue = '화면 녹화';
                    }
                }
            }
        };

        const observer = new MutationObserver(() => {
            updateMenuLocalization();
        });
        observer.observe(pDoc.body, { childList: true, subtree: true });
        updateMenuLocalization();
    } catch(e) {}
</script>
""", height=0)

