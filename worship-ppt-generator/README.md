# 🎵 Worship PPT Generator (찬양 PPT 자동 제작기)

> 교회 찬양팀 콘티 텍스트를 분석하여 16:9 와이드스크린 무결점 찬양 PPT를 자동 생성하는 Streamlit 기반 웹 애플리케이션입니다.

---

## ✨ 주요 기능

1. **텍스트 기반 초고속 콘티 파싱**
   - 곡 정보, 대표기도자, 찬양 순서(루틴), 가사 파트([V], [C], [B], [TAG] 등) 자동 인식
   - 복잡한 악보 루틴 기호 완벽 지원 (Intro(8), (1), (4), Cx2, Bx4, Tag*2, C', (기도), (멘트) 등)

2. **휴먼 게이트 (2단계 기획안 검토)**
   - 생성된 슬라이드 텍스트를 메모장처럼 손쉽게 오탈자 검수 및 [BLANK] 암전 슬라이드 편집 가능

3. **100% 무결점 PPTX 빌드 엔진**
   - 16:9 와이드스크린 (20.0 x 11.25 in), 순수 암전 블랙 배경 (#000000)
   - 81pt Bold 고가독성 흰색 자막 폰트, 양 끝 꽉 찬 텍스트 박스로 자간 줄바꿈 깨짐 원천 방지
   - 순방향(Forward-Only) 진행으로 넘김 실수 제로

---

## 🚀 시작하기

### 1. 필수 요구사항
- Python 3.9 이상

### 2. 패키지 설치
`ash
pip install -r requirements.txt
`

### 3. 애플리케이션 실행
`ash
streamlit run app.py
`
또는 
un_app.bat 파일을 더블 클릭하여 실행합니다.

---

## 📁 프로젝트 구조

`
worship-ppt-generator/
├── app.py                  # Streamlit 웹 애플리케이션 메인
├── core/
│   ├── routine_parser.py   # 찬양 악보 루틴 파싱 엔진
│   ├── text_engine.py      # 콘티 텍스트 분석 및 슬라이드 생성 엔진
│   └── pptx_builder.py     # python-pptx 기반 16:9 슬라이드 생성 엔진
├── requirements.txt        # 의존성 패키지 목록
├── run_app.bat             # 윈도우 원클릭 실행 배치 파일
├── .gitignore              # 깃 무시 파일 설정
└── README.md               # 프로젝트 설명 문서
`

---

## 📝 라이선스 & 제작
- Made by **@loose_lab** (v1.0.0)
