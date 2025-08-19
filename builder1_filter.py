import re
from advanced_text_filter import filter_text_blocks

# Builder1 문서 전용 노이즈 패턴
BUILDER1_NOISE_PATTERNS = [
    # 변환 방식 관련
    r"변환\s*방식\s*[:：]?.*",
    r"표준\s*변환\s*\(빠름\)",
    r"고급\s*변환\s*\(정확\)",
    
    # HTML 템플릿 관련
    r"###\s*HTML\s*템플릿\s*업데이트.*",
    r"```html.*",
    r"`index\.html`.*파일.*",
    
    # 웹 인터페이스 관련
    r"##\s*🎯\s*\d+\.\s*웹\s*인터페이스\s*개선.*",
    r"웹\s*인터페이스\s*환경변수.*",
    r"로딩\s*애니메이션.*",
    r"div\s*요소\s*제거",
    
    # 해결 패턴
    r"#\s*해결\s*[:：]?.*",
    r"해결\s*방법\s*[:：]?.*",
    
    # 기타 UI 관련
    r"업로드\s*섹션.*",
    r"클릭\s*이벤트.*",
    r"JavaScript\s*코드.*",
    r"템플릿\s*파일.*"
]

def filter_builder1_content(text: str) -> str:
    """Builder1 문서 전용 필터링"""
    if not text:
        return ""
    
    lines = text.splitlines()
    filtered_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Builder1 노이즈 패턴 확인
        is_noise = False
        for pattern in BUILDER1_NOISE_PATTERNS:
            if re.search(pattern, line, re.IGNORECASE):
                is_noise = True
                break
        
        if not is_noise:
            filtered_lines.append(line)
    
    # 기본 필터링도 적용
    intermediate_text = "\n".join(filtered_lines)
    final_text = filter_text_blocks(intermediate_text, debug=False)
    
    return final_text

def test_builder1_filtering():
    """Builder1 필터링 테스트"""
    sample_text = """
    변환 방식: 표준 변환 (빠름)
    ### HTML 템플릿 업데이트: ```html
    ## 🎯 4. 웹 인터페이스 개선: ```html
    
    실제 중요한 문서 내용입니다.
    회사 공지사항
    수신: 전 직원
    담당자: 김철수
    
    # 해결: 웹 인터페이스 환경변수 관련 업데이트
    `index.html` 파일 로딩 div 요소 제거
    
    PDF 파일을 업로드하면 PPTX로 변환합니다.
    """
    
    print("=== Builder1 전용 필터링 테스트 ===")
    print("원본:")
    print(sample_text)
    
    print("\n필터링 결과:")
    result = filter_builder1_content(sample_text)
    print(result)

if __name__ == "__main__":
    test_builder1_filtering()