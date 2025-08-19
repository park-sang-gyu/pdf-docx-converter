import re
from advanced_text_filter import filter_text_blocks

def enhanced_ui_filter(text: str) -> str:
    """더 강력한 UI 노이즈 제거"""
    if not text:
        return ""
    
    lines = text.splitlines()
    filtered_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # 특정 패턴 완전 제거
        skip_patterns = [
            r"변환\s*방식\s*[:：].*",  # 변환 방식: ...
            r"###\s*HTML\s*템플릿.*",   # ### HTML 템플릿...
            r"##\s*🎯\s*\d+\..*",      # ## 🎯 4. ...
            r"```\s*html.*",           # ```html
            r"표준\s*변환\s*\(빠름\)",   # 표준 변환 (빠름)
            r"웹\s*인터페이스\s*개선",   # 웹 인터페이스 개선
        ]
        
        should_skip = False
        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                should_skip = True
                break
        
        if not should_skip:
            filtered_lines.append(line)
    
    # 기본 필터링도 적용
    intermediate_text = "\n".join(filtered_lines)
    final_text = filter_text_blocks(intermediate_text, debug=False)
    
    return final_text

# main.py에서 사용
def clean_extracted_text_enhanced(text):
    """향상된 텍스트 정리"""
    if not text:
        return ""
    
    try:
        # 기본 정리
        import re
        cleaned = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)
        cleaned = re.sub(r'\s+', ' ', cleaned)
        cleaned = cleaned.strip()
        
        # 향상된 필터링 적용
        if cleaned:
            filtered = enhanced_ui_filter(cleaned)
            return filtered if filtered else "(추출된 텍스트가 없습니다)"
        else:
            return "(추출된 텍스트가 없습니다)"
        
    except Exception as e:
        print(f"텍스트 정리 중 오류: {e}")
        return text.strip() if text else "(텍스트 처리 오류)"