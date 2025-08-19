import re

# 기존 설정
TH_NOISE = 5

# ---------------- 새로 추가/강화되는 상수 ----------------
NUKE_MODE = True  # True면 아래 핵심 토큰 포함 라인 즉시 제거

# 기존 NUKE_TOKENS에 추가
NUKE_TOKENS = [
    "변환 방식", "표준 변환", "빠름", "html 템플릿 업데이트",
    "웹 인터페이스 개선", "해랍북스", "2024년도", "도서목록",
    "DIAT", "ITO", "수험서", "출간사", "가격", "대상", "비고"
]

# 새로운 패턴 추가 (기존 코드 뒤에 추가)
LONG_REPETITIVE_PATTERN = re.compile(r"(해랍북스|DIAT|ITO|수험서|출간사).{0,50}(해랍북스|DIAT|ITO|수험서|출간사)")
TABLE_METADATA_PATTERN = re.compile(r"^(교재명|출간사|가격|대상|비고)\s*[:：]?\s*")
REPETITIVE_NUMBERS = re.compile(r"\d{1,2}[,-]\d{1,2}[급단계]")

# 새로운 강화된 필터링 함수 추가
def remove_long_repetitive_content(lines: list[str]) -> list[str]:
    """긴 반복 콘텐츠 제거 (테이블 메타데이터, 반복 패턴 등)"""
    if not lines:
        return lines
    
    filtered_lines = []
    repetitive_count = 0
    max_repetitive_lines = 10  # 연속 반복 라인 최대 허용 수
    
    for line in lines:
        line_stripped = line.strip()
        
        # 빈 라인은 통과
        if not line_stripped:
            filtered_lines.append(line)
            repetitive_count = 0
            continue
        
        # 테이블 메타데이터 패턴 감지
        if TABLE_METADATA_PATTERN.search(line_stripped):
            repetitive_count += 1
            if repetitive_count > max_repetitive_lines:
                continue  # 너무 많은 반복이면 스킵
        
        # 긴 반복 패턴 감지
        elif LONG_REPETITIVE_PATTERN.search(line_stripped):
            repetitive_count += 1
            if repetitive_count > max_repetitive_lines:
                continue  # 너무 많은 반복이면 스킵
        
        # 반복적인 숫자 패턴 (급수, 단계 등)
        elif REPETITIVE_NUMBERS.search(line_stripped):
            repetitive_count += 1
            if repetitive_count > max_repetitive_lines:
                continue
        
        # 라인이 너무 길면 (200자 이상) 잠재적 메타데이터로 간주
        elif len(line_stripped) > 200:
            # 하지만 한글 비율이 높으면 실제 내용일 수 있으므로 보존
            hangul_ratio_val = hangul_ratio(line_stripped)
            if hangul_ratio_val < 0.3:  # 한글 비율이 30% 미만이면 제거
                continue
        
        else:
            repetitive_count = 0  # 정상 라인이면 카운터 리셋
        
        filtered_lines.append(line)
    
    return filtered_lines

# 기존 코드에 추가할 함수들

def hangul_ratio(line: str) -> float:
    """한글 비율 계산"""
    if not line.strip():
        return 0.0
    hangul_pattern = re.compile(r"[가-힣]")
    h = len(hangul_pattern.findall(line))
    return h / max(1, len(line))

def early_block_filter(raw_text: str) -> list[str]:
    """초기 블록 필터링"""
    lines = [l for l in raw_text.splitlines()]
    # 공백 제거
    normed = [l.rstrip("\n") for l in lines]
    # 불필요한 완전 공백 제거
    cleaned = [l for l in normed if l.strip()]
    return cleaned

def classify_lines(lines: list[str], debug=False):
    """라인 분류 및 점수 부여"""
    classified = []
    for line in lines:
        score = ui_noise_score(line)
        classified.append((line, score))
        if debug:
            print(f"[{score:2d}] {line[:50]}...")
    return classified

def ui_noise_score(line: str) -> int:
    """UI 노이즈 점수 계산"""
    if not line.strip():
        return 0
    
    score = 0
    line_lower = line.lower()
    
    # NUKE 토큰 체크
    for token in NUKE_TOKENS:
        if token.lower() in line_lower:
            score += 10
    
    # 강한 UI 키워드 체크
    for keyword in STRONG_UI_KEYWORDS:
        if keyword.lower() in line_lower:
            score += 5
    
    # 반복 패턴 체크
    if LONG_REPETITIVE_PATTERN.search(line):
        score += 8
    
    if TABLE_METADATA_PATTERN.search(line):
        score += 6
    
    return score

def dynamic_cutoff(noise_classified):
    """동적 컷오프 적용"""
    return noise_classified

def second_pass_nuke(kept_lines):
    """2차 제거 - UI 키워드 비율이 높으면 해당 라인들 삭제"""
    if not kept_lines:
        return kept_lines
    
    norm_lines = [line.lower().replace(" ", "") for line in kept_lines]
    ui_hits = sum(1 for n in norm_lines if any(k in n for k in NOISE_TRIGGER_KEYWORDS))
    
    if ui_hits / len(kept_lines) >= 0.30:
        # 30% 이상이면 해당 키워드 포함 라인 삭제
        filtered = []
        for l, n in zip(kept_lines, norm_lines):
            if any(k in n for k in NOISE_TRIGGER_KEYWORDS):
                continue
            filtered.append(l)
        return filtered if filtered else kept_lines
    
    return kept_lines

def recover_if_too_few(original, filtered):
    """텍스트가 너무 적으면 공문서 키워드로 복구 시도"""
    PUBLIC_DOC_KEYWORDS = [
        "수신", "제목", "붙임", "담당", "연락처", "회신", "협조", "안내",
        "신청", "요청", "공지", "배포", "검토", "회의", "공문"
    ]
    
    if len(filtered) == 0 or len(filtered) <= 3:
        recover = [l for l in original if any(k in l for k in PUBLIC_DOC_KEYWORDS)]
        if recover:
            return recover
        # 마지막 보호: 원본 상위 15줄
        return original[:15]
    return filtered

def final_compact(lines):
    """최종 압축 - 중복 제거 & 공백 다듬기"""
    seen = set()
    out = []
    for l in lines:
        # 정규화된 키로 중복 체크
        key = "".join(l.split()).lower()
        if key in seen or len(key) < 3:  # 너무 짧은 라인도 제거
            continue
        seen.add(key)
        out.append(l.strip())
    return out

# 새로운 중복 제거 함수 추가
def remove_duplicate_content(lines: list[str]) -> list[str]:
    """중복 콘텐츠 제거 (연속된 중복 블록 감지)"""
    if not lines:
        return lines
    
    filtered_lines = []
    seen_blocks = set()
    current_block = []
    
    for line in lines:
        line_stripped = line.strip()
        
        if not line_stripped:
            # 빈 라인이면 현재 블록 처리
            if current_block:
                block_key = "\n".join(current_block).lower()
                if block_key not in seen_blocks:
                    seen_blocks.add(block_key)
                    filtered_lines.extend(current_block)
                    filtered_lines.append(line)  # 빈 라인도 추가
                current_block = []
            continue
        
        current_block.append(line)
        
        # 블록이 5줄 이상이면 중복 체크
        if len(current_block) >= 5:
            block_key = "\n".join(current_block).lower()
            if block_key not in seen_blocks:
                seen_blocks.add(block_key)
                filtered_lines.extend(current_block)
            current_block = []
    
    # 마지막 블록 처리
    if current_block:
        block_key = "\n".join(current_block).lower()
        if block_key not in seen_blocks:
            filtered_lines.extend(current_block)
    
    return filtered_lines

# filter_text_blocks 함수 수정
def filter_text_blocks(raw_text: str, debug=False) -> str:
    """통합: 기존 filter_text_blocks 재정의 (점수 기반 필터 호출 전)"""
    # 0) Early block nuke
    early_lines = early_block_filter(raw_text)
    
    # 0.3) 중복 콘텐츠 제거 (새로 추가)
    early_lines = remove_duplicate_content(early_lines)
    
    # 0.5) 긴 반복 콘텐츠 제거
    early_lines = remove_long_repetitive_content(early_lines)
    
    # 1) 점수 기반
    classified = classify_lines(early_lines, debug=debug)
    after_cut = dynamic_cutoff(classified)
    kept = [l for (l, sc) in after_cut if sc < TH_NOISE]
    
    kept2 = second_pass_nuke(kept)
    recovered = recover_if_too_few(early_lines, kept2)
    final_lines = final_compact(recovered)
    
    return "\n".join(final_lines)

if __name__ == "__main__":
    # 테스트 코드
    test_text = """
    변환 방식:
    표준 변환 (빠름)
    ### HTML 템플릿 업데이트:
    ```html
    <div>파일을 선택하세요</div>
    ```
    ## 🎯 4. 웹 인터페이스 개선:
    실제 문서 내용입니다.
    이것은 보존되어야 할 텍스트입니다.
    공지사항: 중요한 내용
    담당자: 홍길동
    """
    
    print("=== 원본 텍스트 ===")
    print(test_text)
    
    print("\n=== 강화된 필터링 결과 ===")
    filtered = filter_text_blocks(test_text, debug=True)
    print(filtered)

# STRONG_UI_KEYWORDS에 추가
STRONG_UI_KEYWORDS = [
    "변환방식", "변환방식:", "html템플릿업데이트", "파일을선택", "파일선택",
    "클릭하여파일선택", "pdfpptx", "pptx-pdf", "p p t x", "다운로드됩니다",
    "ocr지원", "이미지품질", "슬라이드당텍스트", "dpi", "pptx", "pdf",
    "업로드", "변환완료", "변환하기",
    # 새로 추가된 키워드들
    "해랍북스", "2024년도", "도서목록", "수험서", "출간사", "교재명",
    "DIAT", "ITO", "급수", "단계", "가격", "대상", "비고"
]

# NOISE_TRIGGER_KEYWORDS에 추가
NOISE_TRIGGER_KEYWORDS = [
    "변환", "파일", "pptx", "pdf", "템플릿", "업데이트",
    # 새로 추가
    "해랍북스", "수험서", "출간사", "교재명", "도서목록"
]