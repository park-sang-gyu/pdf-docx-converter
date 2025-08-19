import os
from pathlib import Path
from file_utils import is_file_locked, find_processes_using_file

def pre_conversion_check(input_path, output_path):
    """
    변환 전 사전 체크
    
    Args:
        input_path (str): 입력 파일 경로
        output_path (str): 출력 파일 경로
    
    Returns:
        tuple: (체크 통과 여부, 메시지 리스트)
    """
    issues = []
    warnings = []
    
    print("=== 변환 전 사전 체크 ===")
    
    # 1. 입력 파일 존재 확인
    if not os.path.exists(input_path):
        issues.append(f"❌ 입력 파일이 존재하지 않습니다: {input_path}")
    else:
        print(f"✅ 입력 파일 존재: {input_path}")
        
        # 파일 크기 확인
        file_size = os.path.getsize(input_path) / (1024 * 1024)  # MB
        if file_size > 100:
            warnings.append(f"⚠️ 파일 크기가 큽니다: {file_size:.1f}MB (처리 시간이 오래 걸릴 수 있음)")
        else:
            print(f"✅ 파일 크기 적정: {file_size:.1f}MB")
    
    # 2. 출력 디렉토리 확인
    output_dir = os.path.dirname(output_path)
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            print(f"✅ 출력 디렉토리 생성: {output_dir}")
        except Exception as e:
            issues.append(f"❌ 출력 디렉토리 생성 실패: {e}")
    else:
        print(f"✅ 출력 디렉토리 존재: {output_dir}")
    
    # 3. 출력 파일 잠금 확인
    if os.path.exists(output_path):
        if is_file_locked(output_path):
            processes = find_processes_using_file(output_path)
            if processes:
                process_names = [p['name'] for p in processes]
                issues.append(f"❌ 출력 파일이 사용 중입니다: {', '.join(process_names)}")
                issues.append("💡 해결 방법: PowerPoint 등 관련 프로그램을 닫고 다시 시도하세요.")
            else:
                warnings.append(f"⚠️ 출력 파일에 접근할 수 없습니다: {output_path}")
        else:
            print(f"✅ 출력 파일 접근 가능: {output_path}")
    
    # 4. 디스크 공간 확인
    try:
        import shutil
        free_space = shutil.disk_usage(output_dir).free / (1024 * 1024)  # MB
        if free_space < 100:  # 100MB 미만
            warnings.append(f"⚠️ 디스크 공간 부족: {free_space:.1f}MB 남음")
        else:
            print(f"✅ 디스크 공간 충분: {free_space:.1f}MB 사용 가능")
    except Exception as e:
        warnings.append(f"⚠️ 디스크 공간 확인 실패: {e}")
    
    # 5. 경로 문자 확인
    if any(ord(c) > 127 for c in input_path):
        print(f"✅ 입력 경로에 한글 포함 (지원됨): {input_path}")
    
    if any(ord(c) > 127 for c in output_path):
        print(f"✅ 출력 경로에 한글 포함 (지원됨): {output_path}")
    
    # 결과 출력
    print("\n=== 체크 결과 ===")
    
    if issues:
        print("❌ 발견된 문제:")
        for issue in issues:
            print(f"  {issue}")
    
    if warnings:
        print("⚠️ 주의사항:")
        for warning in warnings:
            print(f"  {warning}")
    
    if not issues and not warnings:
        print("✅ 모든 체크 통과! 변환을 시작할 수 있습니다.")
    
    return len(issues) == 0, issues + warnings

if __name__ == "__main__":
    # 테스트
    input_file = "uploads/pdf/test.pdf"
    output_file = "outputs/test.pptx"
    
    success, messages = pre_conversion_check(input_file, output_file)
    
    if success:
        print("\n🚀 변환 준비 완료!")
    else:
        print("\n🛑 문제를 해결한 후 다시 시도하세요.")