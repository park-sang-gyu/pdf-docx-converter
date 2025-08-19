import os
import time
import psutil
from pathlib import Path

def is_file_locked(file_path):
    """
    파일이 다른 프로세스에 의해 잠겨있는지 확인
    
    Args:
        file_path (str): 확인할 파일 경로
    
    Returns:
        bool: 파일이 잠겨있으면 True, 아니면 False
    """
    try:
        # 파일이 존재하지 않으면 잠겨있지 않음
        if not os.path.exists(file_path):
            return False
        
        # 파일을 쓰기 모드로 열어보기
        with open(file_path, 'a'):
            pass
        return False
        
    except (IOError, OSError, PermissionError):
        return True

def find_processes_using_file(file_path):
    """
    파일을 사용 중인 프로세스 찾기
    
    Args:
        file_path (str): 확인할 파일 경로
    
    Returns:
        list: 파일을 사용 중인 프로세스 정보 리스트
    """
    processes = []
    file_path = os.path.abspath(file_path)
    
    try:
        for proc in psutil.process_iter(['pid', 'name', 'open_files']):
            try:
                if proc.info['open_files']:
                    for file_info in proc.info['open_files']:
                        if os.path.abspath(file_info.path) == file_path:
                            processes.append({
                                'pid': proc.info['pid'],
                                'name': proc.info['name'],
                                'path': file_info.path
                            })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
                
    except Exception as e:
        print(f"프로세스 검색 오류: {e}")
    
    return processes

def wait_for_file_unlock(file_path, max_wait_seconds=30, check_interval=1):
    """
    파일 잠금이 해제될 때까지 대기
    
    Args:
        file_path (str): 대기할 파일 경로
        max_wait_seconds (int): 최대 대기 시간 (초)
        check_interval (int): 확인 간격 (초)
    
    Returns:
        bool: 파일 잠금이 해제되면 True, 시간 초과시 False
    """
    start_time = time.time()
    
    while time.time() - start_time < max_wait_seconds:
        if not is_file_locked(file_path):
            return True
        
        print(f"파일 잠금 대기 중... ({int(time.time() - start_time)}초)")
        time.sleep(check_interval)
    
    return False

def safe_file_operation(file_path, operation_func, *args, **kwargs):
    """
    안전한 파일 작업 수행
    
    Args:
        file_path (str): 작업할 파일 경로
        operation_func (callable): 수행할 작업 함수
        *args, **kwargs: 작업 함수에 전달할 인자들
    
    Returns:
        tuple: (성공 여부, 결과 또는 오류 메시지)
    """
    try:
        # 1. 파일 잠금 확인
        if is_file_locked(file_path):
            print(f"⚠️ 파일이 사용 중입니다: {file_path}")
            
            # 사용 중인 프로세스 찾기
            processes = find_processes_using_file(file_path)
            if processes:
                print("파일을 사용 중인 프로세스:")
                for proc in processes:
                    print(f"  - {proc['name']} (PID: {proc['pid']})")
                
                # PowerPoint 프로세스인 경우 특별 안내
                ppt_processes = [p for p in processes if 'powerpoint' in p['name'].lower() or 'pptx' in p['name'].lower()]
                if ppt_processes:
                    print("\n💡 PowerPoint가 파일을 사용 중입니다.")
                    print("   해결 방법: PowerPoint에서 파일을 닫고 다시 시도하세요.")
            
            # 파일 잠금 해제 대기
            print("파일 잠금 해제를 기다리는 중...")
            if wait_for_file_unlock(file_path, max_wait_seconds=30):
                print("✅ 파일 잠금이 해제되었습니다.")
            else:
                return False, "파일 잠금 해제 시간 초과"
        
        # 2. 작업 수행
        result = operation_func(*args, **kwargs)
        return True, result
        
    except PermissionError as e:
        return False, f"권한 오류: {e}"
    except Exception as e:
        return False, f"작업 오류: {e}"

def generate_safe_filename(original_name, max_length=100):
    """
    안전한 파일명 생성 (ASCII 문자만 사용)
    
    Args:
        original_name (str): 원본 파일명
        max_length (int): 최대 길이
    
    Returns:
        str: 안전한 파일명
    """
    import re
    import hashlib
    from datetime import datetime
    
    # 파일명과 확장자 분리
    name_part, ext_part = os.path.splitext(original_name)
    
    # ASCII가 아닌 문자 제거 및 공백을 언더스코어로 변경
    safe_name = re.sub(r'[^a-zA-Z0-9_.-]', '_', name_part)
    safe_name = re.sub(r'_+', '_', safe_name)  # 연속된 언더스코어 정리
    safe_name = safe_name.strip('_')  # 앞뒤 언더스코어 제거
    
    # 이름이 너무 짧거나 비어있으면 해시값 사용
    if len(safe_name) < 3:
        hash_value = hashlib.md5(original_name.encode('utf-8')).hexdigest()[:8]
        safe_name = f"file_{hash_value}"
    
    # 길이 제한
    if len(safe_name) > max_length - len(ext_part) - 10:  # 타임스탬프 여유분
        safe_name = safe_name[:max_length - len(ext_part) - 10]
    
    # 타임스탬프 추가 (중복 방지)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    return f"{safe_name}_{timestamp}{ext_part}"

def cleanup_temp_files(temp_dir, pattern="*.png", max_age_hours=24):
    """
    임시 파일 정리
    
    Args:
        temp_dir (str): 임시 파일 디렉토리
        pattern (str): 삭제할 파일 패턴
        max_age_hours (int): 최대 보관 시간 (시간)
    """
    import glob
    from datetime import datetime, timedelta
    
    try:
        temp_path = Path(temp_dir)
        if not temp_path.exists():
            return
        
        cutoff_time = datetime.now() - timedelta(hours=max_age_hours)
        
        for file_path in glob.glob(os.path.join(temp_dir, pattern)):
            try:
                file_stat = os.stat(file_path)
                file_time = datetime.fromtimestamp(file_stat.st_mtime)
                
                if file_time < cutoff_time:
                    os.remove(file_path)
                    print(f"임시 파일 삭제: {file_path}")
                    
            except Exception as e:
                print(f"임시 파일 삭제 실패 {file_path}: {e}")
                
    except Exception as e:
        print(f"임시 파일 정리 오류: {e}")

if __name__ == "__main__":
    # 테스트 코드
    print("=== 파일 유틸리티 테스트 ===")
    
    # 안전한 파일명 생성 테스트
    test_names = [
        "한글파일명.pptx",
        "File with spaces.pdf",
        "Special@#$%Characters.docx",
        "Very_Long_File_Name_That_Exceeds_Normal_Length_Limits_And_Should_Be_Truncated.xlsx"
    ]
    
    print("\n안전한 파일명 생성 테스트:")
    for name in test_names:
        safe_name = generate_safe_filename(name)
        print(f"  원본: {name}")
        print(f"  변환: {safe_name}\n")
    
    # 임시 파일 정리 테스트
    print("임시 파일 정리 테스트:")
    cleanup_temp_files("./temp", "*.png", 1)  # 1시간 이상 된 PNG 파일 삭제