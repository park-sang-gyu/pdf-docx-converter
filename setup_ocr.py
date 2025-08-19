import os
import urllib.request
import zipfile
import subprocess
import sys

def install_tesseract_windows():
    """Windows용 Tesseract OCR 자동 설치"""
    try:
        print("🔍 Tesseract OCR 설치 시작...")
        
        # Tesseract 설치 파일 다운로드 URL
        tesseract_url = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
        installer_path = "tesseract_installer.exe"
        
        print(f"📥 Tesseract 다운로드 중: {tesseract_url}")
        urllib.request.urlretrieve(tesseract_url, installer_path)
        
        print("🔧 Tesseract 설치 중... (관리자 권한 필요)")
        print("⚠️ 설치 창이 나타나면 다음 옵션을 선택하세요:")
        print("   - Additional language data: Korean 체크")
        print("   - Add to PATH 체크")
        
        # 설치 프로그램 실행
        subprocess.run([installer_path], check=True)
        
        # 설치 파일 삭제
        os.remove(installer_path)
        
        print("✅ Tesseract OCR 설치 완료!")
        return True
        
    except Exception as e:
        print(f"❌ Tesseract 설치 실패: {e}")
        return False

def install_python_packages():
    """필요한 Python 패키지 설치"""
    try:
        packages = [
            'pytesseract==0.3.10',
            'opencv-python==4.8.1.78'
        ]
        
        for package in packages:
            print(f"📦 {package} 설치 중...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        
        print("✅ Python 패키지 설치 완료!")
        return True
        
    except Exception as e:
        print(f"❌ 패키지 설치 실패: {e}")
        return False

def download_korean_tessdata():
    """한글 OCR 데이터 다운로드"""
    try:
        print("🇰🇷 한글 OCR 데이터 다운로드 중...")
        
        # Tesseract 데이터 경로 찾기
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tessdata',
            r'C:\Program Files (x86)\Tesseract-OCR\tessdata'
        ]
        
        tessdata_path = None
        for path in possible_paths:
            if os.path.exists(path):
                tessdata_path = path
                break
        
        if not tessdata_path:
            print("❌ Tesseract tessdata 폴더를 찾을 수 없습니다.")
            return False
        
        # 한글 데이터 파일 다운로드
        korean_url = "https://github.com/tesseract-ocr/tessdata/raw/main/kor.traineddata"
        korean_path = os.path.join(tessdata_path, "kor.traineddata")
        
        if not os.path.exists(korean_path):
            print(f"📥 한글 데이터 다운로드: {korean_url}")
            urllib.request.urlretrieve(korean_url, korean_path)
            print(f"✅ 한글 데이터 저장: {korean_path}")
        else:
            print(f"✅ 한글 데이터 이미 존재: {korean_path}")
        
        return True
        
    except Exception as e:
        print(f"❌ 한글 데이터 다운로드 실패: {e}")
        return False

if __name__ == "__main__":
    print("🔍 OCR 환경 설정 시작")
    print("=" * 50)
    
    # 1. Python 패키지 설치
    print("1️⃣ Python 패키지 설치")
    install_python_packages()
    print()
    
    # 2. Tesseract 설치 (Windows)
    if os.name == 'nt':  # Windows
        print("2️⃣ Tesseract OCR 설치")
        install_tesseract_windows()
        print()
        
        # 3. 한글 데이터 다운로드
        print("3️⃣ 한글 OCR 데이터 다운로드")
        download_korean_tessdata()
    
    print()
    print("=" * 50)
    print("🎉 OCR 환경 설정 완료!")
    print("✅ 이제 PDF → PPTX 변환 시 OCR 기능이 자동으로 사용됩니다.")
    print("🔄 서버를 재시작하세요: python main.py")