import os
import urllib.request
import zipfile
import shutil

def download_nanumgothic_font():
    """나눔고딕 폰트 자동 다운로드 및 설치"""
    try:
        print("🔄 나눔고딕 폰트 다운로드 시작...")
        
        # fonts 디렉토리 생성
        fonts_dir = os.path.join(os.getcwd(), 'fonts')
        os.makedirs(fonts_dir, exist_ok=True)
        print(f"📁 fonts 디렉토리 확인: {fonts_dir}")
        
        # 나눔고딕 폰트 파일 경로
        font_ttf_path = os.path.join(fonts_dir, "NanumGothic.ttf")
        font_bold_path = os.path.join(fonts_dir, "NanumGothicBold.ttf")
        
        # 이미 폰트가 있는지 확인
        if os.path.exists(font_ttf_path):
            print(f"✅ 나눔고딕 폰트가 이미 존재합니다: {font_ttf_path}")
            file_size = os.path.getsize(font_ttf_path)
            print(f"📊 파일 크기: {file_size:,} bytes")
            return True
        
        # 나눔고딕 폰트 다운로드 URL (네이버 공식)
        font_url = "https://github.com/naver/nanumfont/releases/download/VER2.6/NanumFont_TTF_ALL.zip"
        font_zip_path = os.path.join(fonts_dir, "NanumFont_TTF_ALL.zip")
        
        print(f"📥 나눔고딕 폰트 다운로드 중...")
        print(f"🔗 URL: {font_url}")
        
        # 폰트 ZIP 파일 다운로드
        urllib.request.urlretrieve(font_url, font_zip_path)
        print(f"✅ 다운로드 완료: {font_zip_path}")
        
        # ZIP 파일 크기 확인
        zip_size = os.path.getsize(font_zip_path)
        print(f"📊 ZIP 파일 크기: {zip_size:,} bytes")
        
        # ZIP 파일 압축 해제
        print(f"📦 ZIP 파일 압축 해제 중...")
        with zipfile.ZipFile(font_zip_path, 'r') as zip_ref:
            # ZIP 파일 내용 확인
            file_list = zip_ref.namelist()
            print(f"📋 ZIP 파일 내용: {len(file_list)}개 파일")
            
            # 나눔고딕 TTF 파일들 추출
            extracted_files = []
            for file_info in zip_ref.filelist:
                filename = file_info.filename
                
                # 나눔고딕 TTF 파일만 추출
                if filename.endswith('.ttf') and 'NanumGothic' in filename:
                    # 파일명 단순화
                    if 'Bold' in filename:
                        target_name = 'NanumGothicBold.ttf'
                    else:
                        target_name = 'NanumGothic.ttf'
                    
                    # 파일 추출
                    with zip_ref.open(filename) as source:
                        target_path = os.path.join(fonts_dir, target_name)
                        with open(target_path, 'wb') as target:
                            shutil.copyfileobj(source, target)
                    
                    extracted_files.append(target_name)
                    print(f"✅ 추출 완료: {target_name}")
        
        # ZIP 파일 삭제
        os.remove(font_zip_path)
        print(f"🗑️ 임시 ZIP 파일 삭제: {font_zip_path}")
        
        # 추출된 폰트 파일 확인
        if os.path.exists(font_ttf_path):
            file_size = os.path.getsize(font_ttf_path)
            print(f"🎉 나눔고딕 폰트 설치 완료!")
            print(f"📍 위치: {font_ttf_path}")
            print(f"📊 크기: {file_size:,} bytes")
            
            # Bold 폰트도 확인
            if os.path.exists(font_bold_path):
                bold_size = os.path.getsize(font_bold_path)
                print(f"📍 Bold 폰트: {font_bold_path}")
                print(f"📊 Bold 크기: {bold_size:,} bytes")
            
            return True
        else:
            print(f"❌ 폰트 추출 실패")
            return False
            
    except Exception as e:
        print(f"❌ 나눔고딕 폰트 다운로드 실패: {e}")
        return False

def check_existing_fonts():
    """기존 폰트 파일 확인"""
    fonts_dir = os.path.join(os.getcwd(), 'fonts')
    
    print(f"🔍 fonts 디렉토리 확인: {fonts_dir}")
    
    if os.path.exists(fonts_dir):
        files = os.listdir(fonts_dir)
        print(f"📁 기존 파일들: {len(files)}개")
        
        for file in files:
            file_path = os.path.join(fonts_dir, file)
            if os.path.isfile(file_path):
                size = os.path.getsize(file_path)
                print(f"  📄 {file}: {size:,} bytes")
    else:
        print(f"📁 fonts 디렉토리가 존재하지 않습니다.")

if __name__ == "__main__":
    print("🔤 나눔고딕 폰트 다운로더")
    print("=" * 50)
    
    # 기존 폰트 확인
    check_existing_fonts()
    print()
    
    # 나눔고딕 폰트 다운로드
    success = download_nanumgothic_font()
    
    print()
    print("=" * 50)
    if success:
        print("🎉 나눔고딕 폰트 다운로드 완료!")
        print("✅ 이제 PDF 변환 시 나눔고딕 폰트가 사용됩니다.")
    else:
        print("❌ 나눔고딕 폰트 다운로드 실패")
        print("⚠️ 수동으로 폰트를 다운로드해주세요.")