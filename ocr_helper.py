import pytesseract
from PIL import Image
from pdf2image import convert_from_path

def extract_text_with_ocr(pdf_path, lang='kor+eng'):
    """
    PDF에서 OCR을 사용하여 텍스트 추출
    
    Args:
        pdf_path (str): PDF 파일 경로
        lang (str): OCR 언어 설정 (기본값: 한국어+영어)
    
    Returns:
        list: 각 페이지별 추출된 텍스트 리스트
    """
    try:
        # PDF를 이미지로 변환
        images = convert_from_path(pdf_path, dpi=300)
        
        extracted_texts = []
        for i, image in enumerate(images):
            print(f"페이지 {i+1} OCR 처리 중...")
            
            # OCR로 텍스트 추출
            text = pytesseract.image_to_string(image, lang=lang)
            extracted_texts.append(text)
            
            print(f"페이지 {i+1} 완료 - {len(text)} 글자 추출")
        
        return extracted_texts
        
    except Exception as e:
        print(f"OCR 처리 오류: {e}")
        return []

def test_ocr_with_sample():
    """
    샘플 이미지로 OCR 테스트
    """
    try:
        # 간단한 테스트 이미지 생성
        from PIL import Image, ImageDraw, ImageFont
        
        # 테스트 이미지 생성
        img = Image.new('RGB', (400, 100), color='white')
        draw = ImageDraw.Draw(img)
        
        # 텍스트 추가 (기본 폰트 사용)
        draw.text((10, 10), "Hello World", fill='black')
        draw.text((10, 40), "안녕하세요 테스트입니다", fill='black')
        
        # 임시 파일로 저장
        test_image_path = "test_ocr.png"
        img.save(test_image_path)
        
        # OCR 테스트
        text = pytesseract.image_to_string(Image.open(test_image_path), lang='kor+eng')
        print(f"OCR 결과: {text}")
        
        # 임시 파일 삭제
        import os
        os.remove(test_image_path)
        
        return text
        
    except Exception as e:
        print(f"OCR 테스트 오류: {e}")
        return None

if __name__ == "__main__":
    print("=== OCR 기능 테스트 ===")
    test_ocr_with_sample()