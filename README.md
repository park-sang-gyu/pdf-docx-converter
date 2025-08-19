# 📄 PDF ↔ DOCX 변환기

완벽한 문서 변환을 위한 웹 애플리케이션

## ✨ 주요 기능

- 📄 **PDF → DOCX 변환**: 고품질 이미지 변환
- 📝 **DOCX → PDF 변환**: 완벽한 서식 보존
- 🖼️ **이미지 처리**: 원본 비율 유지 및 중앙 정렬
- 🔤 **한글 폰트 지원**: 나눔고딕 자동 다운로드
- 📱 **반응형 웹**: 모든 기기에서 사용 가능

## 🚀 Replit에서 실행

[![Run on Replit](https://replit.com/badge/github/[사용자명]/pdf-docx-converter)](https://replit.com/new/github/[사용자명]/pdf-docx-converter)

## 🛠️ 로컬 설치

```bash
# 저장소 클론
git clone https://github.com/[사용자명]/pdf-docx-converter.git
cd pdf-docx-converter

# 의존성 설치
pip install -r requirements.txt

# 서버 실행
python main.py
```

## 📋 시스템 요구사항

- Python 3.8+
- 100MB 이상의 여유 공간
- 인터넷 연결 (폰트 다운로드용)

## 🔧 기술 스택

- **Backend**: Flask, ReportLab, python-docx
- **Frontend**: HTML5, CSS3, JavaScript
- **이미지 처리**: Pillow, pdf2image
- **OCR**: Tesseract (선택사항)

## 📝 라이선스

MIT License