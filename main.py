#!/usr/bin/env python3
"""
PDF ↔ DOCX 변환기 - Replit 배포용
메인 진입점
"""

import os
import sys

# 현재 디렉토리를 Python 경로에 추가
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# final_server 모듈 import 및 실행
if __name__ == '__main__':
    try:
        from final_server import app
        
        # Replit 환경 변수 확인
        port = int(os.environ.get('PORT', 8080))
        host = os.environ.get('HOST', '0.0.0.0')
        
        print("🚀 PDF ↔ DOCX 변환기 시작 (Replit 배포)")
        print(f"📍 서버 주소: {host}:{port}")
        
        # Flask 앱 실행
        app.run(
            host=host,
            port=port,
            debug=False,  # 배포 환경에서는 False
            threaded=True
        )
        
    except ImportError as e:
        print(f"❌ 모듈 import 오류: {e}")
        print("final_server.py 파일을 확인하세요.")
    except Exception as e:
        print(f"❌ 서버 시작 오류: {e}")