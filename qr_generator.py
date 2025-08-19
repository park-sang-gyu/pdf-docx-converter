# 새 파일 생성
import qrcode
from PIL import Image
import json

def generate_document_qr(document_data, output_path):
    """문서 데이터를 QR코드로 생성"""
    try:
        # QR코드에 포함할 데이터
        qr_data = {
            'doc_id': document_data.get('document_id'),
            'kc_num': document_data.get('kc_number'),
            'reg_num': document_data.get('registration_number'),
            'date': document_data.get('date'),
            'url': f"https://your-domain.com/verify/{document_data.get('document_id')}"
        }
        
        # QR코드 생성
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        
        qr.add_data(json.dumps(qr_data, ensure_ascii=False))
        qr.make(fit=True)
        
        # 이미지 생성
        qr_image = qr.make_image(fill_color="black", back_color="white")
        qr_image.save(output_path)
        
        print(f"✅ QR코드 생성: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ QR코드 생성 오류: {e}")
        return False