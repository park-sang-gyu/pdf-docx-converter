import os
import shutil
from datetime import datetime

def safe_cleanup_outputs():
    """안전한 outputs 폴더 정리 (백업 옵션 포함)"""
    outputs_dir = "outputs"
    backup_dir = f"outputs_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(outputs_dir):
        print(f"❌ {outputs_dir} 폴더가 존재하지 않습니다.")
        return
    
    files = [f for f in os.listdir(outputs_dir) if os.path.isfile(os.path.join(outputs_dir, f))]
    
    if not files:
        print("✅ outputs 폴더가 이미 비어있습니다.")
        return
    
    print(f"📁 발견된 파일: {len(files)}개")
    for f in files:
        print(f"  - {f}")
    
    # 사용자 확인
    response = input("\n🗑️ 정말로 모든 파일을 삭제하시겠습니까? (y/N): ")
    
    if response.lower() not in ['y', 'yes']:
        print("❌ 삭제가 취소되었습니다.")
        return
    
    # 백업 생성 옵션
    backup_response = input("💾 삭제 전 백업을 생성하시겠습니까? (Y/n): ")
    
    if backup_response.lower() not in ['n', 'no']:
        try:
            shutil.copytree(outputs_dir, backup_dir)
            print(f"✅ 백업 생성됨: {backup_dir}")
        except Exception as e:
            print(f"⚠️ 백업 생성 실패: {e}")
    
    # 파일 삭제
    deleted_count = 0
    for filename in files:
        file_path = os.path.join(outputs_dir, filename)
        try:
            os.remove(file_path)
            print(f"  ✅ 삭제됨: {filename}")
            deleted_count += 1
        except Exception as e:
            print(f"  ❌ 삭제 실패: {filename} - {e}")
    
    print(f"\n🎉 {deleted_count}개 파일이 삭제되었습니다!")

if __name__ == "__main__":
    safe_cleanup_outputs()