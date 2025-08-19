import os

def check_outputs_status():
    """outputs 폴더 상태 확인"""
    outputs_dir = "outputs"
    
    if not os.path.exists(outputs_dir):
        print(f"❌ {outputs_dir} 폴더가 존재하지 않습니다.")
        return
    
    files = os.listdir(outputs_dir)
    
    if not files:
        print("✅ outputs 폴더가 비어있습니다.")
    else:
        print(f"📁 outputs 폴더에 {len(files)}개 항목이 있습니다:")
        for item in files:
            item_path = os.path.join(outputs_dir, item)
            if os.path.isfile(item_path):
                size = os.path.getsize(item_path)
                print(f"  📄 {item} ({size:,} bytes)")
            else:
                print(f"  📁 {item} (폴더)")

if __name__ == "__main__":
    check_outputs_status()