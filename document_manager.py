import json
import csv
import sqlite3
import os
from datetime import datetime, date
from typing import Dict, List, Optional

class DocumentManager:
    def __init__(self, data_dir="document_data"):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)
        
        # 파일 경로들
        self.db_file = os.path.join(data_dir, "documents.db")
        self.json_file = os.path.join(data_dir, "documents.json")
        self.csv_file = os.path.join(data_dir, "documents.csv")
        
        # 데이터베이스 초기화
        self.init_database()
    
    def init_database(self):
        """데이터베이스 초기화 및 테이블 생성"""
        try:
            # SQL 스키마 파일 읽기
            schema_file = os.path.join(os.path.dirname(__file__), "documents.sql")
            
            if os.path.exists(schema_file):
                with open(schema_file, 'r', encoding='utf-8') as f:
                    schema_sql = f.read()
                
                with sqlite3.connect(self.db_file) as conn:
                    # 여러 SQL 문을 실행
                    conn.executescript(schema_sql)
                    conn.commit()
                    
                print(f"✅ 데이터베이스 초기화 완료: {self.db_file}")
            else:
                print(f"⚠️ 스키마 파일 없음: {schema_file}")
                self._create_basic_tables()
                
        except Exception as e:
            print(f"❌ 데이터베이스 초기화 오류: {e}")
            self._create_basic_tables()
    
    def _create_basic_tables(self):
        """기본 테이블 생성 (폴백)"""
        try:
            with sqlite3.connect(self.db_file) as conn:
                conn.execute('''
                    CREATE TABLE IF NOT EXISTS documents (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        filename VARCHAR(255) NOT NULL,
                        conversion_method VARCHAR(50),
                        success BOOLEAN DEFAULT FALSE,
                        kc_number VARCHAR(100),
                        registration_number VARCHAR(100),
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                conn.commit()
                print("✅ 기본 테이블 생성 완료")
        except Exception as e:
            print(f"❌ 기본 테이블 생성 오류: {e}")
    
    def save_document_data(self, pdf_path: str, extracted_numbers: Dict, 
                          conversion_method: str, success: bool = True, 
                          processing_time: float = 0.0) -> int:
        """문서 데이터 저장 (DB + JSON + CSV)"""
        
        document_data = {
            'timestamp': datetime.now().isoformat(),
            'pdf_path': pdf_path,
            'filename': os.path.basename(pdf_path),
            'conversion_method': conversion_method,
            'success': success,
            'extracted_numbers': extracted_numbers,
            'file_size': os.path.getsize(pdf_path) if os.path.exists(pdf_path) else 0,
            'processing_time': processing_time
        }
        
        # 1. 데이터베이스 저장
        document_id = self._save_to_database(document_data)
        
        # 2. JSON 저장 (백업용)
        self._save_to_json(document_data)
        
        # 3. CSV 저장 (Excel 호환)
        self._save_to_csv(document_data)
        
        # 4. 실패 케이스 별도 처리
        if not success:
            self._save_failed_case(document_id, "Conversion failed")
        
        # 5. 통계 업데이트
        self._update_daily_stats(success, conversion_method, processing_time)
        
        return document_id
    
    def _save_to_database(self, document_data: Dict) -> int:
        """SQLite 데이터베이스에 저장"""
        try:
            with sqlite3.connect(self.db_file) as conn:
                cursor = conn.cursor()
                
                # 문서 데이터 삽입
                cursor.execute('''
                    INSERT INTO documents (
                        filename, original_path, conversion_method, success,
                        kc_number, registration_number, document_number, 
                        business_number, phone_number, file_size, processing_time_seconds
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    document_data['filename'],
                    document_data['pdf_path'],
                    document_data['conversion_method'],
                    document_data['success'],
                    document_data['extracted_numbers'].get('kc_number'),
                    document_data['extracted_numbers'].get('registration_number'),
                    document_data['extracted_numbers'].get('document_number'),
                    document_data['extracted_numbers'].get('business_number'),
                    document_data['extracted_numbers'].get('phone_number'),
                    document_data['file_size'],
                    document_data['processing_time']
                ))
                
                document_id = cursor.lastrowid
                conn.commit()
                
                print(f"💾 DB 저장 완료: ID {document_id}")
                return document_id
                
        except Exception as e:
            print(f"❌ DB 저장 오류: {e}")
            return -1
    
    def _save_failed_case(self, document_id: int, failure_reason: str):
        """실패 케이스 저장"""
        try:
            with sqlite3.connect(self.db_file) as conn:
                conn.execute('''
                    INSERT INTO extraction_failures (
                        document_id, failure_reason, failure_type
                    ) VALUES (?, ?, ?)
                ''', (document_id, failure_reason, 'conversion_failure'))
                conn.commit()
                
        except Exception as e:
            print(f"❌ 실패 케이스 저장 오류: {e}")
    
    def _update_daily_stats(self, success: bool, method: str, processing_time: float):
        """일일 통계 업데이트"""
        try:
            today = date.today()
            
            with sqlite3.connect(self.db_file) as conn:
                # 오늘 통계 확인
                cursor = conn.cursor()
                cursor.execute('SELECT id FROM conversion_stats WHERE date = ?', (today,))
                
                if cursor.fetchone():
                    # 기존 통계 업데이트
                    conn.execute('''
                        UPDATE conversion_stats SET 
                            total_conversions = total_conversions + 1,
                            successful_conversions = successful_conversions + ?,
                            text_based_conversions = text_based_conversions + ?,
                            ocr_based_conversions = ocr_based_conversions + ?,
                            avg_processing_time = (avg_processing_time + ?) / 2
                        WHERE date = ?
                    ''', (
                        1 if success else 0,
                        1 if method == 'text' else 0,
                        1 if method == 'ocr' else 0,
                        processing_time,
                        today
                    ))
                else:
                    # 새 통계 생성
                    conn.execute('''
                        INSERT INTO conversion_stats (
                            date, total_conversions, successful_conversions,
                            text_based_conversions, ocr_based_conversions, avg_processing_time
                        ) VALUES (?, 1, ?, ?, ?, ?)
                    ''', (
                        today,
                        1 if success else 0,
                        1 if method == 'text' else 0,
                        1 if method == 'ocr' else 0,
                        processing_time
                    ))
                
                conn.commit()
                
        except Exception as e:
            print(f"❌ 통계 업데이트 오류: {e}")
    
    def get_failed_documents(self) -> List[Dict]:
        """검수가 필요한 실패 문서 목록 조회"""
        try:
            with sqlite3.connect(self.db_file) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                
                cursor.execute('''
                    SELECT d.*, ef.failure_reason, ef.manual_review_status
                    FROM documents d
                    JOIN extraction_failures ef ON d.id = ef.document_id
                    WHERE ef.manual_review_status = 'pending'
                    ORDER BY d.created_at DESC
                ''')
                
                return [dict(row) for row in cursor.fetchall()]
                
        except Exception as e:
            print(f"❌ 실패 문서 조회 오류: {e}")
            return []
    
    def get_daily_stats(self, days: int = 7) -> List[Dict]:
        """최근 N일 통계 조회"""
        try:
            with sqlite3.connect(self.db_file) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                
                cursor.execute('''
                    SELECT * FROM conversion_stats 
                    WHERE date >= date('now', '-{} days')
                    ORDER BY date DESC
                '''.format(days))
                
                return [dict(row) for row in cursor.fetchall()]
                
        except Exception as e:
            print(f"❌ 통계 조회 오류: {e}")
            return []
    
    def _save_to_json(self, document_data: Dict):
        """JSON 백업 저장"""
        try:
            if os.path.exists(self.json_file):
                with open(self.json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = []
            
            data.append(document_data)
            
            with open(self.json_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"❌ JSON 저장 오류: {e}")
    
    def _save_to_csv(self, document_data: Dict):
        """CSV 저장 (Excel 호환)"""
        try:
            file_exists = os.path.exists(self.csv_file)
            
            with open(self.csv_file, 'a', newline='', encoding='utf-8-sig') as f:
                fieldnames = ['timestamp', 'filename', 'conversion_method', 'success', 
                             'kc_number', 'registration_number', 'document_number', 
                             'business_number', 'processing_time']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                
                if not file_exists:
                    writer.writeheader()
                
                csv_row = {
                    'timestamp': document_data['timestamp'],
                    'filename': document_data['filename'],
                    'conversion_method': document_data['conversion_method'],
                    'success': document_data['success'],
                    'processing_time': document_data['processing_time']
                }
                
                # 번호 필드들 추가
                for key in ['kc_number', 'registration_number', 'document_number', 'business_number']:
                    csv_row[key] = document_data['extracted_numbers'].get(key, '')
                
                writer.writerow(csv_row)
                
        except Exception as e:
            print(f"❌ CSV 저장 오류: {e}")