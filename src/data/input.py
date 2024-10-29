"""
본 모듈은 다음 파일들의 정보를 sqlite db의 형태로 바꾸는 것을 목적으로 한다.
1. 경로 - DFAS의 일괄 추출의 결과물인, *.csv
2. 경로 - 이벤트로그들 *.evtx
3. 경로 - 레지스트리 하이브 파일들
4. 경로 - 인쇄 관련 파일, *spl, *.shd
5. 파일 - 윈도우 타임라인 파일, Activitycache.db
10/30 (수) 까지 완료
"""

import pandas as pd
from Registry import Registry
import sqlite3
from datetime import datetime
import os

def csv_to_db(filename: str):
    """
    파일명에 따라 지정된 칼럼을 생성하여 CSV 데이터를 데이터베이스에 삽입
    """
    # 파일명에 따른 칼럼 설정
    dic = {
        '프로그램_설치_프로그램.csv': ['프로그램명', '게시자', '설치된 시간', '설치경로'],
        '프로그램_실행_프로그램.csv': ['최종 실행시간', '프로그램명', '실행횟수', '사용자'],
        '파일_파일_폴더_접근.csv': ['열어본 시간', '파일명', '경로', '볼륨명', '열어본프로그램', '볼륨 S/N', '사용자명'],
        '파일_폴더_접근_link_파일_정보.csv': ['구분', '파일명', '경로', '크기', '생성시간', '수정시간', '접근시간', 
                                        '대상 파일 생성시간', '대상 파일 수정시간', '대상파일 접근시간', '볼륨 S/N', '사용자명'],
        '파일_점프리스트.csv': ['파일명', '경로', '링크이름', '프로그램', '생성시간', '수정시간', '접근시간', 
                             'Entry CreationTime', 'Entry LastModified', '볼륨 S/N', '원본파일명', '사용자명'],
        '파일_최근열람정보.csv': ['파일명', '경로', '유형'],
        '파일_휴지통.csv': ['삭제시간', '파일/폴더명', '경로', '볼륨', '크기', '사용자명'],
        '파일_shell_bag_mru정보.csv': ['방문시간', '폴더유형', '폴더', '유형', '생성시간', '수정시간', '접근시간', '사용자명'],
        '파일_시스템_볼륨_변경_이벤트($usnjrnl).csv': ['시간', 'USN', '파일명', '경로', '이벤트'],
        '파일_시스템_검색_및_색인_정보.csv': ['시간', '유형', '제목'],
        'activitiescache_정보_activity_정보.csv': ['마지막 수정 시간', '앱 ID', '앱 디스플레이명', '디스플레이 텍스트', '설명', '시작', '종료', 
                                                 '마지막 수정 시간', '만료 시간'],
        '장치_시스템_on_off.csv': ['구분', '시간', '컴퓨터 이름'],
        '장치_저장장치_연결이력.csv': ['연결시간', '장치명', '경로', '사용자'],
        '웹_브라우저_방문_웹사이트.csv': ['방문시간', '수집유형', '구분', '방문 사이트명', '세부내역', '제목', '사용자명'],
        '웹_브라우저_포털_검색어.csv': ['검색시간', '수집유형', '사이트', '구분', '검색어'],
        '웹_브라우저_웹_다운로드.csv': ['시작시간', '종료시간', '수집유형', '다운로드 파일명', 'URL', '상태', '크기', '사용자명'],
        '장치_os_정보.csv': ['운영체제', '버전', '빌드', '설치된 시간', '사용자', '표준시간대'],
        '장치_사용자_계정.csv': ['계정명', '로그인횟수', '비밀번호 여부', '계정 생성시간', '마지막 로그인 시간', '패스워드 변경시간', '마지막 로그인 실패시간'],
    }

    # 지정된 파일명이 dic에 있는지 확인하고 해당 칼럼을 할당
    if filename not in dic:
        print(f"파일명 '{filename}'에 대한 칼럼이 정의되어 있지 않습니다.")
        return

    # 파일에서 데이터 로드
    columns = dic[filename]
    try:
        data = pd.read_csv(filename, names=columns)
    except Exception as e:
        print(f"CSV 파일을 로드하는 중 오류가 발생했습니다: {e}")
        return

    # 데이터베이스 연결 및 테이블 생성
    conn = sqlite3.connect('my_database.db')
    cursor = conn.cursor()
    
    table_name = filename.split('.')[0]
    column_definitions = ', '.join([f"{col} TEXT" for col in columns])  # 모든 칼럼을 TEXT로 설정
    cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({column_definitions})")
    
    # 데이터베이스에 데이터 삽입
    for _, row in data.iterrows():
        placeholders = ', '.join(['?'] * len(columns))
        sql = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({placeholders})"
        cursor.execute(sql, tuple(row))
    
    conn.commit()
    conn.close()
    print(f"{filename}의 데이터를 데이터베이스에 성공적으로 삽입했습니다.")


def evtx_to_db():
    #
    pass

def reg_to_db():
    """Registry files to SQLite database converter"""
    print("\n=== Registry to Database Converter ===")
    print("\n=== Registry Path Input ===")
    print("Please enter the path to the registry files directory.")
    
    path = input("\nRegistry path: ").strip()
    
    if path.lower() == 'q':
        print("Registry analysis skipped.")
        return
        
    # 경로 검증
    if not os.path.exists(path):
        print("\nError: Path does not exist!")
        return
        
    # SOFTWARE 파일 존재 확인
    software_path = os.path.join(path, "SOFTWARE")
    if not os.path.exists(software_path):
        print("\nError: SOFTWARE registry hive not found in the specified path!")
        return
    
    try:
        print(f"\nAnalyzing registry files from: {path}")
        analyzer = OfflinePrinterAnalyzer(path)
        
        print("\nAnalyzing offline registry files...")
        analyzer.update_database()
        
        print("\nDisplaying saved printer information...")
        analyzer.display_printer_info()
        
        print("\nRegistry analysis completed.")
        
    except Exception as e:
        print(f"\nError during registry analysis: {str(e)}")
        
# 지원 클래스 정의
class OfflinePrinterAnalyzer:
    def __init__(self, registry_path, db_path="printer_info.db"):
        """Initialize offline registry analyzer"""
        self.registry_path = registry_path
        self.db_path = db_path
        
        # 데이터베이스 파일이 존재하면 삭제
        if os.path.exists(self.db_path):
            os.remove(self.db_path)
            
        self.create_database()
    
    def create_database(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS printers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                printer_name TEXT,
                port TEXT,
                driver TEXT,
                print_processor TEXT,
                registry_path TEXT,
                is_default INTEGER DEFAULT 0,
                default_printer_data TEXT,
                collection_time TIMESTAMP,
                case_path TEXT
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def parse_installed_printers(self):
        software_path = os.path.join(self.registry_path, "SOFTWARE")
        printers = []
        
        try:
            registry = Registry.Registry(software_path)
            try:
                key_path = "Microsoft\\Windows NT\\CurrentVersion\\Print\\Printers"
                printers_key = registry.open(key_path)
                
                for subkey in printers_key.subkeys():
                    printer_info = {
                        'printer_name': subkey.name(),
                        'port': '',
                        'driver': '',
                        'print_processor': '',
                        'registry_path': software_path
                    }
                    
                    for value in subkey.values():
                        if value.name() == "Port":
                            printer_info['port'] = value.value()
                        elif value.name() == "Printer Driver":
                            printer_info['driver'] = value.value()
                        elif value.name() == "Print Processor":
                            printer_info['print_processor'] = value.value()
                    
                    printers.append(printer_info)
                print(f"Found {len(printers)} printers in registry")
                    
            except Registry.RegistryKeyNotFoundException:
                print(f"Printers key not found in SOFTWARE hive at {software_path}")
                
        except Exception as e:
            print(f"Error accessing SOFTWARE hive: {str(e)}")
            
        return printers
    
    def parse_default_printer(self):
        """Parse default printer settings from SOFTWARE hive"""
        software_path = os.path.join(self.registry_path, "SOFTWARE")
        default_printers = {}
        
        try:
            registry = Registry.Registry(software_path)
            try:
                possible_paths = [
                    "Microsoft\\Windows NT\\CurrentVersion\\Windows\\Devices",
                    "Microsoft\\Windows NT\\CurrentVersion\\Devices",
                    "Microsoft\\Windows NT\\CurrentVersion\\PrinterPorts"
                ]
                
                for path in possible_paths:
                    try:
                        devices_key = registry.open(path)
                        print(f"Found printer settings in: {path}")
                        
                        for value in devices_key.values():
                            printer_name = value.name()
                            printer_data = value.value()
                            default_printers[printer_name] = printer_data
                            print(f"Found printer: {printer_name} = {printer_data}")
                        
                        if default_printers:
                            break
                            
                    except Registry.RegistryKeyNotFoundException:
                        continue
                
                if not default_printers:
                    print("No printer settings found in any expected registry paths")
                    
            except Registry.RegistryKeyNotFoundException:
                print(f"No printer settings found in SOFTWARE hive at {software_path}")
                
        except Exception as e:
            print(f"Error accessing SOFTWARE hive for default printer: {str(e)}")
            
        return default_printers
    
    def update_database(self):
        """Update database with parsed printer information"""
        current_time = datetime.now().isoformat()
        
        installed_printers = self.parse_installed_printers()
        default_printers = self.parse_default_printer()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        for printer in installed_printers:
            is_default = 0
            default_printer_data = ""
            
            if printer['printer_name'] in default_printers:
                is_default = 1
                default_printer_data = default_printers[printer['printer_name']]
            
            cursor.execute('''
                INSERT INTO printers 
                (printer_name, port, driver, print_processor, registry_path,
                 is_default, default_printer_data, collection_time, case_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                printer['printer_name'],
                printer['port'],
                printer['driver'],
                printer['print_processor'],
                printer['registry_path'],
                is_default,
                default_printer_data,
                current_time,
                self.registry_path
            ))
        
        conn.commit()
        conn.close()
    
    def display_printer_info(self):
        """Display printer information from database"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT printer_name, port, driver, print_processor, is_default, 
                   default_printer_data, collection_time 
            FROM printers 
            WHERE case_path = ?
            ORDER BY is_default DESC, printer_name
        ''', (self.registry_path,))
        
        rows = cursor.fetchall()
        
        print(f"\n=== Printer Information from Registry ===")
        print(f"Registry Path: {self.registry_path}")
        print("-" * 50)
        
        if not rows:
            print("No printer information found in database.")
            return
            
        for row in rows:
            print(f"Printer Name: {row[0]}")
            print(f"Port: {row[1]}")
            print(f"Driver: {row[2]}")
            print(f"Print Processor: {row[3]}")
            print(f"Default Printer: {'Yes' if row[4] else 'No'}")
            if row[4]:
                print(f"Default Printer Data: {row[5]}")
            print(f"Collection Time: {row[6]}")
            print("-" * 50)
        
        conn.close()

if __name__ == "__main__":
    reg_to_db()
    pass

def spl_to_db():
    #
    pass

def activitycache_to_db():
    #
    pass