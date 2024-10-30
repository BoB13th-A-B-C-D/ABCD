"""
본 모듈은 다음 파일들의 정보를 sqlite db의 형태로 바꾸는 것을 목적으로 한다.
# 1. 경로 - DFAS의 일괄 추출의 결과물인, *.csv
# 2. 경로 - 이벤트로그들 *.evtx
# 3. 경로 - 레지스트리 하이브 파일들
# 4. 경로 - 인쇄 관련 파일, *spl, *.shd
10/30 (수) 까지 완료
"""
import datetime
import os
import sqlite3
import struct
import xml.etree.ElementTree as ET
import pandas as pd
import win32evtlog
import win32evtlogutil
from datetime import datetime
from Registry import Registry

# 딕셔너리 정의
file_columns = {
    '프로그램_설치_프로그램.csv': ['프로그램명', '게시자', '설치된 시간', '설치경로'],
    '프로그램_실행_프로그램.csv': ['최종 실행시간', '프로그램명', '실행횟수', '사용자'],
    '파일_파일_폴더_접근.csv': ['열어본 시간', '파일명', '경로', '볼륨명', '열어본프로그램', '볼륨 S/N', '사용자명'],
    '파일_폴더_접근_link_파일_정보.csv': ['구분', '파일명', '경로', '크기', '생성시간', '수정시간', '접근시간', '대상 파일 생성시간', '대상 파일 수정시간', '대상 파일 접근시간', '볼륨 S/N', '사용자명'],
    '파일_점프리스트.csv': ['파일명', '경로', '링크이름', '프로그램', '생성시간', '수정시간', '접근시간', 'Entry CreationTime', 'Entry LastModified', '볼륨 S/N', '원본 파일명', '사용자명'],
    '파일_최근열람정보.csv': ['파일명', '경로', '유형'],
    '파일_휴지통.csv': ['삭제시간', '파일/폴더명', '경로', '볼륨', '크기', '사용자명'],
    '파일_shell_bag_mru정보.csv': ['방문시간', '폴더유형', '폴더', '유형', '생성시간', '수정시간', '접근시간', '사용자명'],
    '파일_시스템_볼륨_변경_이벤트($usnjrnl).csv': ['시간', 'USN', '파일명', '경로', '이벤트'],
    '파일_시스템_검색_및_색인_정보.csv': ['시간', '유형', '제목'],
    'activitiescache_정보_activity_정보.csv': ['마지막 수정 시간', '앱 ID', '앱 디스플레이명', '디스플레이 텍스트', '설명', '시작 시간', '종료 시간', '만료 시간'],
    '장치_시스템_on_off.csv': ['구분', '시간', '컴퓨터 이름'],
    '장치_저장장치_연결이력.csv': ['디바이스 정보', '시리얼넘버', '볼륨/볼륨명', 'Volume GUID', '볼륨 S/N', '최초연결시간 (SetupAPI)', '최초연결시간 (Registry)', 
                         '마지막 연결 시간', '연결 해제 시간', 'USB Stor Key', '디스크명'],
    '웹_브라우저_방문_웹사이트.csv': ['방문시간', '수집유형', '구분', '방문 사이트명', '세부내역', '제목', '사용자명'],
    '웹_브라우저_포털_검색어.csv': ['검색시간', '수집유형', '사이트', '구분', '검색어'],
    '웹_브라우저_웹_다운로드.csv': ['시작시간', '종료시간', '수집유형', '다운로드 파일명', 'URL', '상태', '크기', '사용자명'],
    '장치_os_정보.csv': ['운영체제', '버전', '빌드', '설치된 시간', '사용자', '표준시간대'],
    '장치_사용자_계정.csv': ['계정명', '로그인횟수', '비밀번호 여부', '계정 생성시간', '마지막 로그인 시간', '패스워드 변경시간', '마지막 로그인 실패시간'],
    '파일리스트_*.csv': ['파일명', '경로', '생성시간', '수정시간', '접근시간', 'MFT수정시간', '크기', '원본확장자', '변경확장자', '삭제', '해시셋', 'MD5', 'SHA1', '해시 태그']
}

def csv_to_db(path: str):
    # SQLite 데이터베이스 연결
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()

    # 디렉토리 내의 모든 파일 가져오기
    for file_name in os.listdir(path):
        if file_name in file_columns:  # 딕셔너리 키에 파일 이름이 있는지 확인
            file_path = os.path.join(path, file_name)
            columns = file_columns[file_name]  # 필요한 열 가져오기
            try:
                df = pd.read_csv(file_path, usecols=columns, encoding='utf-8', sep='\t')
                # 테이블 이름을 파일 이름에서 .csv 확장자를 제거한 형태로 지정
                table_name = file_name.replace('.csv', '')

                # 데이터베이스에 테이블로 저장
                df.to_sql(table_name, conn, if_exists='replace', index=False)
                #print(f"{table_name} 테이블이 생성되었습니다.")
            except Exception as e:
                print(f"{file_name} 파일을 읽는 중 오류 발생: {e}")

    # 데이터베이스 연결 종료
    conn.close()
    print("모든 파일이 처리되었습니다.")


def evtx_to_db_Diagnostic(evtx_path):
    # SQLite 데이터베이스 및 테이블 초기화
    db_path = "events.db"
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # 테이블 생성 (이벤트 ID, 생성 시간, 연결 상태를 저장)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS event_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER,
            time_generated TEXT,
            connection_status TEXT
        )
    ''')

    # 이벤트 로그 파일 핸들을 백업 로그에서 엽니다
    event_log = win32evtlog.OpenBackupEventLog(None, evtx_path)
    
    try:
        # 특정 이벤트 ID 설정
        target_event_id = 1006
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
        
        # 모든 이벤트를 읽기 위해 반복
        while True:
            events = win32evtlog.ReadEventLog(event_log, flags, 0)
            if not events:
                break  # 더 이상 이벤트가 없으면 루프 종료
            
            # 각 이벤트 검사 및 필터링
            for event in events:
                if event.EventID == target_event_id:
                    # StringInserts의 특정 인덱스에 있는 값을 UserRemovalPolicy로 간주
                    message_data = event.StringInserts
                    
                    connection_status = None
                    if message_data and len(message_data) > 3:
                        user_removal_policy = message_data[3].lower()
                        connection_status = 'Disconnected' if user_removal_policy == 'true' else 'Connected'
                       
                    # 연결 상태가 있으면 데이터베이스에 저장
                    if connection_status:
                        event_id = event.EventID
                        time_generated = event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S')
                        
                        cursor.execute('''
                            INSERT INTO event_log (event_id, time_generated, connection_status)
                            VALUES (?, ?, ?)
                        ''', (event_id, time_generated, connection_status))
                       
        
        # 변경사항 저장
        conn.commit()
        print(f"이벤트 로그가 '{db_path}' 데이터베이스에 저장되었습니다.")

    finally:
        # 데이터베이스 및 이벤트 로그 핸들 닫기
        win32evtlog.CloseEventLog(event_log)
        conn.close()


def evtx_to_db_PrintService(evtx_path):
    # SQLite 데이터베이스 및 테이블 초기화
    db_path = "driver_events.db"  # 별도의 데이터베이스 파일 사용
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # 테이블 생성 (이벤트 ID, 생성 시간, 상태, 추가 필드를 저장)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS driver_event_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER,
            time_generated TEXT,
            status TEXT,
            user TEXT,
            desktop TEXT,
            printer TEXT,
            port TEXT,
            size INTEGER,
            page INTEGER
        )
    ''')

    # 이벤트 로그 파일 핸들을 백업 로그에서 엽니다
    event_log = win32evtlog.OpenBackupEventLog(None, evtx_path)
    
    try:
        # 관심 있는 이벤트 ID 목록과 상태 매핑 (842 제외)
        event_status_mapping = {
            801: "Printing",
            802: "Deleted",
            307: "Completed",
            8421: "Network Connected",
            603: "Error"
        }
        
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
        
        # 모든 이벤트를 읽기 위해 반복
        while True:
            events = win32evtlog.ReadEventLog(event_log, flags, 0)
            if not events:
                break  # 더 이상 이벤트가 없으면 루프 종료
            
            # 각 이벤트 검사 및 필터링
            for event in events:
                if event.EventID in event_status_mapping:
                    # 이벤트 ID와 생성 시간, 상태 가져오기
                    event_id = event.EventID
                    time_generated = event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S')
                    status = event_status_mapping[event_id]

                    # 307 이벤트의 경우 StringInserts에서 추가 정보 추출
                    user = desktop = printer = port = None
                    size = page = None
                    if event_id == 307:
                        # StringInserts에서 직접 값 추출
                        message_data = event.StringInserts
                        if message_data and len(message_data) >= 8:
                            user = message_data[2]
                            desktop = message_data[3]
                            printer = message_data[4]
                            port = message_data[5]
                            size = int(message_data[6])
                            page = int(message_data[7])
                            
                    # 데이터베이스에 삽입
                    cursor.execute('''
                        INSERT INTO driver_event_log (
                            event_id, time_generated, status, user, desktop, printer, port, size, page
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (event_id, time_generated, status, user, desktop, printer, port, size, page))
            
        # 변경사항 저장
        conn.commit()
        print(f"드라이버 이벤트 로그가 '{db_path}' 데이터베이스에 저장되었습니다.")

    finally:
        # 데이터베이스 및 이벤트 로그 핸들 닫기
        win32evtlog.CloseEventLog(event_log)
        conn.close()


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


def spl_to_db(data, offset):
    """UTF-16LE 문자열 읽기"""
    if offset == 0 or offset >= len(data):
        return None
    try:
        end = offset
        while end + 1 < len(data):
            if data[end:end+2] == b'\x00\x00':
                break
            end += 2
        if end > offset:
            return data[offset:end].decode('utf-16le').strip()
    except Exception as e:
        print(f"Error reading UTF-16 string: {str(e)}")
    return None

def parse_shd(file_path):
    try:
        with open(file_path, 'rb') as f:
            data = f.read()

        # 기본 헤더 정보
        signature = struct.unpack('>I', data[0:4])[0]  # big-endian으로 읽기
        header_size = struct.unpack('<I', data[4:8])[0]  # little-endian
        status = struct.unpack('<H', data[8:10])[0]
        job_id = struct.unpack('<I', data[12:16])[0]
        priority = struct.unpack('<I', data[16:20])[0]

        # 오프셋 읽기 (각 4바이트)
        username_offset = struct.unpack('<I', data[20:24])[0]
        notify_name_offset = struct.unpack('<I', data[24:28])[0]
        document_name_offset = struct.unpack('<I', data[40:44])[0]
        printer_name_offset = struct.unpack('<I', data[56:60])[0]

        # 문자열 데이터 읽기
        result = {
            'file_name': os.path.basename(file_path),
            'creation_time': datetime.fromtimestamp(os.path.getctime(file_path)),  # 파일 생성 시간
            'signature': hex(signature),
            'header_size': header_size,
            'status': hex(status),
            'job_id': job_id,
            'priority': priority,
            'strings': {
                'user_name': spl_to_db(data, username_offset),
                'notify_name': spl_to_db(data, notify_name_offset),
                'document_name': spl_to_db(data, document_name_offset),
                'printer_name': spl_to_db(data, printer_name_offset),
            }
        }

        return result

    except Exception as e:
        print(f"Error parsing SHD file {os.path.basename(file_path)}: {str(e)}")
        return None

def initialize_db(db_path):
    """SQLite 데이터베이스 초기화 및 테이블 생성"""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            creation_time TEXT,
            notify_name TEXT,
            document_name TEXT,
            printer_name TEXT
        )
    ''')
    conn.commit()
    return conn

def insert_record(conn, creation_time, notify_name, document_name, printer_name):
    """데이터베이스에 기록 추가"""
    cursor = conn.cursor()
    try:
        cursor.execute('''
            INSERT INTO documents (creation_time, notify_name, document_name, printer_name)
            VALUES (?, ?, ?, ?)
        ''', (creation_time, notify_name, document_name, printer_name))
        conn.commit()
    except Exception as e:
        print(f"Error inserting record into database: {str(e)}")

def main():
    print("=== SHD File Analyzer ===")
    path = input("\nEnter path to SHD file or directory: ").strip()
    
    if not os.path.exists(path):
        print("Error: Path does not exist!")
        return

    db_path = 'shd_documents.db'  # SQLite 데이터베이스 파일 경로
    conn = initialize_db(db_path)  # 데이터베이스 초기화

    def analyze_file(file_path):
        print(f"\nAnalyzing: {os.path.basename(file_path)}")
        result = parse_shd(file_path)
        if result:
            strings = result['strings']
            creation_time = result['creation_time'].isoformat()  # ISO 포맷으로 변환
            print("\nDocument Information:")
            print(f"File Name: {result['file_name']}")
            print(f"Creation Time: {creation_time}")
            if strings['notify_name']:
                print(f"NOTIFY Name: {strings['notify_name']}")
            if strings['document_name']:
                print(f"DOCUMENT Name: {strings['document_name']}")
            if strings['printer_name']:
                print(f"Printer Name: {strings['printer_name']}")

            # 데이터베이스에 기록 추가
            insert_record(conn, creation_time, strings['notify_name'], strings['document_name'], strings['printer_name'])

    if os.path.isfile(path):
        if path.lower().endswith('.shd'):
            analyze_file(path)
        else:
            print("Not a SHD file!")
    else:
        found_files = False
        for root, _, files in os.walk(path):
            for file in files:
                if file.lower().endswith('.shd'):
                    found_files = True
                    analyze_file(os.path.join(root, file))
        
        if not found_files:
            print("No SHD files found in the specified directory.")

    conn.close()  # 데이터베이스 연결 종료