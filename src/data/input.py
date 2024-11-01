"""
본 모듈은 다음 파일들의 정보를 sqlite db의 형태로 바꾸는 것을 목적으로 한다.
# 1. 경로 - DFAS의 일괄 추출의 결과물인, *.csv
# 2. 경로 - 이벤트로그들 *.evtx
# 3. 경로 - 레지스트리 하이브 파일들
# 4. 경로 - 인쇄 관련 파일, *spl, *.shd
10/30 (수) 까지 완료
"""
import datetime
import glob
import os
import sqlite3
import struct
import xml.etree.ElementTree as ET
import pandas as pd
import win32evtlog
import win32evtlogutil
from datetime import datetime
from Registry import Registry

db_name = 'database.db'

# 딕셔너리 정의 (한글 파일명/컬럼명을 영어 테이블명/컬럼명으로 매핑)
file_columns = {
    '프로그램_설치_프로그램.csv': {
        'table_name': 'installed_programs',
        'columns': {
            '프로그램명': 'program_name',
            '게시자': 'publisher',
            '설치된 시간': 'installed_time',
            '설치경로': 'installation_path'
        }
    },
    '프로그램_실행_프로그램.csv': {
        'table_name': 'program_execution',
        'columns': {
            '최종 실행시간': 'last_execution_time',
            '프로그램명': 'program_name',
            '실행횟수': 'execution_count',
            '사용자': 'user'
        }
    },
    '파일_파일_폴더_접근.csv': {
        'table_name': 'file_folder_access',
        'columns': {
            '열어본 시간': 'access_time',
            '파일명': 'file_name',
            '경로': 'path',
            '볼륨명': 'volume_name',
            '열어본프로그램': 'accessed_program',
            '볼륨 S/N': 'volume_serial_number',
            '사용자명': 'user_name'
        }
    },
    '파일_폴더_접근_link_파일_정보.csv': {
        'table_name': 'folder_link_file_info',
        'columns': {
            '구분': 'type',
            '파일명': 'file_name',
            '경로': 'path',
            '크기': 'size',
            '생성시간': 'creation_time',
            '수정시간': 'modification_time',
            '접근시간': 'access_time',
            '대상 파일 생성시간': 'target_file_creation_time',
            '대상 파일 수정시간': 'target_file_modification_time',
            '대상 파일 접근시간': 'target_file_access_time',
            '볼륨 S/N': 'volume_serial_number',
            '사용자명': 'user_name'
        }
    },
    '파일_점프리스트.csv': {
        'table_name': 'file_jumplist',
        'columns': {
            '파일명': 'file_name',
            '경로': 'path',
            '링크이름': 'link_name',
            '프로그램': 'program',
            '생성시간': 'creation_time',
            '수정시간': 'modification_time',
            '접근시간': 'access_time',
            'Entry CreationTime': 'entry_creation_time',
            'Entry LastModified': 'entry_last_modified',
            '볼륨 S/N': 'volume_serial_number',
            '원본 파일명': 'original_file_name',
            '사용자명': 'user_name'
        }
    },
    '파일_최근열람정보.csv': {
        'table_name': 'file_recent_view_info',
        'columns': {
            '파일명': 'file_name',
            '경로': 'path',
            '유형': 'type'
        }
    },
    '파일_휴지통.csv': {
        'table_name': 'file_recycle_bin',
        'columns': {
            '삭제시간': 'deletion_time',
            '파일/폴더명': 'file_folder_name',
            '경로': 'path',
            '볼륨': 'volume',
            '크기': 'size',
            '사용자명': 'user_name'
        }
    },
    '파일_shell_bag_mru정보.csv': {
        'table_name': 'file_shell_bag_mru_info',
        'columns': {
            '방문시간': 'visit_time',
            '폴더유형': 'folder_type',
            '폴더': 'folder',
            '유형': 'type',
            '생성시간': 'creation_time',
            '수정시간': 'modification_time',
            '접근시간': 'access_time',
            '사용자명': 'user_name'
        }
    },
    '파일_시스템_볼륨_변경_이벤트($usnjrnl).csv': {
        'table_name': 'file_system_volume_change_event',
        'columns': {
            '시간': 'time',
            'USN': 'usn',
            '파일명': 'file_name',
            '경로': 'path',
            '이벤트': 'event'
        }
    },
    '파일_시스템_검색_및_색인_정보.csv': {
        'table_name': 'file_system_search_index_info',
        'columns': {
            '시간': 'time',
            '유형': 'type',
            '제목': 'title'
        }
    },
    'activitiescache_정보_activity_정보.csv': {
        'table_name': 'activities_cache_activity_info',
        'columns': {
            '마지막 수정 시간': 'last_modified_time',
            '앱 ID': 'app_id',
            '앱 디스플레이명': 'app_display_name',
            '디스플레이 텍스트': 'display_text',
            '설명': 'description',
            '시작 시간': 'start_time',
            '종료 시간': 'end_time',
            '만료 시간': 'expiration_time'
        }
    },
    '장치_시스템_on_off.csv': {
        'table_name': 'device_system_on_off',
        'columns': {
            '구분': 'type',
            '시간': 'time',
            '컴퓨터 이름': 'computer_name'
        }
    },
    '장치_저장장치_연결이력.csv': {
        'table_name': 'device_storage_connection_history',
        'columns': {
            '디바이스 정보': 'device_info',
            '시리얼넘버': 'serial_number',
            '볼륨/볼륨명': 'volume_name',
            'Volume GUID': 'volume_guid',
            '볼륨 S/N': 'volume_serial_number',
            '최초연결시간 (SetupAPI)': 'first_connection_time_setupapi',
            '최초연결시간 (Registry)': 'first_connection_time_registry',
            '마지막 연결 시간': 'last_connection_time',
            '연결 해제 시간': 'disconnection_time',
            'USB Stor Key': 'usb_stor_key',
            '디스크명': 'disk_name'
        }
    },
    '웹_브라우저_방문_웹사이트.csv': {
        'table_name': 'web_browser_visited_websites',
        'columns': {
            '방문시간': 'visit_time',
            '수집유형': 'collection_type',
            '구분': 'type',
            '방문 사이트명': 'website_name',
            '세부내역': 'details',
            '제목': 'title',
            '사용자명': 'user_name'
        }
    },
    '웹_브라우저_포털_검색어.csv': {
        'table_name': 'web_browser_portal_search_terms',
        'columns': {
            '검색시간': 'search_time',
            '수집유형': 'collection_type',
            '사이트': 'site',
            '구분': 'type',
            '검색어': 'search_term'
        }
    },
    '웹_브라우저_웹_다운로드.csv': {
        'table_name': 'web_browser_web_download',
        'columns': {
            '시작시간': 'start_time',
            '종료시간': 'end_time',
            '수집유형': 'collection_type',
            '다운로드 파일명': 'downloaded_file_name',
            'URL': 'url',
            '상태': 'status',
            '크기': 'size',
            '사용자명': 'user_name'
        }
    },
    '장치_os_정보.csv': {
        'table_name': 'device_os_info',
        'columns': {
            '운영체제': 'operating_system',
            '버전': 'version',
            '빌드': 'build',
            '설치된 시간': 'installation_time',
            '사용자': 'user',
            '표준시간대': 'timezone'
        }
    },
    '장치_사용자_계정.csv': {
        'table_name': 'device_user_account',
        'columns': {
            '계정명': 'account_name',
            '로그인횟수': 'login_count',
            '비밀번호 여부': 'password_set',
            '계정 생성시간': 'account_creation_time',
            '마지막 로그인 시간': 'last_login_time',
            '패스워드 변경시간': 'password_change_time',
            '마지막 로그인 실패시간': 'last_login_failure_time'
        }
    },
    '파일리스트_*.csv': {
        'table_name': 'file_list',
        'columns': {
            '파일명': 'file_name',
            '경로': 'path',
            '생성시간': 'creation_time',
            '수정시간': 'modification_time',
            '접근시간': 'access_time',
            'MFT수정시간': 'MFT_modification_time',
            '크기': 'size',
            '원본확장자': 'original_extension',
            '변경확장자': 'modified_extension',
            '삭제': 'deleted',
            '해시셋': 'hash_set',
            'MD5': 'MD5',
            'SHA1': 'SHA1',
            '해시 태그': 'hash_tag'
        }
    }
}

def csv_to_db(path: str):
    # SQLite 데이터베이스 연결
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    # 디렉토리 내의 모든 파일 가져오기
    for file_name in os.listdir(path):
        if file_name in file_columns:
            file_path = os.path.join(path, file_name)
            table_info = file_columns[file_name]
            table_name = table_info['table_name']
            columns_map = table_info['columns']

            try:
                df = pd.read_csv(file_path, usecols=columns_map.keys(), encoding='utf-8', sep='\t')
                df.rename(columns=columns_map, inplace=True)
                df.to_sql(table_name, conn, if_exists='replace', index=False)
                print(f"{table_name} 테이블이 생성되었습니다.")
            except Exception as e:
                print(f"{file_name} 파일을 읽는 중 오류 발생: {e}")

    # '파일리스트_'로 시작하는 파일들을 glob 패턴을 통해 가져오기
    for file_path in glob.glob(os.path.join(path, '파일리스트_*.csv')):
        table_info = file_columns['파일리스트_*.csv']
        table_name = table_info['table_name']
        columns_map = table_info['columns']

        try:
            df = pd.read_csv(file_path, usecols=columns_map.keys(), encoding='utf-8', sep='\t')
            df.rename(columns=columns_map, inplace=True)
            df.to_sql(table_name, conn, if_exists='replace', index=False)
            print(f"{table_name} 테이블이 생성되었습니다.")
        except Exception as e:
            print(f"{file_path} 파일을 읽는 중 오류 발생: {e}")

    conn.close()
    print("모든 파일이 처리되었습니다.")


def evtx_to_db_Diagnostic(evtx_path):
    # SQLite 데이터베이스 및 테이블 초기화
    conn = sqlite3.connect(db_name)
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
        print(f"이벤트 로그가 '{db_name}' 데이터베이스에 저장되었습니다.")

    finally:
        # 데이터베이스 및 이벤트 로그 핸들 닫기
        win32evtlog.CloseEventLog(event_log)
        conn.close()


def evtx_to_db_PrintService(evtx_path):
    # SQLite 데이터베이스 및 테이블 초기화
    conn = sqlite3.connect(db_name)
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
        print(f"드라이버 이벤트 로그가 '{db_name}' 데이터베이스에 저장되었습니다.")

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
    def __init__(self, registry_path, db_path=db_name):
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

def initialize_db(db_name):
    """SQLite 데이터베이스 초기화 및 테이블 생성"""
    conn = sqlite3.connect(db_name)
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

def shd_to_db(path: str):
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

csv_to_db(r'C:\Users\soke0\Desktop\DFAS')
evtx = r'D:\DFAS_PRO_X64_1.3.0.5_004\DFASPro\Evidence\CASE_202410291429\CollectSourceFile\20241029143259\Partition3\Windows\System32\winevt\Logs'
evtx_to_db_Diagnostic(evtx + '\\' + 'Microsoft-Windows-Partition%4Diagnostic.evtx')
evtx_to_db_PrintService(evtx + '\\' + 'Microsoft-Windows-PrintService%4Operational.evtx')
reg = r'D:\DFAS_PRO_X64_1.3.0.5_004\DFASPro\Evidence\CASE_202410291429\CollectSourceFile\20241029143259\Partition3\Windows\System32\config'
reg_to_db(reg)
shd = r"C:\Users\soke0\Desktop\00004.SHD"
shd_to_db(shd)