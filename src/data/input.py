import os
import sqlite3
import struct
import glob
from datetime import datetime
import pandas as pd
import win32evtlog
from Registry import Registry

class ForensicParser:
    def __init__(self, db_path='databases.db'):
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self.file_columns = {
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
                'table_name': 'file_system_volume_change_usnjrnl',
                'columns': {
                    '시간': 'time',
                    'USN': 'usn',
                    '파일명': 'file_name',
                    '경로': 'path',
                    '이벤트': 'event'
                }
            },
             '파일_시스템_파일_시스템_변경_이벤트($logfile).csv': {
                'table_name': 'file_system_volume_change_logfile',
                'columns': {
                    '시간': 'time',
                    '이벤트': 'event',
                    '파일명': 'file_name',
                    '경로': 'path',
                    '생성시간': 'creation_time',
                    '수정시간': 'modify_time',
                    'MFT 수정시간': 'M_modify_time',
                    '접근시간': 'access_time',
                    '상세정보': 'detail'
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
                    '만료 시간': 'expiration_time',
                    '활동 타입': 'activity_type'
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
                    '삭제': 'deleted'
                }
            }
        }
        self.init_database()

    def init_database(self):
        """Initialize database with complete schema"""
        if os.path.exists(self.db_path):
            os.remove(self.db_path)
            
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        
        # Create all tables
        table_schemas = []
        for file_info in self.file_columns.values():
            table_name = file_info['table_name']
            columns = file_info['columns'].values()
            column_defs = [f"{col} TEXT" for col in columns]
            schema = f'''
                CREATE TABLE IF NOT EXISTS {table_name} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    {', '.join(column_defs)}
                )
            '''
            table_schemas.append(schema)
        
        # Add additional tables
        table_schemas.extend(['''
            CREATE TABLE IF NOT EXISTS event_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_id INTEGER,
                time_generated TEXT,
                connection_status TEXT
            )
        ''', '''
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
        ''', '''
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
        ''', '''
            CREATE TABLE IF NOT EXISTS printer_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                creation_time TEXT,
                notify_name TEXT,
                document_name TEXT,
                printer_name TEXT
            )
        '''])
        
        # Execute all schemas
        for schema in table_schemas:
            self.cursor.execute(schema)
        
        self.conn.commit()

    def parse_all(self, csv_path=None, evtx_path=None, reg_path=None, spool_path=None):
        """Parse all artifact types"""
        try:
            if csv_path and os.path.exists(csv_path):
                print("\nProcessing CSV files...")
                self.parse_csv_files(csv_path)

            if evtx_path and os.path.exists(evtx_path):
                print("\nProcessing Event logs...")
                diagnostic_path = os.path.join(evtx_path, 'Microsoft-Windows-Partition%4Diagnostic.evtx')
                printservice_path = os.path.join(evtx_path, 'Microsoft-Windows-PrintService%4Operational.evtx')
                
                if os.path.exists(diagnostic_path):
                    self.parse_diagnostic_events(diagnostic_path)
                if os.path.exists(printservice_path):
                    self.parse_print_events(printservice_path)

            if reg_path and os.path.exists(reg_path):
                print("\nProcessing Registry files...")
                self.parse_registry(reg_path)

            if spool_path and os.path.exists(spool_path):
                print("\nProcessing Spool files...")
                self.parse_spool_files(spool_path)

            print("\nAll parsing tasks completed successfully")
            return True

        except Exception as e:
            print(f"\nError during parsing: {str(e)}")
            return False
        
    def parse_csv_files(self, path):
        """Parse CSV files with enhanced encoding handling"""
        encodings = ['utf-8-sig', 'cp949', 'euc-kr', 'utf-16']
        
        for file_name in os.listdir(path):
            if file_name in self.file_columns:
                file_path = os.path.join(path, file_name)
                table_info = self.file_columns[file_name]
                
                success = False
                for encoding in encodings:
                    try:
                        # BOM 체크
                        with open(file_path, 'rb') as f:
                            raw = f.read(4)
                            if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
                                encoding = 'utf-16'
                        
                        df = pd.read_csv(file_path, 
                                       encoding=encoding, 
                                       sep='\t',
                                       usecols=table_info['columns'].keys(),
                                       on_bad_lines='skip')
                        
                        df.rename(columns=table_info['columns'], inplace=True)
                        df = df.where(pd.notnull(df), None)
                        df.to_sql(table_info['table_name'], self.conn, 
                                if_exists='append', index=False)
                        
                        print(f"Successfully processed {file_name} using {encoding}")
                        success = True
                        break
                    
                    except Exception as e:
                        if encoding == encodings[-1]:
                            print(f"Failed to process {file_name}: {str(e)}")
                        continue

                if not success:
                    try:
                        with open(file_path, 'rb') as f:
                            content = f.read()
                            if content.startswith(b'\xff\xfe'):
                                content = content[2:]
                            elif content.startswith(b'\xef\xbb\xbf'):
                                content = content[3:]
                            
                            text = content.decode('utf-16le' if content.startswith(b'\xff\xfe') 
                                                else 'cp949', errors='ignore')
                            
                            from io import StringIO
                            df = pd.read_csv(StringIO(text), 
                                           sep='\t',
                                           usecols=table_info['columns'].keys(),
                                           on_bad_lines='skip')
                            
                            df.rename(columns=table_info['columns'], inplace=True)
                            df = df.where(pd.notnull(df), None)
                            df.to_sql(table_info['table_name'], self.conn, 
                                    if_exists='append', index=False)
                            
                            print(f"Successfully processed {file_name} using binary mode")
                            
                    except Exception as e:
                        print(f"All processing methods failed for {file_name}: {str(e)}")

        # Process 파일리스트_*.csv pattern files
        for file_path in glob.glob(os.path.join(path, '파일리스트_*.csv')):
            table_info = self.file_columns['파일리스트_*.csv']
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, 
                                   encoding=encoding, 
                                   sep='\t',
                                   usecols=table_info['columns'].keys(),
                                   on_bad_lines='skip')
                    
                    df.rename(columns=table_info['columns'], inplace=True)
                    df = df.where(pd.notnull(df), None)
                    df.to_sql(table_info['table_name'], self.conn, 
                            if_exists='append', index=False)
                    
                    print(f"Successfully processed {os.path.basename(file_path)} using {encoding}")
                    break
                    
                except Exception as e:
                    if encoding == encodings[-1]:
                        print(f"Failed to process {os.path.basename(file_path)}: {str(e)}")
                    continue

    def parse_diagnostic_events(self, file_path):
        """Parse diagnostic events"""
        event_log = win32evtlog.OpenBackupEventLog(None, file_path)
        try:
            flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
            while True:
                events = win32evtlog.ReadEventLog(event_log, flags, 0)
                if not events:
                    break

                for event in events:
                    if event.EventID == 1006:
                        message_data = event.StringInserts
                        connection_status = None
                        if message_data and len(message_data) > 3:
                            connection_status = ('Disconnected' if message_data[3].lower() == 'true' 
                                               else 'Connected')
                            
                        if connection_status:
                            self.cursor.execute('''
                                INSERT INTO event_log (event_id, time_generated, connection_status)
                                VALUES (?, ?, ?)
                            ''', (
                                event.EventID,
                                event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S'),
                                connection_status
                            ))
                            self.conn.commit()

        finally:
            win32evtlog.CloseEventLog(event_log)

    def parse_print_events(self, file_path):
        """Parse print service events"""
        event_log = win32evtlog.OpenBackupEventLog(None, file_path)
        try:
            status_map = {
                801: "Printing",
                802: "Deleted",
                307: "Completed",
                8421: "Network Connected"
            }
            
            flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
            while True:
                events = win32evtlog.ReadEventLog(event_log, flags, 0)
                if not events:
                    break

                for event in events:
                    if event.EventID in status_map:
                        event_data = {
                            'event_id': event.EventID,
                            'time_generated': event.TimeGenerated.strftime('%Y-%m-%d %H:%M:%S'),
                            'status': status_map[event.EventID],
                            'user': None,
                            'desktop': None,
                            'printer': None,
                            'port': None,
                            'size': None,
                            'page': None
                        }

                        if event.EventID == 307 and event.StringInserts:
                            if len(event.StringInserts) >= 8:
                                event_data.update({
                                    'user': event.StringInserts[2],
                                    'desktop': event.StringInserts[3],
                                    'printer': event.StringInserts[4],
                                    'port': event.StringInserts[5],
                                    'size': int(event.StringInserts[6]),
                                    'page': int(event.StringInserts[7])
                                })

                        self.cursor.execute('''
                            INSERT INTO driver_event_log 
                            (event_id, time_generated, status, user, desktop, printer, 
                             port, size, page)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', tuple(event_data.values()))
                        self.conn.commit()

        finally:
            win32evtlog.CloseEventLog(event_log)

    def parse_registry(self, path):
        """Parse registry for printer information"""
        software_path = os.path.join(path, "SOFTWARE")
        if not os.path.exists(software_path):
            print("SOFTWARE registry hive not found")
            return

        try:
            registry = Registry.Registry(software_path)
            printers_key = registry.open("Microsoft\\Windows NT\\CurrentVersion\\Print\\Printers")
            
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

                self.cursor.execute('''
                    INSERT INTO printers 
                    (printer_name, port, driver, print_processor, registry_path, 
                     collection_time, case_path)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    printer_info['printer_name'],
                    printer_info['port'],
                    printer_info['driver'],
                    printer_info['print_processor'],
                    printer_info['registry_path'],
                    datetime.now().isoformat(),
                    path
                ))
                self.conn.commit()

        except Registry.RegistryKeyNotFoundException:
            print("Printers registry key not found")
        except Exception as e:
            print(f"Error parsing registry: {str(e)}")

    def parse_spool_files(self, path):
        """Parse printer spool files"""
        for shd_file in glob.glob(os.path.join(path, '*.shd')):
            try:
                with open(shd_file, 'rb') as f:
                    data = f.read()

                if len(data) < 60:
                    continue

                document_name = self._read_utf16_string(data, 
                    struct.unpack('<I', data[40:44])[0])
                printer_name = self._read_utf16_string(data, 
                    struct.unpack('<I', data[56:60])[0])
                notify_name = self._read_utf16_string(data, 
                    struct.unpack('<I', data[24:28])[0])

                if any([document_name, printer_name, notify_name]):
                    self.cursor.execute('''
                        INSERT INTO printer_documents 
                        (creation_time, notify_name, document_name, printer_name)
                        VALUES (?, ?, ?, ?)
                    ''', (
                        datetime.fromtimestamp(os.path.getctime(shd_file)).isoformat(),
                        notify_name,
                        document_name,
                        printer_name
                    ))
                    self.conn.commit()

            except Exception as e:
                print(f"Error processing spool file {shd_file}: {str(e)}")

    def _read_utf16_string(self, data, offset):
        """Read UTF-16 string from binary data"""
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
        except Exception:
            pass
        return None

    def close(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()
            self.conn = None
            self.cursor = None

def main():
    """Main execution function"""
    print("=== Unified Forensic Data Parser ===\n")
    
    try:
        # Initialize parser
        parser = ForensicParser()
        
        # Get paths interactively or use defaults
        use_defaults = input("Use default paths? (y/n): ").lower().strip() == 'y'
        
        if use_defaults:
            # 기본 경로 설정
            csv_path = r'C:\Users\bidul\OneDrive\바탕 화면\새 폴더 (2)'
            evtx_path = r'C:\Users\bidul\OneDrive\바탕 화면\CrystalDiskInfo_8_17_14_leehs_bin\DFAS_PRO_X64_1.3.0.5_004\DFASPro\Evidence\CASE_202410291431\CollectSourceFile\20241029144715\Partition3\Windows\System32\winevt\Logs'
            reg_path = r'C:\Users\bidul\OneDrive\바탕 화면\CrystalDiskInfo_8_17_14_leehs_bin\DFAS_PRO_X64_1.3.0.5_004\DFASPro\Evidence\CASE_202410291431\CollectSourceFile\20241029144715\Partition3\Windows\System32\config'
            spool_path = r"C:\Windows\System32\spool\PRINTERS"
        else:
            # 사용자로부터 경로 입력 받기
            csv_path = input("\nEnter path to CSV files directory (or press Enter to skip): ").strip() or None
            evtx_path = input("Enter path to Event Log files directory (or press Enter to skip): ").strip() or None
            reg_path = input("Enter path to Registry files directory (or press Enter to skip): ").strip() or None
            spool_path = input("Enter path to Spool files directory (or press Enter to skip): ").strip() or None

        # Validate paths
        paths = {
            'CSV': csv_path,
            'Event Logs': evtx_path,
            'Registry': reg_path,
            'Spool': spool_path
        }

        print("\nValidating paths...")
        for name, path in paths.items():
            if path:
                if os.path.exists(path):
                    print(f"{name} path valid: {path}")
                else:
                    print(f"Warning: {name} path does not exist: {path}")
            else:
                print(f"Note: {name} path not provided - skipping")

        # Confirm processing
        if input("\nProceed with processing? (y/n): ").lower().strip() != 'y':
            print("Operation cancelled by user")
            return

        # Process files
        success = parser.parse_all(csv_path, evtx_path, reg_path, spool_path)
        
        # Close database connection
        parser.close()

        if success:
            print(f"\nProcessing completed successfully. Results saved to {parser.db_path}")
            
            # Display basic statistics
            conn = sqlite3.connect(parser.db_path)
            cursor = conn.cursor()
            
            print("\nDatabase Statistics:")
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            
            for table in tables:
                table_name = table[0]
                cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                count = cursor.fetchone()[0]
                print(f"- {table_name}: {count} records")
            
            conn.close()
        else:
            print("\nErrors occurred during processing. Check the output above for details.")

    except Exception as e:
        print(f"\nError in main execution: {str(e)}")
        if 'parser' in locals():
            parser.close()

if __name__ == "__main__":
    main()