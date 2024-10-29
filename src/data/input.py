"""
본 모듈은 다음 파일들의 정보를 sqlite db의 형태로 바꾸는 것을 목적으로 한다.
# 1. 경로 - DFAS의 일괄 추출의 결과물인, *.csv
# 2. 경로 - 이벤트로그들 *.evtx
# 3. 경로 - 레지스트리 하이브 파일들
4. 경로 - 인쇄 관련 파일, *spl, *.shd
10/30 (수) 까지 완료
"""

import pandas as pd
import sqlite3
import os

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

def evtx_to_db():
    #
    pass

def reg_to_db():
    #
    pass

def spl_to_db():
    #
    pass

def activitycache_to_db():
    #
    pass