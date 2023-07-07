import sys  # for 실행인자
import os  # for 파일시스템 조회

from common_utils import get_xlsx_file_name, sorted_file_entries, load_json, get_apart_object, get_apart_object2

from data_processing_v1 import update_one_apartment

from counting import 카운팅_업데이트

import json
from datetime import date

import traceback

from colorama import Fore, Back, Style

def get_entries(folder_path):

    pass_files = []
    folder_entries = []
    
    for entry in os.scandir(folder_path):
        
        if entry.is_file(): # 폴더만 순회할 것임. 파일은 기록용으로 출력만
            
            pass_files.append(entry.name)

        else:

            folder_entries.append(entry)

    return pass_files, folder_entries

def update_all_apartment(config, base_path, folder_entries):
    # 서머리용
    결과객체모음 = dict()

    # 각 폴더는 아파트 단지 하나 => 단지명
    print("단지별 업데이트 시작")
    for i, entry in enumerate(folder_entries, start=1):
        try:
            '''
            folder_path => 해당 폴더 아래의 모든 파일을 DirEntry로 추출할때 사용하기 위함.
            단지명 => 아파트 객체 & 해당 아파트 엑셀 파일 경로
            base_path => 해당 아파트 엑셀 파일 경로
            '''
            # 단지명 = name[2:] # 앞에 두자리 제외 => 01서울번동3 => 서울번동3
            순번 = int(entry.name[:2])

            아파트객체 = get_apart_object2(config, 순번)

            print(
                f"{Fore.RED}({i}/{len(folder_entries)}){아파트객체['단지명']}단지 {아파트객체['대상세대수']}세대에 대한 작업 시작{Style.RESET_ALL}")

            결과객체 = update_one_apartment(
                config, entry.path, 아파트객체["단지명"], base_path)

            결과객체모음[아파트객체["단지명"]] = 결과객체

            print(f"{Fore.BLUE}({i}/{len(folder_entries)}) 작업 종료{Style.RESET_ALL}")

        except Exception as error:

            print(f"{entry.name}폴더명을 가진 작업분에 대해 실패")
            print(error)
            traceback.print_exc()

    return 결과객체모음

def 카운팅(config, 결과객체모음):
    
    print("카운팅 시작")

    try:

        카운팅_업데이트(config, 결과객체모음, 프린트여부=True, 파일업데이트여부=True)

    except Exception as e:

        print(f"카운팅 과정에서 오류: {e}")
        traceback.print_exc()

def logging(결과객체모음):
    
    print("덤프 시작")
    
    try:
        today = date.today()
        filename = f"log_{today.strftime('%Y%m%d')}.json"

        with open(filename, "w", encoding="utf-8") as json_file:
            json.dump(결과객체모음, json_file, ensure_ascii=False)

    except Exception as e:
        print(f"덤프 과정에서 오류: {e}")
        traceback.print_exc()

DEFAULT_CONFIG_FILE_PATH = "./apartments.json"

'''
python3 update_in_folders.py <folder_path> <base_path>

- folder_path: 업데이트할 타겟은 사진들이 폴더별로 정리되어있는 루트 주소
- base_path: 업데이트될 엑셀 파일의 디렉토리 주소

'''
if __name__ == '__main__':

    config = load_json(DEFAULT_CONFIG_FILE_PATH)

    folder_path = sys.argv[1]
    base_path = "." if len(sys.argv) < 3 else sys.argv[2]

    pass_files, folder_entries = get_entries(folder_path)
    
    결과객체모음 = update_all_apartment(config, base_path, folder_entries)
    카운팅(config, 결과객체모음)
    logging(결과객체모음)

    print("전체 프로세스 종료")
