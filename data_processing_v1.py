# for openpyxl, 엑셀류를 다루는 표준을 따르는 라이브러리, Image를 다룰때 내부적으로 pillow를 사용해서 인스톨 필수
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment
#from PIL import Image as PILImage # 이미지 크기를 줄이고 싶을때, 이것을 활용해야할듯.

import json # json 파일에서 설정을 사용하고자함
import sys # 파일 실행시 인자 받아서 활용하기
import os # 해당경로 하위 파일 파악하는데 활용
import re # 파일명에서 정보 추출하는데 활용

from collections import defaultdict

# 두개이상의 파일에서 공통으로 가져가야할 규칙은 import해서 사용하기
from common_utils import get_worksheet_name, get_xlsx_file_name, sorted_files, get_config_from_json

DEBUG_MODE = False

PX_TO_PT = 3 / 4
PT_TO_PX = 4 / 3 # 1.33

IMAGE_CELL_HEIGHT_PT = 160
IMAGE_CELL_HEIGHT_PX = IMAGE_CELL_HEIGHT_PT * PT_TO_PX

IMAGE_CELL_WIDTH_PT = 240
IMAGE_CELL_WIDTH_PT_6 = IMAGE_CELL_WIDTH_PT / 6

CELL_HEAD = 'A1'
MERGED_CELL_HEAD_RANGE = 'A1:D1'
HEAD_STRING = '세대별 부착 사진대지'

CELL_LABEL = 'A2'
CELL_VALUE = 'B2'
LABEL_STRING = '단지명:'

FIRST_ITEM_ROW_INDEX = 3

def insert_image_with_cell_height(ws, image_path, cell):
    img = Image(image_path)

    # 이미지 사이즈 조절 => 원본 크기를 조절하는것은 아니였던듯? 하려면 pillow의 Image 직접 사용해야하는듯
    img.width = IMAGE_CELL_HEIGHT_PX * img.width / img.height
    img.height = IMAGE_CELL_HEIGHT_PX

    # 이미지 삽입
    ws.add_image(img, cell.coordinate)

ENUM_현관사진 = 0
ENUM_큐알사진 = 1

def check_extension(filename, extname):
    pattern = rf"\.{extname}$"
    return bool(re.search(pattern, filename))

def extract_extension(filename):
    pattern = r"\.([^.]+)$"
    match = re.search(pattern, filename)
    if match:
        return match.group(1)
    else:
        return None

def extract_filename_components(filename):
    pattern = r"^(.+)\.([^.]+)$"
    match = re.match(pattern, filename)
    if match:
        basename, extension = match.groups("")
        return basename, extension
    else:
        return "", "", ""

def 동호수_유효성체크(동호수목록, 동, 호수):
    try:
        return 동호수목록[동].index(int(호수))
    except KeyError:
        return -1
    except ValueError:
        return -1

def 워크시트하나에_이미지_한개_삽입하기(ws, 현관사진_파일명, 큐알사진_파일명, 인덱스):
    
    
    row_index = FIRST_ITEM_ROW_INDEX + 인덱스 * 2
    
    현관사진_cell = ws.cell(row=row_index, column=1)
    큐알사진_cell = ws.cell(row=row_index, column=3)
    insert_image_with_cell_height(ws, 현관사진_파일명, 현관사진_cell)
    insert_image_with_cell_height(ws, 큐알사진_파일명, 큐알사진_cell)


def 작업분_기존엑셀에_반영하기(wb, foldername, 동호수목록, filenames):
    # 동별로 하나의 시트 => 동별로 구분 짓기.
    동별_작업분 = defaultdict(defaultdict(lambda: ['', '', -1]))
    ''' 예시
    동별_작업분[101][777] = [현관사진경로, 큐알사진경로, 워크시트명, 호수 - 시작호수]
    - 워크시트명과 호수-시작호수는 필요 없어진것일수도 있지만.... 일단 냅두자
    '''
    
    성공적으로업데이트, 파일명이_매칭이_안되어_실패, 유효성체크_실패, 사진_하나라도_없어_실패 = [], [], []
    
    for filename in filenames:

        r = r'^(\d+)동\s*(\d+)호\s*(\(1\))?(\(2\))?\s*.(png|jpg|jpeg)$'
        matched = re.fullmatch(r, filename)

        if not matched:
            파일명이_매칭이_안되어_실패.append(filename)
            continue
    
        동, 호수, 큐알사진여부 = matched.groups(False)
        인덱스 = 동호수_유효성체크(동호수목록, 동, 호수)

        if 인덱스 == -1:
            유효성체크_실패.append(filename)
            continue
    
        동, 호수 = map(int, [동, 호수])
        사진모드 = ENUM_큐알사진 if 큐알사진여부 else ENUM_현관사진
    
        # 아마 이렇게 바뀔것
        동별_작업분[동][호수][사진모드] = os.path.join(foldername, filename) # 이걸 어떻게...?
        동별_작업분[동][호수][3] = 인덱스
        
    
    for 동, 호별작업물 in 동별_작업분.items():
        
        워크시트명 = get_worksheet_name(단지명, 동)
        ws = wb[워크시트명]
        
        for 현관사진_파일명, 큐알사진_파일명, 인덱스 in 호별작업물.values():
            
            if 현관사진_파일명 != '' and 큐알사진_파일명 != '':
                워크시트하나에_이미지_한개_삽입하기(ws, 현관사진_파일명, 큐알사진_파일명, 인덱스)
                성공적으로업데이트.append(value)
            else:
                사진_하나라도_없어_실패.append(value)

    print("성공", 성공적으로업데이트)
    print("실패1", 파일명이_매칭이_안되어_실패)
    print("실패2", 유효성체크_실패)
    print("실패3", 사진_하나라도_없어_실패)

def 아파트객체_추출(config, 단지명):
    for 아파트객체 in config["아파트목록"]:
        if 단지명 == 아파트객체["단지명"]:
            return 아파트객체
    return None

DEFAULT_CONFIG_FILE_PATH = "./apartments.json"

'''
1. 먼저 설정파일을 읽어서 아파트목록 데이터를 취한다.
2. 반영할 폴더명 & 단지명을 커맨드라인인수로 입력 받는다.
3. 단지명의 유효성 체크후, 유효하면 해당 아파트객체와 해당 아파트의 엑셀파일을 준비한다.
4. 폴더를 순회하여, 모든 파일에 대하여 엑셀파일에 삽입할 준비를 한다. => 덮어쓰기모드?
5. 업데이트된 엑셀파일을 저장한다.
'''
if __name__ == '__main__':
    
    config = get_config_from_json(DEFAULT_CONFIG_FILE_PATH)
    
    # 1. 커맨드라인 인수로부터 정보 입력 받음
    folder_path = sys.argv[1]
    단지명 = sys.argv[2]
    
    # 2. 유효한 아파트 단지인지 체크
    아파트객체 = 아파트객체_추출(config, 단지명)
    if 아파트객체 == None:
        raise Exception(f"단지명 '{단지명}'은 유효하지 않습니다.")
    
    # 3. 폴더 하위 모든 파일 추출
    filenames = sorted_files(folder_path)
    
    # 3-1. 하나도 파일이 없으면 에러
    if not filenames:
        raise Exception(f"{folder_path}디렉토리 아래에 파일이 존재하지 않습니다.")

    # 4. 아파트단지 하나에 대한 엑셀파일(워크북)을 연다.
    파일경로 = get_xlsx_file_name(단지명)
    wb = load_workbook(filename = 파일경로)
    
    # 5. 아파트객체에서 동호수목록을 가지고 유효성체크하며 기존 엑셀에 반영
    동호수목록 = 아파트객체["동호수목록"]
    작업분_기존엑셀에_반영하기(wb, folder_path, 동호수목록, filenames)
    
    # 6. 작업 완료후 아래 코드로 저장하여 반영하기
    wb.save(파일경로)
    wb.close()