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

    # 이미지 사이즈 조절
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

def load_excel(config):
    
    filename = config["파일명"]
    extname = extract_extension(filename)
    
    if extname == None:
        config["확장자포함파일명"] = filename + ".xlsx"
    elif extname != "xlsx":
        print(f"파일명의 확장자가 올바르지 않습니다. (config[\"파일명\"]: {filename}")
        sys.exit()
    else:
        config["확장자포함파일명"] = filename

    wb = load_workbook(filename = config["확장자포함파일명"])

    return wb
    
def 동호수_유효성체크(동호수목록, target_동, target_호수):
    
    for item in 동호수목록:
        동, 시작호수, 끝호수 = item
        if target_동 == 동 and 시작호수 <= target_호수 and target_호수 <= 끝호수:
            return item
    return None

def 기존엑셀에_이미지_한개_삽입하기(wb, 현관사진_파일명, 큐알사진_파일명, 워크시트명, 시작호수와의차이):
    print(f"워크시트명: {워크시트명}")
    ws = wb[워크시트명]
    
    row_index = FIRST_ITEM_ROW_INDEX + 시작호수와의차이 * 2
    
    현관사진_cell = ws.cell(row=row_index, column=1)
    큐알사진_cell = ws.cell(row=row_index, column=3)
    insert_image_with_cell_height(ws, 현관사진_파일명, 현관사진_cell)
    insert_image_with_cell_height(ws, 큐알사진_파일명, 큐알사진_cell)


def 작업분_기존엑셀에_반영하기(wb, foldername, 동호수목록, sorted_files):
    작업분 = defaultdict(lambda: ['', '', '', -1])
    실패1 = []
    for filename in sorted_files:
        r"*.xlxs"
        r = r'^(\d+)동(\d+)호(\(2\))?.png$'
        matched = re.fullmatch(r, filename)
        if matched:
            동, 호수, 큐알사진여부 = matched.groups(False)
            동, 호수 = map(int, [동, 호수])
            사진모드 = ENUM_큐알사진 if 큐알사진여부 else ENUM_현관사진

            # 유효성 체크
            validate_result = 동호수_유효성체크(동호수목록, 동, 호수)
            
            # 워크시트명, 시작호수와의차이 = 동호수_유효성체크(동호수목록, 동, 호수)
            if validate_result != None:
                동, 시작호수, 끝호수 = validate_result
                작업분[(동, 호수)][사진모드] = os.path.join(foldername, filename)
                작업분[(동, 호수)][2] = get_worksheet_name(동, 끝호수)
                작업분[(동, 호수)][3] = 호수 - 시작호수
                continue
        # 실패 항목 기록
        실패1.append(filename)
    
    성공 = []
    실패2 = []
    for key, value in 작업분.items():
        동, 호수 = key
        현관사진_파일명, 큐알사진_파일명, 워크시트명, 시작호수와의차이 = value
        if 현관사진_파일명 != '' and 큐알사진_파일명 != '':
            기존엑셀에_이미지_한개_삽입하기(wb, 현관사진_파일명, 큐알사진_파일명, 워크시트명, 시작호수와의차이)
            성공.append(value)
        else:
            실패2.append(value)
    print("실패1", 실패1)
    print("성공", 성공)
    print("실패2", [x[0] for x in 실패2])

def 아파트객체_추출(config, 단지명):
    for 아파트객체 in config["아파트목록"]:
        if 단지명 == 아파트객체["단지명"]:
            return 아파트객체
    return None

DEFAULT_CONFIG_FILE_PATH = "./apartments.json"

if __name__ == '__main__':
    
    config = get_config_from_json(DEFAULT_CONFIG_FILE_PATH)
    
    folder_path = sys.argv[1]
    단지명 = sys.argv[2]
    
    # 유효한 아파트 단지인지 체크
    아파트객체 = 아파트객체_추출(config, 단지명)
    if 아파트객체 == None:
        raise Exception(f"단지명 '{단지명}'은 유효하지 않습니다.")
    
    # 폴더 하위 모든 파일 추출
    sorted_files = sorted_files(folder_path)
    
    # 하나도 파일이 없으면
    if not sorted_files:
        raise Exception(f"{folder_path}디렉토리 아래에 파일이 존재하지 않습니다.")

    # 하나의 단지에 대해서 수행하면 끝
    
    # 1.기존 엑셀 파일 열고
    # 기존의 load_excel 함수는 안 써도 될듯 => 파일명 규칙만 정해서, 단지명 => 파일 경로 => 해당 경로로 열기
    파일경로 = get_xlsx_file_name(단지명)
    wb = load_workbook(filename = 파일경로)
    # wb = load_excel(config)
    
    # 2. 아파트객체에서 동호수목록을 가지고 유효성체크하며 기존 엑셀에 반영
    # 단지당 엑셀파일 하나 => 동별 시트로 구분 => 모든 파일을 순회하며 유효한 것에 대해서 기존 엑셀에 반영
    동호수목록 = 아파트객체["동호수목록"]
    작업분_기존엑셀에_반영하기(wb, folder_path, 동호수목록, sorted_files)
    
    # 3. 작업 완료후 아래 코드로 저장하여 반영하기
    wb.save(파일경로)
    # wb.save(config["확장자포함파일명"])
    wb.close()