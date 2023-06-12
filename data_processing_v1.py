from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment
from PIL import Image as PILImage

import json

import sys
import os

import re

from collections import defaultdict

from openpyxl import load_workbook

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

def get_worksheet_name(몇동, 끝호수):
    return f"{몇동}동(~{끝호수}호)"

def insert_image_with_cell_height(ws, image_path, cell):
    img = Image(image_path)

    # 이미지 사이즈 조절
    img.width = IMAGE_CELL_HEIGHT_PX * img.width / img.height
    img.height = IMAGE_CELL_HEIGHT_PX

    # 이미지 삽입
    ws.add_image(img, cell.coordinate)

def sorted_filelist(foldername):
    # 폴더 내의 파일 목록을 가져옴
    # files = os.listdir(foldername)
    
    # # 폴더 내부의 모든 파일 및 하위 폴더 순회
    # for root, dirs, files in os.walk(foldername):
    #     # 파일들을 이름순으로 정렬
    #     sorted_files = sorted(files)
        
    #     # 정렬된 파일들을 출력
    #     if DEBUG_MODE:        
    #         for file in sorted_files:
    #             print(file)
        
    #     # 파일 목록에 추가
    #     filelist.extend(sorted_files)
    
    # 하위 1개의 폴더까지만 순회
    for entry in os.scandir(foldername):
        if entry.is_file():
            # 파일인 경우 파일 목록에 추가
            # filelist.append(entry.name)
        elif entry.is_dir():
            # 폴더인 경우 해당 폴더 내의 파일 목록을 가져와 파일 목록에 추가
            subfolder_files = sorted(os.listdir(entry.path))
            filelist.extend(subfolder_files)

    return sorted_files

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

def load_json(filepath='aparts.json'):
    try:
        with open(filepath, 'r') as f:
            config = json.load(f)
    except FileNotFoundError:
        raise Exception(f"파일 '{filepath}'을(를) 찾을 수 없습니다.")
    except json.JSONDecodeError:
        raise Exception(f"파일 '{filepath}'의 JSON 형식이 올바르지 않습니다.")

    return config

if __name__ == '__main__':
    
    config = load_json()
    config["foldername"] = sys.argv[1]
    
    sorted_files = sorted_filelist(config["foldername"])
    
    print(sorted_files)
    
    # if not sorted_files:
    #     print(f"{config['foldername']}폴더 아래에 파일이 존재하지 않습니다.")
    #     sys.exit()
        
    # wb = load_excel(config)

    # 작업분_기존엑셀에_반영하기(wb, config["foldername"], config["동호수목록"], sorted_files)
    
    # wb.save(config["확장자포함파일명"])
    # wb.close()