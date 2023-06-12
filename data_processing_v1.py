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

# 이미지의 크기와 픽셀간 조절값(실험적) => 픽셀의 크기가 다른 환경에서는 추가적인 조절 필요
PX_TO_PT = 3 / 4
PT_TO_PX = 4 / 3 # 1.33

# 이미지 4:3으로 조절
IMAGE_CELL_HEIGHT_PT = 160
IMAGE_CELL_HEIGHT_PX = IMAGE_CELL_HEIGHT_PT * PT_TO_PX

IMAGE_ASPECT_HEIGHT = 4
IMAGE_ASPECT_WIDTH = 3

IMAGE_CELL_WIDTH_PX = IMAGE_CELL_HEIGHT_PX * IMAGE_ASPECT_WIDTH / IMAGE_ASPECT_HEIGHT

# 각 시트의 상단부에 들어갈 컨텐츠를 정하는 부분
CELL_HEAD = 'A1'
MERGED_CELL_HEAD_RANGE = 'A1:D1'
HEAD_STRING = '세대별 부착 사진대지'

CELL_LABEL = 'A2'
CELL_VALUE = 'B2'
LABEL_STRING = '단지명:'

# 시트 상단부에 들어갈 컨텐츠에 따라, 목록의 시작 위치가 결정.
FIRST_ITEM_ROW_INDEX = 3
ENUM_현관사진_COLUMN = 1
ENUM_큐알사진_COLUMN = 3

def get_worksheet_name(단지명, 몇동):
    return f"{단지명}-{몇동}동"

def insert_image_with_cell_height(ws, image_path, cell):
    img = Image(image_path)

    # 이미지 사이즈 조절, 4:3로 강제 조절
    img.height = IMAGE_CELL_HEIGHT_PX
    img.width = IMAGE_CELL_WIDTH_PX

    # 이미지 삽입
    ws.add_image(img, cell.coordinate)

def sorted_filelist(folder_path):
    # 폴더 내의 파일 목록을 가져옴
    # files = os.listdir(folder_path)
    
    # # 폴더 내부의 모든 파일 및 하위 폴더 순회
    # for root, dirs, files in os.walk(folder_path):
    #     # 파일들을 이름순으로 정렬
    #     sorted_files = sorted(files)
        
    #     # 정렬된 파일들을 출력
    #     if DEBUG_MODE:        
    #         for file in sorted_files:
    #             print(file)
        
    #     # 파일 목록에 추가
    #     filelist.extend(sorted_files)
    filelist = []
    # 하위 1개의 폴더까지만 순회
    for entry in os.scandir(folder_path):
        if entry.is_file():
            # 파일인 경우 파일 목록에 추가
            # filelist.append(entry.name)
            pass
        elif entry.is_dir():
            # 폴더인 경우 해당 폴더 내의 파일 목록을 가져와 파일 목록에 추가
            subfolder_files = os.listdir(entry.path)
            filelist.extend(subfolder_files)

    return sorted(filelist)

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
    
    filename = config["엑셀파일명"]
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
    if target_동 in 동호수목록:
        호수목록 = 동호수목록[target_동]
        return target_호수 in 호수목록
    # for item in 동호수목록:
    #     동, 시작호수, 끝호수 = item
    #     if target_동 == 동 and 시작호수 <= target_호수 and target_호수 <= 끝호수:
    #         return item
    return None

def 기존엑셀에_이미지_한개_삽입하기(wb, 현관사진_파일경로, 큐알사진_파일경로, 워크시트이름, 시작호수와의차이):
    ws = wb[워크시트이름]
    
    row_index = FIRST_ITEM_ROW_INDEX + 시작호수와의차이 * 2
    
    현관사진_cell = ws.cell(row=row_index, column=ENUM_현관사진_COLUMN)
    큐알사진_cell = ws.cell(row=row_index, column=ENUM_큐알사진_COLUMN)

    insert_image_with_cell_height(ws, 현관사진_파일경로, 현관사진_cell)
    insert_image_with_cell_height(ws, 큐알사진_파일경로, 큐알사진_cell)


def 작업분_기존엑셀에_반영하기(wb, folder_path, 동호수목록, sorted_files):
    
    작업분 = defaultdict(lambda: ['', '', '', -1])
    실패1 = []

    for filename in sorted_files:
        rstr = r'^(\d+)동\s*(\d+)호\s*(\(1\))?(\(2\))?\s*.(png|jpg|jpeg)$' #공백 허용 버전, 이미지 포맷 확대
        # r = r'^(\d+)동\s*(\d+)호(\(2\))?\s*.png$'
        matched = re.fullmatch(rstr, filename)
        if matched:
            동_문자열, 호수_문자열, _원, 큐알사진여부, 확장자 = matched.groups(False)
            동, 호수 = map(int, [동_문자열, 호수_문자열])
            사진모드 = ENUM_큐알사진 if 큐알사진여부 else ENUM_현관사진

            # 유효성 체크
            # validate_result = 동호수_유효성체크(동호수목록, 동, 호수)
            
            # # 워크시트이름, 시작호수와의차이 = 동호수_유효성체크(동호수목록, 동, 호수)
            # if validate_result != None:
            #     동, 시작호수, 끝호수 = validate_result
            #     작업분[(동, 호수)][사진모드] = os.path.join(foldername, filename)
            #     작업분[(동, 호수)][2] = get_worksheet_name(동, 끝호수)
            #     작업분[(동, 호수)][3] = 호수 - 시작호수
            #     continue
            # 워크시트이름, 시작호수와의차이 = 동호수_유효성체크(동호수목록, 동, 호수)
            # 동, 시작호수, 끝호수 = validate_result

            if 동_문자열 not in 동호수목록:
                print(f"<{동}>동은 동호수목록{동호수목록.keys()}에 포함되지 않습니다.")
                실패1.append(filename)
                continue
            호수목록 = 동호수목록[동_문자열]

            for i, value in enumerate(호수목록):
                if value == 호수:
                    차이 = i
                    break
            else:
                print(f"<{호수}>호수는 호수목록{호수목록}에 포함되지 않습니다.")
                실패1.append(filename)
                continue

            작업분[(동, 호수)][사진모드] = os.path.join(folder_path, filename)
            작업분[(동, 호수)][2] = get_worksheet_name(단지명, 동)
            작업분[(동, 호수)][3] = 차이
        # 실패 항목 기록
        else:
            실패1.append(filename)
    
    성공 = []
    실패2 = []
    # print(len(작업분))
    for key, value in 작업분.items():
        동, 호수 = key
        현관사진_파일경로, 큐알사진_파일경로, 워크시트이름, 시작호수와의차이 = value
        if 현관사진_파일경로 != '' and 큐알사진_파일경로 != '':
            기존엑셀에_이미지_한개_삽입하기(wb, 현관사진_파일경로, 큐알사진_파일경로, 워크시트이름, 시작호수와의차이)
            성공.append(value)
        else:
            실패2.append(value)
    print("실패1", 실패1)
    # print("성공", 성공)
    print("실패2", [x[0] for x in 실패2])

def load_json(filepath='apartments.json'):
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
    폴더경로 = sys.argv[1]
    config["foldername"] = sys.argv[1]
    
    for 아파트객체 in config["아파트목록"]:
        if 아파트객체["단지명"] == "서울월계":
            동호수목록 = 아파트객체["동호수목록"]
    if 동호수목록:
        
        wb = load_excel(config)

        for entry in os.scandir(폴더경로):
            if entry.is_file():
                # 파일인 경우 파일 목록에 추가
                # filelist.append(entry.name)
                pass
            elif entry.is_dir():
                # 폴더인 경우 해당 폴더 내의 파일 목록을 가져와 파일 목록에 추가
                # print(entry.path)
                
                subfolder_files = os.listdir(entry.path)
                sorted_subfolder_files = sorted(os.listdir(entry.path))
                
                # print(sorted_subfolder_files)
                
                작업분_기존엑셀에_반영하기(wb, entry.path, 동호수목록, sorted_subfolder_files)
        
        wb.save(config["확장자포함파일명"])
        wb.close()

    # sorted_files = sorted_filelist(config["foldername"])
    
    # # print(sorted_files[:10])
    
    # if not sorted_files:
    #     print(f"{config['foldername']}폴더 아래에 파일이 존재하지 않습니다.")
    #     sys.exit()
        
    # wb = load_excel(config)



    # 작업분_기존엑셀에_반영하기(wb, config["foldername"], config["동호수목록"], sorted_files)
    
    # wb.save(config["확장자포함파일명"])
    # wb.close()