from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from common_utils import get_config_from_json, get_xlsx_file_name, get_worksheet_name, get_apart_object

import sys

FIRST_ITEM_ROW_INDEX = 3

def insert_image_with_cell_height(ws, image_path, cell):
    img = Image(image_path)
    ws.add_image(img, cell.coordinate)

# def is_image_cell(ws, cell):
#     return any(image.anchor == cell.coordinate for image in ws._images)

def is_image_cell(image_obj, cell):
    return cell.coordinate in image_obj.anchors

def 사진체크하기(image_obj, ws, row):
    
    현관사진_cell = ws.cell(row=row, column=1)
    큐알사진_cell = ws.cell(row=row, column=3)
    
    return is_image_cell(image_obj, 현관사진_cell) and is_image_cell(image_obj, 큐알사진_cell)

def count_images(ws, image_obj, n):
    
    image_count = 0
    
    for idx in range(n):
        row_index = FIRST_ITEM_ROW_INDEX + 2 * idx
        if 사진체크하기(image_obj, ws, row_index):
            image_count += 1

    return image_count

def get_image_object(ws):
    image_obj = {
        images: ws._images,
        anchors: { img.anchor: img for img in ws._images },
    }
    return image_obj

def 워크시트_하나_조사하기(ws, n):
    
    image_obj = get_image_object(ws)
    
    cnt = count_images(ws, image_obj, n)

def 워크시트_유니크_사진수(ws):
    count = len(set([ img.anchor for img in ws._images ]))
    return count

def 워크시트_작업량(ws):
    return 워크시트_유니크_사진수(ws) // 2

def 진행률체크위한_사진체크하기(아파트객체, 엑셀파일명):
    
    단지명 = 아파트객체["단지명"]
    동호수목록 = 아파트객체["동호수목록"]
    대상세대수 = 아파트객체["대상세대수"]
    
    # 1. 타겟 파일 열기
    파일경로 = get_xlsx_file_name(단지명)
    wb = load_workbook(filename = 파일경로)
    
    # 2. 워크시트를 전체를 순회하면서, 각 시트에서 이미지 카운트 => 누적
    # 완료세대수 = sum([ 워크시트_하나_조사하기(wb[get_worksheet_name(단지명, 몇동)], len(호수목록)) for 몇동, 호수목록 in 동호수목록.items() ])
    완료세대수 = sum([ 워크시트_작업량(wb[get_worksheet_name(단지명, 몇동)]) for 몇동 in 동호수목록.keys() ])
    
    # 3. 열었던 타겟파일은 닫기
    wb.close()
    
    return 완료세대수

def 아파트단지_하나_카운팅(config, 단지명, 프린트여부 = True, 파일업데이트여부 = False):
    
    아파트목록 = config['아파트목록']
    # 설정파일명 = config["설정파일명"]
    엑셀파일명 = config['엑셀파일명']
    
    아파트객체 = get_apart_object(config, 단지명)

    완료세대수 = 진행률체크위한_사진체크하기(아파트객체, 엑셀파일명)
    대상세대수 = 아파트객체["대상세대수"]
    
    # 1. 프린트
    if 프린트여부:    
        진행률 = (완료세대수 / 대상세대수) * 100
        print(f"{단지명}: {완료세대수}호 / 전체 {대상세대수} 호 = {진행률}%")
    
    단지명_열 = 2
    대상세대수_열 = 3
    완료세대수_열 = 4
    
    # 2. 요약 파일 업데이트
    if 파일업데이트여부:
        wb = load_workbook(filename = 엑셀파일명)
        ws = wb.active
        for row in ws.iter_rows():
            if row[단지명_열] == 단지명:
                row[완료세대수_열] = 완료세대수
                break
        else:
            raise Exception(f"기존 요약파일에서 단지명({단지명})에 대한 row를 찾지 못했습니다.")
        wb.save(엑셀파일명)
        wb.close()

if __name__ == '__main__':

    config = get_config_from_json('./apartments.json')
    # 1. 커맨드라인 인수 체크
    단지명 = sys.argv[1]
    프린트여부 = True
    파일업데이트여부 = False
    아파트단지_하나_카운팅(config, 단지명, 프린트여부, 파일업데이트여부)