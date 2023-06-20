from openpyxl import load_workbook
import json
from collections import defaultdict

def is_dong_ho(dong, ho):
    return isinstance(dong, int) and isinstance(ho, int)

def process_index_dong_ho(apartment_data, max_index, index, dong, ho):
    if is_dong_ho(dong, ho):
        apartment_data[dong].append(ho)
        if isinstance(index, int):
            max_index = max(max_index, index)
    return max_index

def get_area(단지명):
    data = {
        "서울번동3": '강북구',
        "서울번동5": '강북구',
        "서울번동2": '강북구',
        "서울가양": '강서구',
        "서울등촌9": '강서구',
        "서울등촌7": '강서구',
        "서울등촌1": '강서구',
        "서울등촌4": '강서구',
        "서울등촌6": '강서구',
        "서울등촌11": '강서구',
        "서울중계1": '노원구',
        "서울중계3": '노원구',
        "서울중계3(주거복지동)": '노원구',
        "서울중계9": '노원구',
        "서울중계9(주거복지동)": '노원구',
        "서울월계": '노원구',
        "서울오류": '노원구',
        "서울공릉": '노원구',
        "서울가좌": '마포구',
        "서울중구": '중구',
    }
    return data[단지명]

def get_apartment_info(wb, sheetname, index):
    
    # 시트 열기
    ws = wb[sheetname]
    
    # 시트별로 동호수를 저장할 리스트 생성
    apartment_data = defaultdict(list)
    max_index = 0

    # 시트의 모든 행 순회
    for row in ws.iter_rows(min_row=3):
        # 동호수가 있는 열의 값을 조회 + 순번(처음부터 빈 집은 순번이 비어있을 수 있다! => 빼고 카운팅 됨)
        max_index = process_index_dong_ho(apartment_data, max_index, row[0].value, row[1].value, row[2].value)
        max_index = process_index_dong_ho(apartment_data, max_index, row[5].value, row[6].value, row[7].value)

    # 정렬하기
    for k, v in apartment_data.items():
        apartment_data[k] = sorted(v)
    
    apartment_info = {
        '단지명': sheetname,
        '동호수목록': apartment_data,
        # '동호수목록2': new_data,
        '대상세대수': max_index,
        '순번': index, # 1~20,
        '지역구명': get_area(sheetname),
    }

    return apartment_info

def get_new_data(apartment_data):
    result = dict()
    for k, v in apartment_data.items():
        lst = []
        i = 0
        while i < len(v):
            j = 1
            while i + j < len(v) and v[i]+j == v[i+j]:
                j += 1
            lst.append([v[i], v[i] + j - 1])
            i += j 
        result[k] = lst
    return result

def analyze_apartments(추출할엑셀파일경로, 설정파일명, 엑셀파일명):

    # 엑셀 파일 열기
    wb = load_workbook(filename = 추출할엑셀파일경로)

    # 각 시트를 순회하여 아파트목록 데이터 생성
    아파트목록 = [get_apartment_info(wb, sheetname, i+1) for i, sheetname in enumerate(wb.sheetnames)]

    # JSON 파일로 저장
    config = {
        '아파트목록': 아파트목록,
        '설정파일명': 설정파일명,
        '엑셀파일명': 엑셀파일명,
    }

    with open(설정파일명, 'w', encoding='utf-8') as json_file:
        json.dump(config, json_file, ensure_ascii=False, indent=4)

''' 
1. excel 파일 하나에 단지별로 시트 하나에 명부가 적혀있다.
2. 각 시트별로 순회하며 단지 정보를 추출하여 하나의 아파트 객체를 만든다.
3. 만들어진 아파트 객체 리스트를 json 형식으로 파일로 저장한다.
'''
if __name__ == '__main__':
    
    추출할엑셀파일경로 = '00 - 입주자서명부.xlsx'
    설정파일명 = 'apartments.json'
    엑셀파일명 = 'summary.xlsx'

    # 아파트 분석 실행
    analyze_apartments(추출할엑셀파일경로, 설정파일명, 엑셀파일명)
