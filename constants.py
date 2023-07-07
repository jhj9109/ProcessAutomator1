ORIENTATION_TAG = 0x0112

ori_kr = {
    1: "회전 없음",
    3: "180",
    6: "시계 90",
    8: "반시계 90",
}

ori_to_degree = {
    1: 0,
    3: 180,
    6: -90,
    8: 90,
}


''' constant values
현재 구조에서 이미지는 3번째 행부터 홀수 행에서 1, 3열에 위치한다.
- (행 - 첫번째아이템행)이 짝수여야 한다.
- 열은 1, 3열이여야 한다.
'''
현관사진_COLUMN = 1
큐알사진_COLUMN = 3
FIRST_ITEM_ROW_INDEX = 3
SUMMARY_FIRST_ITEM_ROW_INDEX = 2

단지명_to_지역구 = {
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

PX_TO_PT = 3 / 4
PT_TO_PX = 4 / 3 # 1.33

# Cell 한칸의 비율은 세로:가로 = 160:120 = 4:3
IMAGE_CELL_HEIGHT_PT = 160
IMAGE_CELL_HEIGHT_PX = IMAGE_CELL_HEIGHT_PT * PT_TO_PX
IMAGE_CELL_WIDTH_PX = IMAGE_CELL_HEIGHT_PX * 3 / 4

IMAGE_CELL_WIDTH_PT = 120
IMAGE_CELL_WIDTH_PT_6 = IMAGE_CELL_WIDTH_PT / 6
