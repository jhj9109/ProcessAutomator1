import sys # for 실행인자
import os # for 파일시스템 조회

from common_utils import get_xlsx_file_name, sorted_file_entries, load_json, get_apart_object, get_apart_object2

from data_processing_v1 import update_one_apartment

DEFAULT_CONFIG_FILE_PATH = "./apartments.json"

if __name__ == '__main__':
    
    config = load_json(DEFAULT_CONFIG_FILE_PATH)\
    
    # 0. 커맨드라인 인수로부터 정보 입력 받음
    folder_path = sys.argv[1]
    base_path = "." if len(sys.argv) < 3 else sys.argv[2]
    
    pass_files = []
    folder_entries = []
    for entry in os.scandir(folder_path):
        if entry.is_file():
            # 폴더만 순회할 것임. 파일은 기록용으로 출력만
            pass_files.append(entry.name)
        else:
            folder_entries.append(entry)
    
    # 각 폴더는 아파트 단지 하나 => 단지명
    for entry in folder_entries:
        try:
            순번 = int(entry.name[:2])
            # 단지명 = name[2:] # 앞에 두자리 제외 => 01서울번동3 => 서울번동3
            '''
            folder_path => 해당 폴더 아래의 모든 파일을 DirEntry로 추출할때 사용하기 위함.
            단지명 => 아파트 객체 & 해당 아파트 엑셀 파일 경로
            base_path => 해당 아파트 엑셀 파일 경로
            '''
            아파트객체 = get_apart_object2(config, 순번)
            print(f"{아파트객체['단지명']}단지 {아파트객체['대상세대수']}세대에 대한 작업 시작")
            update_one_apartment(config, entry.path, 아파트객체["단지명"], base_path)
        except:
            print(f"{entry.name}폴더명을 가진 작업분에 대해 실패")
    