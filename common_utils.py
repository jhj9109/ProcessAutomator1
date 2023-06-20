def get_worksheet_name(단지명, 몇동):
    return f"{단지명}-{몇동}동"

def get_xlsx_file_name(단지명):
    return f"{단지명}_요약.xlsx"

def sorted_files(depth, folder_path):
    import os
    
    files = []
    
    for entry in os.scandir(folder_path):
        if entry.is_file():
            # 파일인 경우 파일 목록에 추가
            files.append(entry.name)
        elif entry.is_dir() and depth != 0:
            # 폴더인 경우 && depth가 남아있다면, 재귀 호출을 통해 파일 목록 추가
            if depth == -1: # 무한
                sorted_subfolder_files = sorted_files(-1, entry.path)
            else:    
                sorted_subfolder_files = sorted_files(depth - 1, entry.path)
            files.extend(sorted_subfolder_files)

    files.sort() # 정렬

    return files

def sorted_file_entries(depth, folder_path):
    import os
    
    entries = []
    
    for entry in os.scandir(folder_path):
        if entry.is_file():
            # 파일인 경우 파일 목록에 추가
            entries.append(entry)
        elif entry.is_dir() and depth != 0:
            # 폴더인 경우 && depth가 남아있다면, 재귀 호출을 통해 파일 목록 추가
            if depth == -1: # 무한
                sorted_subfolder_entries = sorted_file_entries(-1, entry.path)
            else:    
                sorted_subfolder_entries = sorted_file_entries(depth - 1, entry.path)
            entries.extend(sorted_subfolder_entries)

    entries.sort(key=lambda entry: entry.name) # 정렬

    return entries

def get_config_from_json(filepath):
    import json
    try:
        with open(filepath, 'r') as f:
            config = json.load(f)
    except FileNotFoundError:
        raise Exception(f"파일 '{filepath}'을(를) 찾을 수 없습니다.")
    except json.JSONDecodeError:
        raise Exception(f"파일 '{filepath}'의 JSON 형식이 올바르지 않습니다.")

    return config

def get_apart_object(config, 단지명):
    for 아파트객체 in config["아파트목록"]:
        if 단지명 == 아파트객체["단지명"]:
            return 아파트객체
    raise Exception(f"config에서 단지명({단지명})에 대한 아파트객체_추출에 실패했습니다.")

def get_apart_object2(config, 순번):
    for 아파트객체 in config["아파트목록"]:
        if 순번 == 아파트객체["순번"]:
            return 아파트객체
    raise Exception(f"config에서 순번({순번})에 대한 아파트객체_추출에 실패했습니다.")