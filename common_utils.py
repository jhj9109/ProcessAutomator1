def get_worksheet_name(단지명, 몇동):
    return f"{단지명}-{몇동}동"

def get_xlsx_file_name(단지명):
    return f"{단지명}_요약"

def sorted_files(depth, folder_path):
    import os
    
    files = []
    
    for entry in os.scandir(folder_path):
        if entry.is_file():
            # 파일인 경우 파일 목록에 추가
            files.append(entry.name)
        elif entry.is_dir() and depth > 0:
            # 폴더인 경우 && depth가 남아있다면, 재귀 호출을 통해 파일 목록 추가
            sorted_subfolder_files = sorted_files(depth - 1, entry.path)
            files.extend(sorted_subfolder_files)

    files.sort() # 정렬

    return files

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