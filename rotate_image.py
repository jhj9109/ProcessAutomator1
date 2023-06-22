from PIL import Image as PILImage

from image_utils import get_orientation, get_degree_by_orientaion

def rotate_image(image_path):

    with PILImage.open(image_path) as origin:
        
        orientation = get_orientation(origin)

        if orientation is None:
            return None

        try:
            degree = get_degree_by_orientaion(orientation)
        except:
            print(f"get_degree_by_orientaion 실패 => {image_path}, 오리엔테이션:{orientation}")
            return None

        if degree == 0:
            return None

        rotated_image = origin.rotate(degree, expand=True)
        
        return rotated_image

def save_rotated_image(image_path):
    try:
        rotated_image = rotate_image(image_path)

        if rotated_image is not None:
            
            rotated_image.save(image_path)
    except:
        print(f"save_rotated_image 실패: {image_path}")

# def get_temp_file_name(file_path):
#     # 파일 경로와 파일 이름, 확장자 분리
#     dir_path, file_name = os.path.split(file_path)
#     file_name_without_extension, extension = os.path.splitext(file_name)

#     # 수정된 파일 이름 생성
#     new_file_name = file_name_without_extension + "_temp" + extension

#     # 새로운 파일 경로 생성
#     new_file_path = os.path.join("temp", new_file_name)

#     print(new_file_path)

#     return new_file_path
