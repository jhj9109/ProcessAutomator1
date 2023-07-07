from PIL import Image as PILImage

from image_utils import get_orientation, get_degree_by_orientaion, is_mpo

def rotate_image(image_path):

    with PILImage.open(image_path) as origin:

        orientation = get_orientation(origin)

        if orientation is None:
            return None

        try:
            
            degree = get_degree_by_orientaion(orientation)

        except Exception as error:
            
            print(f"get_degree_by_orientaion 실패 => {image_path}, 오리엔테이션:{orientation}, error:{error}")
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

    except Exception as error:

        print(f"{error}: save_rotated_image 실패: {image_path}")

def mpo_to_jpeg(image_path):
    
    try:
            
        with PILImage.open(image_path) as origin:

            if is_mpo(origin):
                origin.save(image_path, 'JPEG')
    
    except Exception as error:

        print(f"{error}: mpo to jpeg에 실패했습니다.")
        
'''
이미지를 적절한 처리후 다시 저장하는 과정으로서
정상적인 경우에 Exception이 발생하지 않았으며,
혹시나 발생하더라도 코드를 멈추게 하지 않기 위해 try except 처리
'''
def handle_rotated_or_mpo_image(image_path):
    save_rotated_image(image_path)
    mpo_to_jpeg(image_path)
