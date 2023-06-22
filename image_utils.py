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

def get_orientation(pil_image):
    try:
        return pil_image._getexif().get(ORIENTATION_TAG)
    except:
        return None

def get_degree_by_orientaion(orientation):
    return ori_to_degree[orientation]

def print_image_info(image_path):

    from PIL import Image as PilImage

    img = PilImage.open(image_path)
    
    print(f"{image_path}: 크기: {img.size}")
    
    orientation = get_orientation(img)
    
    if orientation is not None:

        print(f"이미지 회전: {orientation} ({ori_kr[orientation]})")

# def reset_rotation_info(image):
#     # 이미지의 exif 정보 가져오기
#     print(image)
#     exif = image.info.get('exif')
#     print(exif)

#     if exif is None:
#         return image

#     # exif 정보를 디코딩하여 태그 이름으로 변경
#     exif_dict = {TAGS[key]: exif[key] for key in exif.keys(
#     ) if key in TAGS and TAGS[key] == 'Orientation'}

#     # 회전 정보를 리셋
#     if 'Orientation' in exif_dict:
#         exif_dict['Orientation'] = 1

#     # 디코딩된 exif 정보를 다시 인코딩하여 이미지에 설정
#     new_exif = {TAGS[key]: exif_dict[key] for key in exif_dict}
#     image.info['exif'] = new_exif

#     return image