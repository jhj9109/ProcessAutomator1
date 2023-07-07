from collections import defaultdict

from constants import ORIENTATION_TAG, ori_kr, ori_to_degree, 현관사진_COLUMN, 큐알사진_COLUMN, FIRST_ITEM_ROW_INDEX


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


def is_valid_anchor(row, col):
    return row >= FIRST_ITEM_ROW_INDEX and \
        (row - FIRST_ITEM_ROW_INDEX) % 2 == 0 and \
        (col != 현관사진_COLUMN and col != 큐알사진_COLUMN)


def is_pair(images_in_row):
    return 현관사진_COLUMN in images_in_row and 큐알사진_COLUMN in images_in_row


def get_image_records(ws, raise_exception_flag=True):

    image_records = defaultdict(dict)

    try:

        for index_in__image, img in enumerate(ws._images):
            '''
            ws.cell(row, col) 과정에서 row와 col은 1부터 시작하는값.
            img.anchor._from 의 row, col은 0부터 시작하는 값.
            '''
            ROW_OFFSET = 1
            COLUMN_OFFSET = 1

            row = img.anchor._from.row + ROW_OFFSET
            col = img.anchor._from.col + COLUMN_OFFSET

            if not is_valid_anchor(row, col):  # 이미지가 위치가능한 좌표가 있음
                raise Exception(
                    f"Invalid image anchor's (row, col): ({row, col})")

            if col in image_records[row]:  # 하나의 셀에 여러 이미지 앵커가 가르킴
                raise Exception(
                    f"Duplicated image anchor at ({row, col}): {image_records[row][col], index_in__image}")

            image_records[row][col] = index_in__image

        for row, images_in_row in image_records.items():
            if not is_pair(images_in_row):
                raise Exception(
                    f"짝이 안맞음 row, images_in_row: {row, images_in_row}")

    except Exception as error:

        if raise_exception_flag:
            raise (error)
        print(f"get_image_records()에서 Exception 발생")
        print(error)

    return image_records


def is_mpo(image):
    return image.format.upper() == 'MPO'

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
