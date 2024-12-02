from app.classes import *
from app.utils import *
from app.controller import *
from tqdm import tqdm


class TransformClass:
    def transform_excel(items: str, folder_path: str, classes: classmethod):
        file_path = f'{folder_path}/{items}'
        back_up(file_path)
        result = transform_to_pivot(file_path, classes, classes.sheet_name, classes.bi_sheet_name)

        return result

def start():
    file_list = [f.name for f in folder_path.iterdir() if f.is_file()]
    file_datas = ["QE", "PCR"]
    set_class = [CapacityClass(), PromotionClass(), PcrClass()]
    filtered_list = []

    for item in file_datas:
        filtered_list.append([f for f in file_list if item in f])

    print("")
    print("※ 주의: 작업 진행 도중 프로그램이 종료되면 엑셀 파일이 손상될 수 있습니다.")
    print("")

    total_files = len(filtered_list[0]) * 2 + len(filtered_list[1])
    complete_count = 0

    with tqdm(total=total_files, desc="BI 변환 작업 진행 상황", bar_format="{l_bar}{bar} [파일 {n_fmt}/{total_fmt}]") as pbar:
        for classes in set_class:
            if classes.bi_sheet_name != PcrClass().bi_sheet_name:
                for items in filtered_list[0]:
                    result = TransformClass.transform_excel(items, folder_path, classes)

                    if result == True:
                        complete_count += 1
                        pbar.update(1)
            else:
                for items in filtered_list[1]:
                    result = TransformClass.transform_excel(items, folder_path, classes)

                    if result == True:
                        complete_count += 1
                        pbar.update(1)

    print(f"총 파일 갯수: {total_files}, 변환 성공 갯수: {complete_count}")

start()