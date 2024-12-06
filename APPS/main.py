from app.classes import *
from app.utils import *
from app.controller import *
from tqdm import tqdm


def start(): # 작업 시작
    TransformClass().terminate_excel()

    file_list = [f.name for f in folder_path.iterdir() if f.is_file()] # 1. WORKING 폴더 내에 파일 감지
    file_datas = ["QE", "PCR"]
    set_class = [CapacityClass(), PromotionClass(), PcrClass()]
    filtered_list = []

    for item in file_datas:
        filtered_list.append([f for f in file_list if item in f]) # QE, PCR로 파일 필터링 후 각개로 작동하게 로직 구현

    print("\n ※ 주의: 작업 진행 도중 프로그램이 종료되면 엑셀 파일이 손상될 수 있습니다. \n")

    total_count = len(filtered_list[0]) * 2 + len(filtered_list[1]) # 작업 상황 관련 총 횟수
    complete_count = 0

    with tqdm(total=total_count, desc="BI 변환 작업 진행 상황", bar_format="{l_bar}{bar} [변환 {n_fmt}/{total_fmt}]") as pbar: # 변환 완료 횟수 퍼센테이지 게이지로 활성화
        for classes in set_class: # app.classes 내부 파일 class 변수 작업
            if classes.bi_sheet_name == PcrClass().bi_sheet_name:
                for items in filtered_list[1]: # PCR파일 작업
                    TransformClass.transform_excel(items, folder_path, classes)
                    complete_count += 1
                    pbar.update(1)
            
            else:
                for items in filtered_list[0]: # QE파일 작업
                    TransformClass.transform_excel(items, folder_path, classes)

                    complete_count += 1
                    pbar.update(1)

    print(f"총 횟수: {total_count}, 변환 성공 갯수: {complete_count}")

start()