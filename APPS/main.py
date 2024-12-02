from app.classes import *
from app.utils import *
from app.controller import *
from tqdm import tqdm


def get_integer_input(prompt):
    while True:
        user_input = input(prompt)
        
        if not user_input or not user_input.isdigit():
            print("값이 입력되지 않았습니다.")
            
        else:
            number = int(user_input)

            if number > 3:
                print("값이 올바르지 않습니다.")
            else:
                return number

def start():
    file_list = [f.name for f in folder_path.iterdir() if f.is_file()]

    set_cycle = get_integer_input('Capacity는 1, Promotion은 2, PCR은 3입니다. 입력하여 주십시오: ')

    if set_cycle == 1:
        file_data = "QE"
        set_class = CapacityClass()

    elif set_cycle == 2:
        file_data = "QE"
        set_class = PromotionClass()

    elif set_cycle == 3:
        file_data = "PCR"
        set_class = PcrClass()

    filtered_list = [item for item in file_list if file_data in item]


    print("")
    print("※ 주의: 작업 진행 도중 프로그램이 종료되면 엑셀 파일이 손상될 수 있습니다.")
    print("")

    total_files = len(filtered_list)
        
    complete_count = 0

    with tqdm(total=total_files, desc="BI 변환 작업 진행 상황", bar_format="{l_bar}{bar} [파일 {n_fmt}/{total_fmt}]") as pbar:
        for items in filtered_list:
            file_path = f'{folder_path}/{items}'
            complete = transform_to_pivot(file_path, set_class, set_class.sheet_name, set_class.bi_sheet_name)

            if complete == True:
                complete_count += 1

            pbar.update(1)

    print(f"총 파일 갯수: {len(filtered_list)}, 변환 성공 갯수: {complete_count}")

start()