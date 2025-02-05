import pandas as pd
import time
import psutil

from openpyxl import load_workbook
from app.classes import *
from app.utils import *


class TransformClass: # excel 파일 경로 설정 및 백업 생성 class화
    def transform_excel(items: str, folder_path: str, classes: classmethod):
        file_path = f'{folder_path}/{items}'
        split_item= items.split(" ")
        
        try:
            file_year = int(split_item[0])

        except:
            file_year = None

        back_up(file_path)        
        result = transform_to_pivot(file_path, classes, classes.sheet_name, classes.bi_sheet_name, file_year)

        return result
    
    def terminate_excel(self): # 엑셀 강제종료
        for process in psutil.process_iter(['name']):
            if process.info['name'] == 'EXCEL.EXE':
                process.terminate()
                print("\n 현재 열린 모든 Excel창을 강제로 종료했습니다. \n")
                time.sleep(1)


def transform_to_pivot(file_path: str, classes: classmethod, sheet_name: str, bi_sheet_name: str, file_year: int | None = None): # 변환 Controller
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
    except ValueError as e:
        if f"Worksheet named '{sheet_name}' not found" in str(e):

            return False

    remove_sheets(file_path, bi_sheet_name) # Excel 시트에 bi_sheet_name으로 되어있는 Sheet 삭제

    if hasattr(classes, 'set_df'):
        classes.set_df(df)

    set_melted = classes.melted

    all_data = data_cycles(classes, sheet_name, df, set_melted, file_year) # Sheet 변환 사이클 돌리기

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer: # classes.py 에 있는 bi_sheet_name 이름으로 새 Sheet 작성
        all_data.to_excel(writer, sheet_name=bi_sheet_name, index=False)

    workbook = load_workbook(file_path)

    workbook.save(file_path) # Excel 저장
    workbook.close() # Excel 닫기

    return True

def data_cycles(classes: classmethod, sheet_name: str, df: pd.DataFrame, melted_column: list, file_year: int | None = None):
    all_data = pd.DataFrame() # 빈 데이터프레임 생성

    if hasattr(classes, 'start_column'):
        start_column = classes.start_column # 열 시작 컬럼
    
    column_value = classes.column_value

    fixed_cols = df.columns[:column_value] # 고정 컬럼 지정
    cycles = []

    if sheet_name != PcrClass().sheet_name:

        for i in range(4): # 고정컬럼 이후 컬럼 범위 생성 반복문
            increase_count = start_column + (i * 13)
            range_count = {'range': range(increase_count, increase_count + 12), 'qetable': f'QE{i + 1}'}
            cycles.append(range_count)

    else:
        cycles = classes.cycles(file_year) # PCR은 열 이름이 계속 달라져서 하드코딩(추후 변경 예정)

    for cycle in cycles:
        cycle_range = cycle['range'] # 컬럼 범위 불러오기
        qetable_value = cycle['qetable'] # qe테이블 값 변수저장
        date_cols = df.columns[cycle_range] # 날짜 컬럼 생성
        melted = df.melt(id_vars=fixed_cols, value_vars=date_cols, var_name='DATES', value_name='VALUE')

        if hasattr(classes, 'update_row'): # PCR일 경우 데이터 중 -값을 0으로 치환
            classes.update_row(melted)

        melted['VALUE'] = pd.to_numeric(melted['VALUE'], errors='coerce').fillna(0).astype(int) # VALUE값 중 Null 값 0으로 치환
        melted.columns = melted_column
        melted['QETABLE'] = qetable_value # QETABLE 열의 행 값 설정
        classes.add_used_row(melted) # USED 열 추가 및 행 값 추가
        
        all_data = pd.concat([all_data, melted], ignore_index=True)
    
    return all_data
