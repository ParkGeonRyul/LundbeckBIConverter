import pandas as pd
import shutil
import os

from openpyxl import load_workbook
from datetime import datetime
from app.classes import *


def back_up(file_path: str): # 백업파일 생성 로직
    current_datetime = datetime.now().strftime('%Y%m%d%H%M') # 년, 월, 일, 시, 분에 맞춰 이름 생성
    base_file_name = os.path.basename(file_path)
    backup_file_name = f"{os.path.splitext(base_file_name)[0]}_{current_datetime}.xlsx" # 백업파일 이름 지정
    backup_file_path = os.path.join(backup_folder, backup_file_name)
    
    if not os.path.exists(backup_folder): # 백업 폴더 없을 경우 생성
        os.makedirs(backup_folder)
    
    shutil.copy(file_path, backup_file_path) # 원본파일 복사 후 백업파일 생성

def remove_sheets(file_path: str, bi_sheet_name: str): # Sheet 삭제 로직
    workbook = load_workbook(file_path) # Excel 파일 불러오기

    if bi_sheet_name in workbook.sheetnames:
        std = workbook[bi_sheet_name]
        workbook.remove(std) # 해당 시트 삭제 후 저장
        workbook.save(file_path)
    workbook.close()