import pandas as pd

from openpyxl import load_workbook
from app.classes import *
from app.utils import *


def transform_to_pivot(file_path: str, classes: classmethod, sheet_name: str, bi_sheet_name: str):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    remove_sheets(file_path, bi_sheet_name)

    if classes.sheet_name == CapacityClass().sheet_name:
        df = df[~df.iloc[:, 3].astype(str).str.startswith('ZGR', na=False)]
        df = df[df.iloc[:, 2].astype(str).str.strip().str.len() == 10]
        set_melted = classes.melted

    else:
        set_melted = classes.melted

    all_data = data_cycles(sheet_name, df, set_melted)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        all_data.to_excel(writer, sheet_name=bi_sheet_name, index=False)

    workbook = load_workbook(file_path)
    sheet = workbook[bi_sheet_name]
    
    if sheet_name == CapacityClass().sheet_name:
        sheet['I1'] = 'USED'

    elif sheet_name == PromotionClass().sheet_name:
        sheet['J1'] = 'USED'

    elif sheet_name == PcrClass().sheet_name:
        sheet['I1'] = 'Category5'
        sheet['J1'] = 'USED'

    workbook.save(file_path)
    workbook.close()