import pandas as pd

from openpyxl import load_workbook
from app.classes import *
from app.utils import *


def transform_to_pivot(set_cycle: int, file_path: str, sheet_name: str, bi_sheet_name: str):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    remove_sheets(file_path, bi_sheet_name)

    if set_cycle == 1:
        df = df[~df.iloc[:, 3].astype(str).str.startswith('ZGR', na=False)]
        df = df[df.iloc[:, 2].astype(str).str.strip().str.len() == 10]
        set_melted = CapacityClass.melted

    elif set_cycle == 2:
        set_melted = PromotionClass.melted

    elif set_cycle == 3:
        set_melted = PcrClass.melted

    all_data = data_cycles(set_cycle, df, set_melted)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        all_data.to_excel(writer, sheet_name=bi_sheet_name, index=False)

    workbook = load_workbook(file_path)
    sheet = workbook[bi_sheet_name]
    
    if set_cycle == 1:
        sheet['I1'] = 'USED'

    elif set_cycle == 2:
        sheet['J1'] = 'USED'

    elif set_cycle == 3:
        sheet['I1'] = 'Category5'
        sheet['J1'] = 'USED'

    workbook.save(file_path)
    workbook.close()

    return True