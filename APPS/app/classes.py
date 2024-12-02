from pathlib import Path
from pandas import DataFrame
import pandas as pd
import os


class CapacityClass:
    def __init__(self):
        self.sheet_name = 'Capacity'
        self.bi_sheet_name = 'Capacity_BI'
        self.melted = ['Capacity', 'Cost_Center', 'Code', 'Account group', 'Versions', 'DATES', 'VALUE']
        self.start_column = 6 # 열 시작 컬럼
        self.column_value = 5
        self.result_path = os.path.join('3. RESULT', 'test.xlsx')

    def add_used_row(self, melted_df: DataFrame):
        melted_df['USED'] = melted_df.apply(
        lambda melted: (
            "X" if (
                (pd.isna(melted[self.melted[3]]) and isinstance(melted[self.melted[3]], float)) or
                not melted[self.melted[3]].startswith("5")
            ) else "O"
        ),
        axis=1
    )
        
    def set_df(self, df: DataFrame):
        df = df[~df.iloc[:, 3].astype(str).str.startswith('ZGR', na=False)]
        df = df[df.iloc[:, 2].astype(str).str.strip().str.len() == 10]


class PromotionClass:
    def __init__(self):
        self.sheet_name = 'Promotion'
        self.bi_sheet_name = 'Promotion_BI'
        self.melted = ['PROMOTION', 'CODE', 'CODE_NAME', 'PRODUCT_GROUP', 'PRODUCT_CODE', 'VERSION', 'DATES', 'VALUE']
        self.start_column = 7 # 열 시작 컬럼
        self.column_value = 6

    def add_used_row(self, melted_df: DataFrame):
            melted_df['USED'] = melted_df.apply(
            lambda melted: (
                "X" if (
                    "/" not in melted[self.melted[1]] or
                    melted[self.melted[2]] == "Admin Common" or
                    melted[self.melted[3]] == "All Product Groups" or
                    (pd.isna(melted[self.melted[4]]) and isinstance(melted[self.melted[4]], float)) or
                    not melted[self.melted[4]].startswith("5")
                ) else "O"
            ),
            axis=1
        )

class PcrClass:
    def __init__(self):
        self.sheet_name = 'PCR_Power'
        self.bi_sheet_name = 'PCR_POWERBI'
        self.melted = ['COA', 'Account Group','Function1#','Product Grp','Category4', 'DATES', 'VALUE', 'QETABLE']
        self.cycles = [
            {'range': range(11, 23), 'qetable': 'FY ACT 2023 @BUD rate'},
            {'range': range(24, 36), 'qetable': 'FY BUD 2024 @BUD rate'},
            {'range': range(37, 49), 'qetable': '24 QE1'},
            {'range': range(50, 62), 'qetable': '24 QE2'},
            {'range': range(63, 75), 'qetable': '24 QE3'},
            {'range': range(76, 88), 'qetable': '24 QE4'},
            {'range': range(89, 101), 'qetable': 'YTD ACT 2024 @BUD rate'},
            {'range': range(102,114 ), 'qetable': 'FY BUD 2025 @BUD rate'}
        ]
        self.column_value = 6

    def update_value_row(self, melted: DataFrame):
        melted['VALUE'] = melted['VALUE'].apply(lambda x: 0 if x == '-' else x)

folder_path = Path('./1. WORKING')
backup_folder = './2. BACKUP'
result_folder = './3. RESULT'