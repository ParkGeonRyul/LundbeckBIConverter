import pandas as pd
import os

from datetime import datetime
from pathlib import Path
from pandas import DataFrame


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
        lambda row: (
            "X" if (
                row[self.melted[3]] == "Result"
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
        self.start_column = 7
        self.column_value = 6

    def add_used_row(self, melted_df: DataFrame):
            melted_df['USED'] = melted_df.apply(
            lambda row: (
                "X" if (
                    "/" not in row[self.melted[1]] or
                    row[self.melted[2]] == "Admin Common" or
                    row[self.melted[3]] == "All Product Groups" or
                    (pd.isna(row[self.melted[4]]) and isinstance(row[self.melted[4]], float)) or
                    not row[self.melted[4]].startswith("5")
                ) else "O"
            ),
            axis=1
        )

class PcrClass:
    def __init__(self):

        self.sheet_name = 'PCR_Power'
        self.bi_sheet_name = 'PCR_POWERBI'
        self.melted = ['COA', 'Account Group','Function1#','Product Grp','Category4', 'DATES', 'VALUE']
        self.column_value = 5

    def cycles(self, year: int | None = None):
        
        if year is None:
            year = datetime.now().year
        short_year = str(year)[-2:]
        cycles = [
            {'range': range(11, 23), 'qetable': f'FY ACT {year - 1} @BUD rate'},
            {'range': range(24, 36), 'qetable': f'FY BUD {year - 1} @BUD rate'},
            {'range': range(37, 49), 'qetable': f'{short_year} QE1'},
            {'range': range(50, 62), 'qetable': f'{short_year} QE2'},
            {'range': range(63, 75), 'qetable': f'{short_year} QE3'},
            {'range': range(76, 88), 'qetable': f'{short_year} QE4'},
            {'range': range(89, 101), 'qetable': f'YTD ACT {year} @BUD rate'},
            {'range': range(102,114 ), 'qetable': f'FY BUD {year + 1} @BUD rate'}
        ]

        return cycles

    def update_row(self, melted: DataFrame):
        melted['VALUE'] = melted['VALUE'].apply(lambda x: 0 if x == '-' else x)

    def add_used_row(self, melted_df: DataFrame):
        melted_df['Category5'] = melted_df.apply(
            lambda row: (
                row[self.melted[0]] +
                (f"_{row[self.melted[1]]}" if not pd.isna(row[self.melted[1]]) else "") +
                (f"_{row[self.melted[2]]}" if not pd.isna(row[self.melted[2]]) else "") +
                (f"_{row[self.melted[3]]}" if not pd.isna(row[self.melted[3]]) else "") +
                (f"_{row[self.melted[4]]}" if not pd.isna(row[self.melted[4]]) else "")
            ),
            axis=1
        )

        melted_df['USED'] = melted_df['Category5'].apply(
            lambda row: "X" if row in [
                "Sales_Net Sales",
                "Sales_Gross Sales",
                "Sales_Sales Adjustments",
                "Production Cost",
                "Production Cost_Manufacuring Costs",
                "Production Cost_Other Variable Cost",
                "SG&A",
                "SG&A_Total Promotion Cost",
                "SG&A_Total Promotion Cost_Promotion Cost",
                "SG&A_Total Promotion Cost_Medical Affairs activity cost",
                "SG&A_Total Sales Cost",
                "Profit Centre Result_Profit Centre Result"
            ] else "O"
        )

        melted_df['USED'] = melted_df.apply(
            lambda row: "X" if (
                row['USED'] == "O" and
                (
                    row[self.melted[1]] == "Net Sales incl. other revenue" or
                    row[self.melted[2]] == "Pricing and Market Access" or
                    row[self.melted[3]] == "All Product Groups"
                )
            ) else row['USED'],
            axis=1
        )

folder_path = Path('./1.WORKING')
backup_folder = './2.BACKUP'
result_folder = './3.RESULT'