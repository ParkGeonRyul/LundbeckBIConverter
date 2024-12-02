from pathlib import Path
import os

FOLDER_PATH = os.getenv('FOLDER_PATH')
BACKUP_PATH = os.getenv('BACKUP_PATH')

class CapacityClass:
    sheet_name = 'Capacity'
    bi_sheet_name = 'Capacity_BI'
    melted = ['Capacity', 'Cost_Center', 'Code', 'Account group', 'Versions', 'DATES', 'VALUE']

class PromotionClass:
    sheet_name = 'Promotion'
    bi_sheet_name = 'Promotion_BI'
    melted = ['PROMOTION', 'CODE', 'CODE_NAME', 'PRODUCT_GROUP', 'PRODUCT_CODE', 'VERSION', 'DATES', 'VALUE']

class PcrClass:
    sheet_name = 'PCR_Power'
    bi_sheet_name = 'PCR_POWERBI'
    melted = ['COA', 'Account Group','Function1#','Product Grp','Category4', 'DATES', 'VALUE', 'QETABLE']
    cycles = [
        {'range': range(11, 23), 'qetable': 'FY ACT 2023 @BUD rate'},
        {'range': range(24, 36), 'qetable': 'FY BUD 2024 @BUD rate'},
        {'range': range(37, 49), 'qetable': '24 QE1'},
        {'range': range(50, 62), 'qetable': '24 QE2'},
        {'range': range(63, 75), 'qetable': '24 QE3'},
        {'range': range(76, 88), 'qetable': '24 QE4'},
        {'range': range(89, 101), 'qetable': 'YTD ACT 2024 @BUD rate'},
        {'range': range(102,114 ), 'qetable': 'FY BUD 2025 @BUD rate'}
    ]

folder_path = Path('./1.WORKING')
backup_folder = './2.BACKUP'