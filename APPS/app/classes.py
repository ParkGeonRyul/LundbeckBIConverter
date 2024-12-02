from pathlib import Path

class CapacityClass:
    def __init__(self):
        self.sheet_name = 'Capacity'
        self.bi_sheet_name = 'Capacity_BI'
        self.melted = ['Capacity', 'Cost_Center', 'Code', 'Account group', 'Versions', 'DATES', 'VALUE']


class PromotionClass:
    def __init__(self):
        self.sheet_name = 'Promotion'
        self.bi_sheet_name = 'Promotion_BI'
        self.melted = ['PROMOTION', 'CODE', 'CODE_NAME', 'PRODUCT_GROUP', 'PRODUCT_CODE', 'VERSION', 'DATES', 'VALUE']

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

folder_path = Path('./1. WORKING')
backup_folder = './2. BACKUP'