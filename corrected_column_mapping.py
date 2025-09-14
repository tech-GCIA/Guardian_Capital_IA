#!/usr/bin/env python
"""
Corrected column mapping based on actual Base Sheet analysis from row 6
This fixes the import-export alignment issues
"""

def get_corrected_import_column_mapping():
    """
    Returns the corrected column mapping for data import based on actual Base Sheet structure
    """
    return {
        'basic_info': {
            's_no': 0,           # S. No.
            'company_name': 1,   # Company Name  
            'accord_code': 2,    # Accord Code
            'sector': 3,         # Sector
            'cap': 4,            # Cap
            'free_float': 5,     # Free Float
            'revenue_6yr_cagr': 6, # 6 Year CAGR Revenue
            'revenue_ttm': 7,    # TTM Revenue
            'pat_6yr_cagr': 8,   # 6 Year CAGR PAT
            'pat_ttm': 9,        # TTM PAT
            'current': 10,       # Current
            'two_yr_avg': 11,    # 2 Yr Avg
            'reval_deval': 12,   # Reval/deval
        },
        
        # Time-series data sections based on Row 6 analysis
        'time_series': {
            # Market Cap dates: columns 14-41 (Row 6 shows Market Cap at col 15-16)
            'market_cap_start': 14,
            'market_cap_end': 41,
            
            # Market Cap Free Float: columns 43-70 (Row 6 shows Market Cap FF at col 44-48)
            'market_cap_ff_start': 43,
            'market_cap_ff_end': 70,
            
            # TTM Revenue: columns 72-107 (Row 6 shows TTM Revenue at col 73-75)  
            'ttm_revenue_start': 72,
            'ttm_revenue_end': 107,
            
            # TTM Revenue Free Float: columns 108-143 (Row 6 shows TTM Revenue FF at col 109-112)
            'ttm_revenue_ff_start': 108,
            'ttm_revenue_ff_end': 143,
            
            # TTM PAT: columns 144-179 (Row 6 shows TTM PAT at col 145-147)
            'ttm_pat_start': 144,
            'ttm_pat_end': 179,
            
            # TTM PAT Free Float: columns 180-215 (Row 6 shows TTM PAT FF at col 181-183)
            'ttm_pat_ff_start': 180,
            'ttm_pat_ff_end': 215,
            
            # Quarterly Revenue: columns 216-251 (Row 6 shows Quarterly Revenue at col 217-219)
            'qtr_revenue_start': 216,
            'qtr_revenue_end': 251,
            
            # Quarterly Revenue Free Float: columns 252-287 (Row 6 shows at col 253-255)
            'qtr_revenue_ff_start': 252,
            'qtr_revenue_ff_end': 287,
            
            # Quarterly PAT: columns 288-323 (Row 6 shows at col 289-291)
            'qtr_pat_start': 288,
            'qtr_pat_end': 323,
            
            # Quarterly PAT Free Float: columns 324-359 (Row 6 shows at col 325-327)
            'qtr_pat_ff_start': 324,
            'qtr_pat_ff_end': 359,
            
            # Annual ratios based on Row 6 analysis
            'roce_start': 360,    # Row 6 shows ROCE at col 361
            'roce_end': 372,      # 12 years of data
            
            'roe_start': 373,     # Row 6 shows ROE at col 374  
            'roe_end': 385,       # 12 years of data
            
            'retention_start': 386,  # Row 6 shows Retention at col 387
            'retention_end': 398,    # 12 years of data
            
            # Stock Price data
            'price_start': 399,   # Row 6 shows Share Price at col 400
            'price_end': 409,     # Various price dates
            
            'pe_start': 410,      # Row 6 shows PE at col 413 (approximate)
            'pe_end': 420,        # PE ratios
        },
        
        # Identifiers at the end
        'identifiers': {
            'bse_code_col': 422,   # Column position for BSE Code
            'nse_code_col': 423,   # Column position for NSE Symbol  
            'isin_col': 424,       # Column position for ISIN
        }
    }

def get_date_period_mappings():
    """
    Returns the exact date and period mappings used in the Base Sheet
    """
    return {
        'market_cap_dates': [
            '2025-08-19', '2025-06-30', '2025-03-28', '2025-01-31', '2024-12-31',
            '2024-09-30', '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29',
            '2023-06-30', '2023-03-31', '2022-12-30', '2022-09-30', '2022-06-30',
            '2022-03-31', '2021-12-31', '2021-09-30', '2021-06-30', '2021-03-31',
            '2020-12-31', '2020-09-30', '2020-06-30', '2020-03-31', '2019-12-31',
            '2019-09-30', '2019-06-28', '2019-04-01'
        ],
        
        'ttm_periods': [
            '202506', '202503', '202412', '202409', '202406', '202403',
            '202312', '202309', '202306', '202303', '202212', '202209',
            '202206', '202203', '202112', '202109', '202106', '202103',
            '202012', '202009', '202006', '202003', '201912', '201909',
            '201906', '201903', '201812', '201809', '201806', '201803',
            '201712', '201709', '201706', '201703', '201612'
        ],
        
        'annual_years': [
            '2024-25', '2023-24', '2022-23', '2021-22', '2020-21', '2019-20',
            '2018-19', '2017-18', '2016-17', '2015-16', '2014-15', '2013-14'
        ],
        
        'price_dates': [
            '2025-08-19', '2025-06-30', '2025-03-28', '2024-12-31', '2024-09-30',
            '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29', '2023-06-30'
        ]
    }

def get_corrected_header_structure():
    """
    Returns corrected header structure based on actual Base Sheet rows 6-8 analysis
    """
    
    # Initialize all rows with 426 empty columns
    headers = {
        'row_1': [''] * 426,  # Row 1: Empty (before multi-level headers)
        'row_2': [''] * 426,  # Row 2: Empty  
        'row_3': [''] * 426,  # Row 3: Empty
        'row_4': [''] * 426,  # Row 4: Empty
        'row_5': [''] * 426,  # Row 5: Empty
        'row_6': [''] * 426,  # Row 6: Category descriptions (actual first header row)
        'row_7': [''] * 426,  # Row 7: Sub-category details  
        'row_8': [''] * 426,  # Row 8: Actual column headers
    }
    
    # ROW 6: Category descriptions (based on analysis)
    row6_mapping = {
        6: 'Stock wise Fundamentals and Valuations',
        14: 'Market Cap (in crores)',
        15: 'Market Cap (in crores)',
        43: 'Market Cap- Free Float (in crores)',
        44: 'Market Cap- Free Float (in crores)',
        45: 'Market Cap- Free Float (in crores)',
        46: 'Market Cap- Free Float (in crores)',
        47: 'Market Cap- Free Float (in crores)',
        72: 'TTM Revenue',
        73: 'TTM Revenue',
        74: 'TTM Revenue',
        108: 'TTM Revenue- Free Float',
        109: 'TTM Revenue- Free Float',
        110: 'TTM Revenue- Free Float',
        111: 'TTM Revenue- Free Float',
        144: 'TTM PAT',
        145: 'TTM PAT',
        146: 'TTM PAT',
        180: 'TTM PAT- Free Float',
        181: 'TTM PAT- Free Float',
        182: 'TTM PAT- Free Float',
        216: 'Quarterly- Revenue',
        217: 'Quarterly- Revenue',
        218: 'Quarterly- Revenue',
        252: 'Quarterly- Revenue- Free Float',
        253: 'Quarterly- Revenue- Free Float',
        254: 'Quarterly- Revenue- Free Float',
        288: 'Quarterly- PAT',
        289: 'Quarterly- PAT',
        290: 'Quarterly- PAT',
        324: 'Quarterly-PAT- Free Float',
        325: 'Quarterly-PAT- Free Float',
        326: 'Quarterly-PAT- Free Float',
        360: 'ROCE (%)',
        373: 'ROE (%)',
        386: 'Retention (%)',
        399: 'Share Price'
    }
    
    for col, value in row6_mapping.items():
        headers['row_6'][col] = value
    
    # ROW 7: Sub-category details (based on analysis)
    row7_mapping = {
        6: 'Revenue (Q1 FY-26)',
        8: 'PAT (Q1 FY-26)',
        10: 'PR',
        14: 'Market Cap (in crores)',
        15: 'Market Cap (in crores)',
        43: 'Market Cap- Free Float (in crores)',
        72: 'TTM Revenue',
        108: 'TTM Revenue- Free Float',
        144: 'TTM PAT',
        180: 'TTM PAT- Free Float',
        216: 'Quarterly- Revenue',
        252: 'Quarterly- Revenue- Free Float',
        288: 'Quarterly- PAT',
        324: 'Quarterly-PAT- Free Float',
        360: 'ROCE (%)',
        373: 'ROE (%)',
        386: 'Retention (%)',
        399: 'Share Price',
        401: 'PR',
        412: 'PE'
    }
    
    for col, value in row7_mapping.items():
        headers['row_7'][col] = value
    
    return headers