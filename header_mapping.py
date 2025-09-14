#!/usr/bin/env python
"""
Comprehensive header mapping based on original Base Sheet analysis
This module provides the exact header structure for recreating the Base Sheet format
"""

def get_complete_header_structure():
    """
    Returns the complete 8-row header structure matching the original Base Sheet.xlsx
    Based on detailed analysis of the original file structure.
    """
    
    # Initialize all rows with 426 empty columns
    headers = {
        'row_1': [''] * 426,  # Row 1: All empty
        'row_2': [''] * 426,  # Row 2: Column numbers 1-425
        'row_3': [''] * 426,  # Row 3: Field descriptions with API names
        'row_4': [''] * 426,  # Row 4: Shareholding info and L/L-1 sequences
        'row_5': [''] * 426,  # Row 5: Formulas and calculations
        'row_6': [''] * 426,  # Row 6: Category descriptions
        'row_7': [''] * 426,  # Row 7: Sub-category details
        'row_8': [''] * 426,  # Row 8: Actual column headers
    }
    
    # ROW 2: Column numbers (0-based indexing, so col 1 = index 1)
    for i in range(1, 426):
        headers['row_2'][i] = str(i)
    
    # ROW 3: Field descriptions with API field names (0-based indexing)
    row3_mapping = {
        21: 'Market Cap',
        22: 'NDP_MCAP',
        77: 'Net Sales', 
        78: 'TTM_TTMNS',
        149: 'Profit After Tax',
        150: 'TTM_TTMNP',
        221: 'Net Sales & Other Operating Income',
        222: 'QTR_NET_SALES',
        293: 'Profit after tax',
        294: 'QTR_PAT',
        361: 'ROCE (%)',
        362: 'FR_ROCE',
        373: 'ROE (%)',
        374: 'FR_ROE',
        386: 'Dividend Pay Out Ratio(%)',
        388: 'FR_DIVIDEND_PAYOUT_PER',
        399: 'NDP_ADJCLOSE',
        402: 'NDP_ADJCLOSE',
        403: 'NDP_ADJCLOSE',
        423: 'BSE Code',
        424: 'CD_SCRIPCODE'
    }
    for col, value in row3_mapping.items():
        headers['row_3'][col] = value
    
    # ROW 4: Shareholding info and sequences
    row4_mapping = {
        21: 'Total Public Shareholding',
        22: 'SHP_TPTOTALPUBLIC',
        77: 'Other Income',
        78: 'TTM_TTM_OTHERINC',
        423: 'NSE Symbol',
        424: 'CD_SYMBOL'
    }
    for col, value in row4_mapping.items():
        headers['row_4'][col] = value
    
    # Add L/L-1/L-2... sequences for ROCE, ROE, Retention sections
    l_sequences = {
        360: 'L',      # ROCE section
        361: 'L-1', 362: 'L-2', 363: 'L-3', 364: 'L-4', 365: 'L-5',
        366: 'L-6', 367: 'L-7', 368: 'L-8', 369: 'L-9', 370: 'L-10', 371: 'L-11',
        373: 'L',      # ROE section  
        374: 'L-1', 375: 'L-2', 376: 'L-3', 377: 'L-4', 378: 'L-5',
        379: 'L-6', 380: 'L-7', 381: 'L-8', 382: 'L-9', 383: 'L-10', 384: 'L-11',
        386: 'L',      # Retention section
        387: 'L-1', 388: 'L-2', 389: 'L-3', 390: 'L-4', 391: 'L-5',
        392: 'L-6', 393: 'L-7', 394: 'L-8', 395: 'L-9', 396: 'L-10', 397: 'L-11'
    }
    for col, value in l_sequences.items():
        headers['row_4'][col] = value
    
    # ROW 5: ISIN info and formulas (simplified for export)
    headers['row_5'][423] = 'ISIN No'
    headers['row_5'][424] = 'CD_ISIN'
    
    # ROW 6: Category descriptions with exact positioning
    row6_mapping = {
        6: 'Stock wise Fundamentals and Valuations',
        14: 'Market Cap (in crores)',
        15: 'Market Cap (in crores)',
        43: 'Market Cap- Free Float  (in crores)',
        44: 'Market Cap- Free Float  (in crores)',
        45: 'Market Cap- Free Float  (in crores)',
        46: 'Market Cap- Free Float  (in crores)',
        47: 'Market Cap- Free Float  (in crores)',
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
        183: 'TTM PAT- Free Float',
        216: 'Quarterly- Revenue',
        217: 'Quarterly- Revenue',
        218: 'Quarterly- Revenue',
        252: 'Quarterly- Revenue-  Free Float',
        253: 'Quarterly- Revenue-  Free Float',
        254: 'Quarterly- Revenue-  Free Float',
        288: 'Quarterly- PAT',
        289: 'Quarterly- PAT',
        290: 'Quarterly- PAT',
        324: 'Quarterly-PAT-  Free Float',
        325: 'Quarterly-PAT-  Free Float',
        326: 'Quarterly-PAT-  Free Float',
        360: 'ROCE (%)',
        373: 'ROE (%)',
        386: 'Retention (%)',
        399: 'Share Price'
    }
    for col, value in row6_mapping.items():
        headers['row_6'][col] = value
    
    # ROW 7: Sub-category details
    row7_mapping = {
        6: 'Revenue (Q1 FY-26)',
        8: 'PAT (Q1 FY-26)',
        10: 'PR',
        14: 'Market Cap (in crores)',
        15: 'Market Cap (in crores)',
        43: 'Market Cap- Free Float  (in crores)',
        72: 'TTM Revenue',
        108: 'TTM Revenue- Free Float',
        144: 'TTM PAT',
        180: 'TTM PAT- Free Float',
        216: 'Quarterly- Revenue',
        252: 'Quarterly- Revenue-  Free Float',
        288: 'Quarterly- PAT',
        324: 'Quarterly-PAT-  Free Float',
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

def get_data_column_mapping():
    """
    Returns the mapping for actual data columns (Row 8) based on analysis
    """
    # Basic info columns
    basic_columns = [
        'S. No.',           # Col 1
        'Company Name',     # Col 2
        'Accord Code',      # Col 3
        'Sector',          # Col 4
        'Cap',             # Col 5
        'Free Float',      # Col 6
        '6 Year CAGR',     # Col 7
        'TTM',             # Col 8
        '6 Year CAGR',     # Col 9
        'TTM',             # Col 10
        'Current',         # Col 11
        '2 Yr Avg',        # Col 12
        'Reval/deval',     # Col 13
        '',                # Col 14 - separator
    ]
    
    # Market cap dates (columns 15-42)
    market_cap_dates = [
        '2025-08-19', '2025-06-30', '2025-03-28', '2025-01-31', '2024-12-31',
        '2024-09-30', '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29',
        '2023-06-30', '2023-03-31', '2022-12-30', '2022-09-30', '2022-06-30',
        '2022-03-31', '2021-12-31', '2021-09-30', '2021-06-30', '2021-03-31',
        '2020-12-31', '2020-09-30', '2020-06-30', '2020-03-31', '2019-12-31',
        '2019-09-30', '2019-06-28', '2019-04-01'
    ]
    
    # TTM periods for various sections
    ttm_periods = [
        '202506', '202503', '202412', '202409', '202406', '202403',
        '202312', '202309', '202306', '202303', '202212', '202209',
        '202206', '202203', '202112', '202109', '202106', '202103',
        '202012', '202009', '202006', '202003', '201912', '201909',
        '201906', '201903', '201812', '201809', '201806', '201803',
        '201712', '201709', '201706', '201703', '201612'
    ]
    
    # Annual years for ratios
    annual_years = [
        '2024-25', '2023-24', '2022-23', '2021-22', '2020-21', '2019-20',
        '2018-19', '2017-18', '2016-17', '2015-16', '2014-15', '2013-14'
    ]
    
    # Price dates
    price_dates = [
        '2025-08-19', '2025-06-30', '2025-03-28', '2024-12-31', '2024-09-30',
        '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29', '2023-06-30'
    ]
    
    return {
        'basic_columns': basic_columns,
        'market_cap_dates': market_cap_dates,
        'ttm_periods': ttm_periods,
        'annual_years': annual_years,
        'price_dates': price_dates
    }

def get_column_positions():
    """
    Returns exact column positions for different data sections
    """
    return {
        'basic_info': list(range(0, 14)),  # Columns 1-13 + separator
        'market_cap': list(range(14, 42)),  # Columns 15-42
        'market_cap_separator': 42,
        'market_cap_ff': list(range(43, 71)),  # Columns 44-71
        'market_cap_ff_separator': 71,
        'ttm_revenue_separator': 72,
        'ttm_revenue': list(range(73, 108)),  # 35 periods
        'ttm_revenue_ff_separator': 108,
        'ttm_revenue_ff': list(range(109, 144)),  # 35 periods
        'ttm_pat_separator': 144,
        'ttm_pat': list(range(145, 180)),  # 35 periods
        'ttm_pat_ff_separator': 180,
        'ttm_pat_ff': list(range(181, 216)),  # 35 periods
        'qtr_revenue_separator': 216,
        'qtr_revenue': list(range(217, 252)),  # 35 periods
        'qtr_revenue_ff_separator': 252,
        'qtr_revenue_ff': list(range(253, 288)),  # 35 periods
        'qtr_pat_separator': 288,
        'qtr_pat': list(range(289, 324)),  # 35 periods
        'qtr_pat_ff_separator': 324,
        'qtr_pat_ff': list(range(325, 360)),  # 35 periods
        'roce_separator': 360,
        'roce': list(range(361, 373)),  # 12 years
        'roe_separator': 373,
        'roe': list(range(374, 386)),  # 12 years
        'retention_separator': 386,
        'retention': list(range(387, 399)),  # 12 years
        'price_separator': 399,
        'share_price': list(range(400, 410)),  # 10 dates
        'pe_separator': 410,
        'pe_ratio': list(range(411, 421)),  # 10 dates
        'identifiers_separator': 421,
        'identifiers': list(range(422, 426)),  # BSE, NSE, ISIN codes
    }