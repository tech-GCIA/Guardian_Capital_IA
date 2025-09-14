# header_mapping.py
# Column structure and date mappings for stock export functionality

def get_complete_header_structure():
    """
    Returns the complete 8-row header structure for the Base Sheet Excel format
    """

    # Row 1: Empty row (426 columns)
    row_1 = [''] * 426

    # Row 2: Column numbers (1-426)
    row_2 = [''] + list(range(1, 425)) + ['', '', '']

    # Row 3: Field descriptions - extracted from CSV
    row_3 = [''] * 426
    row_3[21] = 'Market Cap'
    row_3[22] = 'NDP_MCAP'
    row_3[72] = 'Net Sales'
    row_3[73] = 'TTM_TTMNS'
    row_3[144] = 'Profit After Tax'
    row_3[145] = 'TTM_TTMNP'
    row_3[216] = 'Net Sales & Other Operating Income'
    row_3[217] = 'QTR_NET_SALES'
    row_3[288] = 'Profit after tax'
    row_3[289] = 'QTR_PAT'
    row_3[360] = 'ROCE (%)'
    row_3[361] = 'FR_ROCE'
    row_3[373] = 'ROE (%)'
    row_3[374] = 'FR_ROE'
    row_3[386] = 'Dividend Pay Out Ratio(%)'
    row_3[387] = 'FR_DIVIDEND_PAYOUT_PER'
    row_3[399] = 'NDP_ADJCLOSE'
    row_3[410] = 'NDP_ADJCLOSE'
    row_3[411] = 'NDP_ADJCLOSE'
    row_3[422] = 'BSE Code'
    row_3[423] = 'CD_SCRIPCODE'

    # Row 4: Shareholding info and period descriptions
    row_4 = [''] * 426
    row_4[21] = 'Total Public Shareholding'
    row_4[22] = 'SHP_TPTOTALPUBLIC'
    row_4[73] = 'Other Income'
    row_4[74] = 'TTM_TTM_OTHERINC'

    # Add L, L-1, L-2 etc. for periods
    for i in range(12):
        if i == 0:
            row_4[361 + i] = 'L'
            row_4[374 + i] = 'L'
            row_4[387 + i] = 'L'
        else:
            row_4[361 + i] = f'L-{i}'
            row_4[374 + i] = f'L-{i}'
            row_4[387 + i] = f'L-{i}'

    row_4[423] = 'NSE Symbol'
    row_4[424] = 'CD_SYMBOL'

    # Row 5: Formulas (mostly #DIV/0! and calculations)
    row_5 = [''] * 426
    # Add some formula examples
    for i in range(67, 74):
        row_5[i] = '#DIV/0!'
    for i in range(180, 187):
        row_5[i] = '#DIV/0!'
    row_5[399] = '=EQNXTDZ($B5,$LB$3,LA$4,"X")'
    row_5[400] = '=EQNXTDZ($B5,$LB$3,LA$4,"X")'
    row_5[387] = '=100-EQNXTDZ($B5,$LZ$3,LY$4,"X")'
    row_5[425] = 'ISIN No'
    row_5[424] = 'CD_ISIN'

    # Row 6: Category descriptions
    row_6 = [''] * 426
    row_6[5] = 'Stock wise Fundamentals and Valuations'
    row_6[14] = 'Market Cap (in crores)'
    row_6[15] = 'Market Cap (in crores)'
    row_6[43] = 'Market Cap- Free Float  (in crores)'
    row_6[44] = 'Market Cap- Free Float  (in crores)'
    row_6[45] = 'Market Cap- Free Float  (in crores)'
    row_6[46] = 'Market Cap- Free Float  (in crores)'
    row_6[47] = 'Market Cap- Free Float  (in crores)'
    row_6[72] = 'TTM Revenue'
    row_6[73] = 'TTM Revenue'
    row_6[74] = 'TTM Revenue'
    row_6[108] = 'TTM Revenue- Free Float'
    row_6[109] = 'TTM Revenue- Free Float'
    row_6[110] = 'TTM Revenue- Free Float'
    row_6[111] = 'TTM Revenue- Free Float'
    row_6[144] = 'TTM PAT'
    row_6[145] = 'TTM PAT'
    row_6[146] = 'TTM PAT'
    row_6[180] = 'TTM PAT- Free Float'
    row_6[181] = 'TTM PAT- Free Float'
    row_6[182] = 'TTM PAT- Free Float'
    row_6[183] = 'TTM PAT- Free Float'
    row_6[216] = 'Quarterly- Revenue'
    row_6[217] = 'Quarterly- Revenue'
    row_6[218] = 'Quarterly- Revenue'
    row_6[252] = 'Quarterly- Revenue-  Free Float'
    row_6[253] = 'Quarterly- Revenue-  Free Float'
    row_6[254] = 'Quarterly- Revenue-  Free Float'
    row_6[288] = 'Quarterly- PAT'
    row_6[289] = 'Quarterly- PAT'
    row_6[290] = 'Quarterly- PAT'
    row_6[324] = 'Quarterly-PAT-  Free Float'
    row_6[325] = 'Quarterly-PAT-  Free Float'
    row_6[326] = 'Quarterly-PAT-  Free Float'
    row_6[360] = 'ROCE (%)'
    row_6[373] = 'ROE (%)'
    row_6[386] = 'Retention (%)'
    row_6[399] = 'Share Price'

    # Row 7: Sub-category details
    row_7 = [''] * 426
    row_7[5] = 'Revenue (Q1 FY-26)'
    row_7[8] = 'PAT (Q1 FY-26)'
    row_7[10] = 'PR'
    row_7[14] = 'Market Cap (in crores)'
    row_7[43] = 'Market Cap- Free Float  (in crores)'
    row_7[72] = 'TTM Revenue'
    row_7[108] = 'TTM Revenue- Free Float'
    row_7[144] = 'TTM PAT'
    row_7[180] = 'TTM PAT- Free Float'
    row_7[216] = 'Quarterly- Revenue'
    row_7[252] = 'Quarterly- Revenue-  Free Float'
    row_7[288] = 'Quarterly- PAT'
    row_7[324] = 'Quarterly-PAT-  Free Float'
    row_7[360] = 'ROCE (%)'
    row_7[373] = 'ROE (%)'
    row_7[386] = 'Retention (%)'
    row_7[399] = 'Share Price'
    row_7[400] = 'PR'
    row_7[410] = 'PE'

    return {
        'row_1': row_1,
        'row_2': row_2,
        'row_3': row_3,
        'row_4': row_4,
        'row_5': row_5,
        'row_6': row_6,
        'row_7': row_7
    }

def get_data_column_mapping():
    """
    Returns the column mappings for different data types
    """

    # Basic columns (columns 0-12)
    basic_columns = [
        'S. No.', 'Company Name', 'Accord Code', 'Sector', 'Cap',
        'Free Float', '6 Year CAGR', 'TTM', '6 Year CAGR', 'TTM',
        'Current', '2 Yr Avg', 'Reval/deval'
    ]

    # Market cap dates (columns 14-41) - based on CSV structure
    market_cap_dates = [
        '2025-08-19', '2025-06-30', '2025-03-28', '2025-01-31', '2024-12-31',
        '2024-09-30', '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29',
        '2023-06-30', '2023-03-31', '2022-12-30', '2022-09-30', '2022-06-30',
        '2022-03-31', '2021-12-31', '2021-09-30', '2021-06-30', '2021-03-31',
        '2020-12-31', '2020-09-30', '2020-06-30', '2020-03-31', '2019-12-31',
        '2019-09-30', '2019-06-28', '2019-04-01'
    ]

    # TTM periods for financial data (202506, 202503, etc.)
    ttm_periods = [
        '202506', '202503', '202412', '202409', '202406', '202403',
        '202312', '202309', '202306', '202303', '202212', '202209',
        '202206', '202203', '202112', '202109', '202106', '202103',
        '202012', '202009', '202006', '202003', '201912', '201909',
        '201906', '201903', '201812', '201809', '201806', '201803',
        '201712', '201709', '201706', '201703', '201612'
    ]

    # Annual financial years
    annual_years = [
        '2024-25', '2023-24', '2022-23', '2021-22', '2020-21', '2019-20',
        '2018-19', '2017-18', '2016-17', '2015-16', '2014-15', '2013-14'
    ]

    # Price dates (subset of market cap dates)
    price_dates = [
        '2025-08-19', '2025-06-30', '2025-03-28', '2024-12-31', '2024-09-30',
        '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29', '2023-06-30',
        '2023-03-31'
    ]

    return {
        'basic_columns': basic_columns,
        'market_cap_dates': market_cap_dates,
        'ttm_periods': ttm_periods,
        'annual_years': annual_years,
        'price_dates': price_dates
    }

def get_fund_holding_columns():
    """
    Returns additional columns for fund holding data that will be prepended to stock data
    """
    return [
        'Fund Name', 'Fund ISIN', 'Holding Date', 'Holding Percentage (%)',
        'Market Value', 'Number of Shares'
    ]

def get_fund_analysis_header_structure():
    """
    Returns the complete 8-row header structure for Fund Analysis Excel format
    Based on sample file: "Old Bridge Focused Equity Fund.xlsm"
    """

    # Row 1: Fund name only in first cell, rest empty
    row_1 = [''] * 405
    # Fund name will be populated dynamically

    # Row 2: Column numbers (1, 2, 3, 4, 14, ..., 389, 5, 6, 7, 8, 9)
    row_2 = [''] * 405
    row_2[0] = '1'
    row_2[1] = '2'
    row_2[2] = '3'
    row_2[3] = '4'
    row_2[4] = '14'
    # Skip 5-8 for now, will be filled in sequence
    row_2[389] = '389'
    row_2[9] = '5'
    row_2[10] = '6'
    row_2[11] = '7'
    row_2[12] = '8'
    row_2[13] = '9'

    # Row 3: Portfolio date (will be populated dynamically)
    row_3 = [''] * 405
    # "Portfolio as on: 30th April 2025" will be populated dynamically

    # Row 4-6: Mostly empty
    row_4 = [''] * 405
    row_5 = [''] * 405
    row_6 = [''] * 405

    # Row 6 has some labels
    row_6[5] = 'To be filled'
    row_6[8] = 'To be filled'
    row_6[11] = 'Stock wise Fundamentals and Valuations'

    # Row 7: Sub-category details
    row_7 = [''] * 405
    row_7[11] = 'Revenue (Q3 FY-25)'
    row_7[13] = 'PAT (Q3 FY-25)'

    return {
        'row_1': row_1,
        'row_2': row_2,
        'row_3': row_3,
        'row_4': row_4,
        'row_5': row_5,
        'row_6': row_6,
        'row_7': row_7
    }

def get_fund_analysis_column_mapping():
    """
    Returns the column mappings for Fund Analysis format
    Based on sample file structure with integrated fund holding data
    """

    # Fund-integrated basic columns (0-17)
    fund_basic_columns = [
        'Company Name',      # 0
        'Accord Code',       # 1
        'Sector',           # 2
        'Cap',              # 3
        'Market Cap',       # 4 - calculated from stock data
        'Weights',          # 5 - holding_percentage / 100
        'Factor',           # 6 - calculated field
        'Value',            # 7 - market_value from holding
        'No.of shares',     # 8 - number_of_shares from holding
        'Share Price',      # 9 - from stock price data
        'Free Float',       # 10
        '6 Year CAGR',      # 11 - Revenue
        'TTM',              # 12 - Revenue
        '6 Year CAGR',      # 13 - PAT
        'TTM',              # 14 - PAT
        'Current',          # 15
        '2 Yr Avg',         # 16
        'Reval/deval'       # 17
    ]

    # Market cap dates - start from column 19 (shifted from 14)
    market_cap_dates = [
        '2025-03-31', '2024-12-31', '2024-09-30', '2024-06-30', '2024-03-31',
        '2023-12-31', '2023-09-30', '2023-06-30', '2023-03-31', '2022-12-31',
        '2022-09-30', '2022-06-30', '2022-03-31', '2021-12-31', '2021-09-30',
        '2021-06-30', '2021-03-31', '2020-12-31', '2020-09-30', '2020-06-30',
        '2020-03-31', '2019-12-31', '2019-09-30', '2019-06-30', '2019-03-31'
    ]

    # TTM periods for financial data
    ttm_periods = [
        '202503', '202412', '202409', '202406', '202403', '202312', '202309',
        '202306', '202303', '202212', '202209', '202206', '202203', '202112',
        '202109', '202106', '202103', '202012', '202009', '202006', '202003',
        '201912', '201909', '201906', '201903', '201812', '201809', '201806',
        '201803', '201712', '201709', '201706', '201703', '201612', '201609'
    ]

    # Annual financial years
    annual_years = [
        '2024-25', '2023-24', '2022-23', '2021-22', '2020-21', '2019-20',
        '2018-19', '2017-18', '2016-17', '2015-16', '2014-15', '2013-14'
    ]

    # Price dates (subset of market cap dates)
    price_dates = [
        '2025-03-31', '2024-12-31', '2024-09-30', '2024-06-30', '2024-03-31',
        '2023-12-31', '2023-09-30', '2023-06-30', '2023-03-31', '2022-12-31'
    ]

    return {
        'fund_basic_columns': fund_basic_columns,
        'market_cap_dates': market_cap_dates,
        'ttm_periods': ttm_periods,
        'annual_years': annual_years,
        'price_dates': price_dates,

        # Column position mappings for fund analysis format
        'market_cap_start': 19,        # Start column for market cap data
        'market_cap_ff_start': 50,     # Market cap free float
        'ttm_revenue_start': 80,       # TTM Revenue
        'ttm_revenue_ff_start': 120,   # TTM Revenue Free Float
        'ttm_pat_start': 160,          # TTM PAT
        'ttm_pat_ff_start': 200,       # TTM PAT Free Float
        'qtr_revenue_start': 240,      # Quarterly Revenue
        'qtr_revenue_ff_start': 280,   # Quarterly Revenue Free Float
        'qtr_pat_start': 320,          # Quarterly PAT
        'qtr_pat_ff_start': 360,       # Quarterly PAT Free Float
        'roce_start': 370,             # ROCE ratios
        'roe_start': 380,              # ROE ratios
        'retention_start': 390,        # Retention ratios
        'share_price_start': 395,      # Share prices
        'pe_ratio_start': 400,         # PE ratios
        'identifiers_start': 402       # BSE, NSE, ISIN codes
    }

def adapt_headers_for_fund_format():
    """
    Adapts the working original header structure for 405-column fund format
    Keeps all professional field descriptions and category labels
    """
    # Get the working original structure (426 columns)
    original_structure = get_complete_header_structure()

    # Adapt each row for 405 columns with integrated fund data
    adapted_structure = {}

    for row_key, row_data in original_structure.items():
        # Trim to 405 columns and adjust for integrated fund columns
        adapted_row = [''] * 405

        # Copy basic columns 0-3 (Company Name, Accord Code, Sector, Cap)
        for i in range(4):
            if i < len(row_data):
                adapted_row[i] = row_data[i]

        # Columns 4-8 are now integrated fund columns (Market Cap, Weights, Factor, Value, No.of shares)
        # Column 9 is Share Price, Column 10 is Free Float
        # Copy remaining columns with position shift

        # Original stock columns 4-9 (Free Float, 6 Year CAGR, etc.) now go to positions 10-15
        for i in range(4, 10):
            if i < len(row_data):
                adapted_row[i + 6] = row_data[i]  # Shift by 6 positions

        # Original stock columns 10-12 (Current, 2 Yr Avg, Reval/deval) now go to positions 15-17
        for i in range(10, 13):
            if i < len(row_data):
                adapted_row[i + 5] = row_data[i]  # Shift by 5 positions

        # Market cap and other data sections shift from original positions
        # Original market cap started at 14, now starts at 19
        shift_offset = 5  # 19 - 14 = 5

        # Copy market cap and subsequent data sections with shift
        for i in range(14, min(len(row_data), 400)):  # Original columns 14-399
            new_pos = i + shift_offset
            if new_pos < 405:  # Ensure we don't exceed 405 columns
                adapted_row[new_pos] = row_data[i]

        adapted_structure[row_key] = adapted_row

    return adapted_structure

def get_fund_integrated_column_mapping():
    """
    Returns column mappings for fund format with integrated fund holdings
    Based on adapted original structure (405 columns)
    """
    # Get original mapping
    original_mapping = get_data_column_mapping()

    # Fund-integrated basic columns (0-17) with fund data at positions 4-8
    fund_integrated_columns = [
        'Company Name',      # 0
        'Accord Code',       # 1
        'Sector',           # 2
        'Cap',              # 3
        'Market Cap',       # 4 - calculated from stock data (FUND INTEGRATED)
        'Weights',          # 5 - holding_percentage / 100 (FUND INTEGRATED)
        'Factor',           # 6 - calculated field (FUND INTEGRATED)
        'Value',            # 7 - market_value from holding (FUND INTEGRATED)
        'No.of shares',     # 8 - number_of_shares from holding (FUND INTEGRATED)
        'Share Price',      # 9 - from stock price data
        'Free Float',       # 10 - original position 5, shifted
        '6 Year CAGR',      # 11 - Revenue, original position 6, shifted
        'TTM',              # 12 - Revenue, original position 7, shifted
        '6 Year CAGR',      # 13 - PAT, original position 8, shifted
        'TTM',              # 14 - PAT, original position 9, shifted
        'Current',          # 15 - original position 10, shifted
        '2 Yr Avg',         # 16 - original position 11, shifted
        'Reval/deval'       # 17 - original position 12, shifted
    ]

    # Adjust all start positions by +5 due to fund column integration
    shift_offset = 5

    return {
        'fund_integrated_columns': fund_integrated_columns,
        'market_cap_dates': original_mapping['market_cap_dates'],
        'ttm_periods': original_mapping['ttm_periods'],
        'annual_years': original_mapping['annual_years'],
        'price_dates': original_mapping['price_dates'],

        # Adjusted column positions for 405-column format
        'market_cap_start': 14 + shift_offset,        # 19 (was 14)
        'market_cap_ff_start': 43 + shift_offset,     # 48 (was 43)
        'ttm_revenue_start': 72 + shift_offset,       # 77 (was 72)
        'ttm_revenue_ff_start': 108 + shift_offset,   # 113 (was 108)
        'ttm_pat_start': 144 + shift_offset,          # 149 (was 144)
        'ttm_pat_ff_start': 180 + shift_offset,       # 185 (was 180)
        'qtr_revenue_start': 216 + shift_offset,      # 221 (was 216)
        'qtr_revenue_ff_start': 252 + shift_offset,   # 257 (was 252)
        'qtr_pat_start': 288 + shift_offset,          # 293 (was 288)
        'qtr_pat_ff_start': 324 + shift_offset,       # 329 (was 324)
        'roce_start': 360 + shift_offset,             # 365 (was 360)
        'roe_start': 373 + shift_offset,              # 378 (was 373)
        'retention_start': 386 + shift_offset,        # 391 (was 386)
        'share_price_start': 399 + shift_offset,      # 404 (was 399) - CAREFUL: might exceed 405!
        'pe_ratio_start': 410 + shift_offset,         # 415 (was 410) - EXCEEDS 405!
        'identifiers_start': 400                      # Adjust to fit in 405 columns
    }

def get_full_fund_integrated_headers():
    """
    Creates 431-column structure: ALL 426 stock columns + 5 fund columns
    Preserves every single stock data column without loss
    Inserts fund columns at positions 4-8, shifts stock data by +5
    """
    # Get the complete original 426-column structure
    original_structure = get_complete_header_structure()

    # Create 431-column structure (426 stock + 5 fund = 431)
    integrated_structure = {}

    for row_key, row_data in original_structure.items():
        # Create 431-column row
        integrated_row = [''] * 431

        # Copy first 4 columns (0-3): Company Name, Accord Code, Sector, Cap
        for i in range(4):
            if i < len(row_data):
                integrated_row[i] = row_data[i]

        # Positions 4-8 are reserved for fund columns (will be set in specific rows)
        # Market Cap, Weights, Factor, Value, No.of shares

        # Copy ALL remaining stock columns (4-425) to positions 9-430
        for i in range(4, len(row_data)):
            new_position = i + 5  # Shift by +5 to account for fund columns
            if new_position < 431:
                integrated_row[new_position] = row_data[i]

        integrated_structure[row_key] = integrated_row

    # Set fund column headers in appropriate rows
    # Row 7 (final header row) gets the fund column names
    if 'row_7' in integrated_structure:
        integrated_structure['row_7'][4] = 'Market Cap'
        integrated_structure['row_7'][5] = 'Weights'
        integrated_structure['row_7'][6] = 'Factor'
        integrated_structure['row_7'][7] = 'Value'
        integrated_structure['row_7'][8] = 'No.of shares'

    return integrated_structure

def get_complete_fund_column_mapping():
    """
    Complete column mapping for 431-column fund format
    Preserves ALL original stock data positions with +5 shift
    """
    # Get original mapping for reference
    original_mapping = get_data_column_mapping()

    # Fund-integrated basic columns (0-8 + shifted stock columns)
    fund_integrated_columns = [
        'Company Name',      # 0
        'Accord Code',       # 1
        'Sector',           # 2
        'Cap',              # 3
        'Market Cap',       # 4 - calculated from stock data (FUND INTEGRATED)
        'Weights',          # 5 - holding_percentage / 100 (FUND INTEGRATED)
        'Factor',           # 6 - calculated field (FUND INTEGRATED)
        'Value',            # 7 - market_value from holding (FUND INTEGRATED)
        'No.of shares',     # 8 - number_of_shares from holding (FUND INTEGRATED)
        'Free Float',       # 9 - original position 4, shifted by +5
        '6 Year CAGR',      # 10 - Revenue, original position 5, shifted by +5
        'TTM',              # 11 - Revenue, original position 6, shifted by +5
        '6 Year CAGR',      # 12 - PAT, original position 7, shifted by +5
        'TTM',              # 13 - PAT, original position 8, shifted by +5
        'Current',          # 14 - original position 9, shifted by +5
        '2 Yr Avg',         # 15 - original position 10, shifted by +5
        'Reval/deval'       # 16 - original position 11, shifted by +5
    ]

    # ALL original positions shifted by +5 (no data loss)
    shift_offset = 5

    return {
        'fund_integrated_columns': fund_integrated_columns,
        'market_cap_dates': original_mapping['market_cap_dates'],
        'ttm_periods': original_mapping['ttm_periods'],
        'annual_years': original_mapping['annual_years'],
        'price_dates': original_mapping['price_dates'],

        # ALL original column positions preserved with +5 shift
        'market_cap_start': 14 + shift_offset,        # 19 (preserves all market cap data)
        'market_cap_ff_start': 43 + shift_offset,     # 48 (preserves all free float data)
        'ttm_revenue_start': 72 + shift_offset,       # 77 (preserves all TTM revenue)
        'ttm_revenue_ff_start': 108 + shift_offset,   # 113 (preserves all TTM revenue FF)
        'ttm_pat_start': 144 + shift_offset,          # 149 (preserves all TTM PAT)
        'ttm_pat_ff_start': 180 + shift_offset,       # 185 (preserves all TTM PAT FF)
        'qtr_revenue_start': 216 + shift_offset,      # 221 (preserves all quarterly revenue)
        'qtr_revenue_ff_start': 252 + shift_offset,   # 257 (preserves all quarterly revenue FF)
        'qtr_pat_start': 288 + shift_offset,          # 293 (preserves all quarterly PAT)
        'qtr_pat_ff_start': 324 + shift_offset,       # 329 (preserves all quarterly PAT FF)
        'roce_start': 360 + shift_offset,             # 365 (preserves all ROCE data)
        'roe_start': 373 + shift_offset,              # 378 (preserves all ROE data)
        'retention_start': 386 + shift_offset,        # 391 (preserves all retention ratios)
        'share_price_start': 399 + shift_offset,      # 404 (preserves all share prices)
        'pe_ratio_start': 410 + shift_offset,         # 415 (preserves all PE ratios)
        'identifiers_start': 422 + shift_offset       # 427 (preserves BSE, NSE, ISIN codes)
    }