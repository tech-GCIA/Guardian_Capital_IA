"""
Dynamic Admin Export Patch
==========================

This module contains the updated admin export functions that use dynamic periods
instead of static hardcoded periods from header_mapping.py.
"""

def export_stocks_base_sheet_dynamic(self, request, queryset):
    """Export selected stocks data in Base Sheet format with DYNAMIC periods"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from django.utils import timezone
    import tempfile
    from io import BytesIO
    from django.http import HttpResponse

    # Import the dynamic header generator
    from .dynamic_admin_export import DynamicAdminExportGenerator

    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Stocks Base Sheet"

    # Generate dynamic export structure based on actual database data
    dynamic_generator = DynamicAdminExportGenerator()
    export_structure = dynamic_generator.get_complete_dynamic_export_structure()

    header_structure = export_structure['header_structure']
    column_mapping = export_structure['column_mapping']
    periods = export_structure['periods']
    total_columns = export_structure['total_columns']

    print(f"=== DYNAMIC ADMIN EXPORT ===")
    print(f"Total columns: {total_columns} (vs static 426)")
    print(f"Market cap periods: {len(periods['market_cap_dates'])}")
    print(f"TTM periods: {len(periods['ttm_periods'])}")
    print(f"Quarterly periods: {len(periods['quarterly_periods'])}")
    print(f"Annual years: {len(periods['annual_years'])}")
    print(f"Price dates: {len(periods['price_dates'])}")

    # Add all 8 header rows with DYNAMIC structure
    ws.append(header_structure['row_1'])  # Row 1: Empty
    ws.append(header_structure['row_2'])  # Row 2: Column numbers
    ws.append(header_structure['row_3'])  # Row 3: Field descriptions
    ws.append(header_structure['row_4'])  # Row 4: Shareholding info
    ws.append(header_structure['row_5'])  # Row 5: Formulas
    ws.append(header_structure['row_6'])  # Row 6: Category descriptions
    ws.append(header_structure['row_7'])  # Row 7: Sub-category details
    ws.append(header_structure['row_8'])  # Row 8: DYNAMIC column headers with actual periods

    # Add stock data using DYNAMIC COLUMN POSITIONING
    for i, stock in enumerate(queryset, 1):
        # Initialize row with DYNAMIC total columns (not fixed 426)
        row_data = [''] * total_columns

        # Basic info columns (0-12) - same as before
        row_data[0] = i  # S.No
        row_data[1] = stock.company_name
        row_data[2] = stock.accord_code
        row_data[3] = stock.sector
        row_data[4] = stock.cap
        row_data[5] = stock.free_float
        row_data[6] = stock.revenue_6yr_cagr
        row_data[7] = stock.revenue_ttm
        row_data[8] = stock.pat_6yr_cagr
        row_data[9] = stock.pat_ttm
        row_data[10] = stock.current_value
        row_data[11] = stock.two_yr_avg
        row_data[12] = stock.reval_deval

        # Market Cap data - DYNAMIC positions
        market_cap_data = {}
        for mc in stock.market_cap_data.all():
            market_cap_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap

        for idx, date_str in enumerate(periods['market_cap_dates']):
            col_position = column_mapping['market_cap_start'] + idx
            if col_position <= column_mapping['market_cap_end']:
                row_data[col_position] = market_cap_data.get(date_str)

        # Market Cap Free Float data - DYNAMIC positions
        market_cap_ff_data = {}
        for mc in stock.market_cap_data.all():
            market_cap_ff_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap_free_float

        for idx, date_str in enumerate(periods['market_cap_dates']):
            col_position = column_mapping['market_cap_ff_start'] + idx
            if col_position <= column_mapping['market_cap_ff_end']:
                row_data[col_position] = market_cap_ff_data.get(date_str)

        # TTM Revenue data - DYNAMIC positions
        ttm_data = {}
        for ttm in stock.ttm_data.all():
            ttm_data[ttm.period] = ttm

        for idx, period in enumerate(periods['ttm_periods']):
            # TTM Revenue
            col_position = column_mapping['ttm_revenue_start'] + idx
            if col_position <= column_mapping['ttm_revenue_end']:
                ttm_obj = ttm_data.get(period)
                row_data[col_position] = ttm_obj.ttm_revenue if ttm_obj else None

            # TTM Revenue Free Float
            col_position = column_mapping['ttm_revenue_ff_start'] + idx
            if col_position <= column_mapping['ttm_revenue_ff_end']:
                ttm_obj = ttm_data.get(period)
                row_data[col_position] = ttm_obj.ttm_revenue_free_float if ttm_obj else None

            # TTM PAT
            col_position = column_mapping['ttm_pat_start'] + idx
            if col_position <= column_mapping['ttm_pat_end']:
                ttm_obj = ttm_data.get(period)
                row_data[col_position] = ttm_obj.ttm_pat if ttm_obj else None

            # TTM PAT Free Float
            col_position = column_mapping['ttm_pat_ff_start'] + idx
            if col_position <= column_mapping['ttm_pat_ff_end']:
                ttm_obj = ttm_data.get(period)
                row_data[col_position] = ttm_obj.ttm_pat_free_float if ttm_obj else None

        # Quarterly data - DYNAMIC positions
        quarterly_data = {}
        for qtr in stock.quarterly_data.all():
            quarterly_data[qtr.period] = qtr

        for idx, period in enumerate(periods['quarterly_periods']):
            # Quarterly Revenue
            col_position = column_mapping['qtr_revenue_start'] + idx
            if col_position <= column_mapping['qtr_revenue_end']:
                qtr_obj = quarterly_data.get(period)
                row_data[col_position] = qtr_obj.quarterly_revenue if qtr_obj else None

            # Quarterly Revenue Free Float
            col_position = column_mapping['qtr_revenue_ff_start'] + idx
            if col_position <= column_mapping['qtr_revenue_ff_end']:
                qtr_obj = quarterly_data.get(period)
                row_data[col_position] = qtr_obj.quarterly_revenue_free_float if qtr_obj else None

            # Quarterly PAT
            col_position = column_mapping['qtr_pat_start'] + idx
            if col_position <= column_mapping['qtr_pat_end']:
                qtr_obj = quarterly_data.get(period)
                row_data[col_position] = qtr_obj.quarterly_pat if qtr_obj else None

            # Quarterly PAT Free Float
            col_position = column_mapping['qtr_pat_ff_start'] + idx
            if col_position <= column_mapping['qtr_pat_ff_end']:
                qtr_obj = quarterly_data.get(period)
                row_data[col_position] = qtr_obj.quarterly_pat_free_float if qtr_obj else None

        # Annual ratios - DYNAMIC positions
        annual_data = {}
        for ratio in stock.annual_ratios.all():
            annual_data[ratio.financial_year] = ratio

        for idx, year in enumerate(periods['annual_years']):
            # ROCE
            col_position = column_mapping['roce_start'] + idx
            if col_position <= column_mapping['roce_end']:
                ratio_obj = annual_data.get(year)
                row_data[col_position] = ratio_obj.roce_percentage if ratio_obj else None

            # ROE
            col_position = column_mapping['roe_start'] + idx
            if col_position <= column_mapping['roe_end']:
                ratio_obj = annual_data.get(year)
                row_data[col_position] = ratio_obj.roe_percentage if ratio_obj else None

            # Retention
            col_position = column_mapping['retention_start'] + idx
            if col_position <= column_mapping['retention_end']:
                ratio_obj = annual_data.get(year)
                row_data[col_position] = ratio_obj.retention_percentage if ratio_obj else None

        # Price data - DYNAMIC positions
        price_data = {}
        pe_data = {}
        for price in stock.price_data.all():
            price_data[price.price_date.strftime('%Y-%m-%d')] = price.share_price
            pe_data[price.price_date.strftime('%Y-%m-%d')] = price.pe_ratio

        for idx, date_str in enumerate(periods['price_dates']):
            # Share Price
            col_position = column_mapping['price_start'] + idx
            if col_position <= column_mapping['price_end']:
                row_data[col_position] = price_data.get(date_str)

            # PE Ratio
            col_position = column_mapping['pe_start'] + idx
            if col_position <= column_mapping['pe_end']:
                row_data[col_position] = pe_data.get(date_str)

        # Identifiers - DYNAMIC positions
        row_data[column_mapping['bse_col']] = stock.bse_code
        row_data[column_mapping['nse_col']] = stock.nse_symbol
        row_data[column_mapping['isin_col']] = stock.isin

        ws.append(row_data)

    # Save to BytesIO buffer
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_content = excel_buffer.getvalue()
    excel_buffer.close()

    # Create response
    response = HttpResponse(
        excel_content,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = f'attachment; filename="stocks_base_sheet_dynamic_{timezone.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'

    self.message_user(request, f"Successfully exported {len(queryset)} stocks with DYNAMIC periods to Excel file.")
    return response

def export_all_stocks_base_sheet_dynamic(self, request, queryset):
    """Export all stocks data in Base Sheet format with DYNAMIC periods"""
    from django.utils import timezone

    # Get all stocks instead of just the queryset
    all_stocks_queryset = Stock.objects.all()

    # Use the same dynamic logic but with all stocks
    return export_stocks_base_sheet_dynamic(self, request, all_stocks_queryset)