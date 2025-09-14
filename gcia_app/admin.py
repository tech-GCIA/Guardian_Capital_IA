# Register your models here.
from django.contrib import admin, messages
from django.http import HttpResponse
from gcia_app.models import *
from import_export.admin import ImportExportModelAdmin
import string
import random
from django.utils.crypto import get_random_string
from django.contrib.auth.hashers import make_password
import os
import datetime
import tempfile
from io import BytesIO

class ReadOnlyModelAdmin(admin.ModelAdmin):
    """Make all fields read-only for non-staff users."""

    def has_add_permission(self, request):
        if request.user.is_superuser:
            return super().has_add_permission(request)
        return False

    def has_change_permission(self, request, obj=None):
        if request.user.is_superuser:
            return super().has_change_permission(request, obj)
        return False

    def has_delete_permission(self, request, obj=None):
        if request.user.is_superuser:
            return super().has_delete_permission(request, obj)
        return False
    
    def get_import_permission(self, request):
        if request.user.is_superuser:
            return super().get_import_permission(request)
        return False
    
class CustomerAdmin(admin.ModelAdmin):
    list_display = ('customer_id', 'first_name', 'last_name', 'email', 'phone_number', 'pan_number') # List the fields to display in the list view
    search_fields = ('first_name', 'last_name', 'email') # Add search capability for first name, last name, and email
    ordering = ('customer_id',) # Allow sorting by customer_id
    list_per_page = 100 # Enable pagination
    # readonly_fields = ('password',)

    # actions = ['generate_random_password']
    
    # Custom queryset for displaying records
    def get_queryset(self, request):
        queryset = super().get_queryset(request)
        return queryset
    
    # def generate_random_password(self, request, queryset):
    #     """Admin action to generate and save a strong random password."""
    #     characters = string.ascii_letters + string.digits + string.punctuation
    #     for customer in queryset:
    #         raw_password = get_random_string(12, characters)
    #         customer.password = make_password(raw_password)
    #         customer.save()
    #         self.message_user(
    #             request,
    #             f"Password for {customer.email} updated to: {raw_password}",
    #             level=messages.INFO
    #         )

    # generate_random_password.short_description = "Generate and set a strong random password (12 characters)"

        
class AMCFundSchemeAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['name']
    list_display = ('amcfundscheme_id', 'name', 'isin_number', 'scheme_benchmark', 'is_active', 'latest_nav', 'latest_nav_as_on_date')
    readonly_fields = ['amcfundscheme_id', 'created', 'modified','assets_under_management']

    list_filter = ['is_active', 'is_direct_fund', 'is_scheme_benchmark']

    def get_queryset(self, request):
        if 'change' not in request.get_full_path():
            return super(AMCFundSchemeAdmin, self).get_queryset(request)
        else:
            return super(AMCFundSchemeAdmin, self).get_queryset(request)
        
class AMCFundSchemeNavLogAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['amcfundscheme__name']
    list_display = ('amcfundschemenavlog_id', 'amcfundscheme__name', 'as_on_date', 'nav')

    raw_id_fields = ['amcfundscheme']

    def get_queryset(self, request):
        if 'change' not in request.get_full_path():
            return super(AMCFundSchemeNavLogAdmin, self).get_queryset(request).select_related('amcfundscheme')
        else:
            return super(AMCFundSchemeNavLogAdmin, self).get_queryset(request)

class StockAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['company_name', 'accord_code', 'sector', 'bse_code', 'nse_symbol', 'isin']
    list_display = ('stock_id', 'company_name', 'accord_code', 'sector', 'cap', 'bse_code', 'nse_symbol', 'created')
    list_filter = ['sector', 'cap', 'created', 'modified']
    readonly_fields = ['stock_id', 'created', 'modified']
    
    actions = ['export_stocks_base_sheet', 'export_all_stocks_base_sheet']
    
    def export_stocks_base_sheet(self, request, queryset):
        """Export selected stocks data in Base Sheet format"""
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
        from django.utils import timezone
        import tempfile
        
        # Import the corrected header structure
        import sys
        import os
        sys.path.append(os.path.dirname(os.path.abspath(__file__)))
        from header_mapping import get_complete_header_structure, get_data_column_mapping
        
        # Create workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Stocks Base Sheet"
        
        # Get the complete header structure
        header_structure = get_complete_header_structure()
        column_mapping = get_data_column_mapping()
        
        # Add all 8 header rows with correct structure
        ws.append(header_structure['row_1'])  # Row 1: Empty
        ws.append(header_structure['row_2'])  # Row 2: Column numbers
        
        ws.append(header_structure['row_3'])  # Row 3: Field descriptions
        ws.append(header_structure['row_4'])  # Row 4: Shareholding info
        ws.append(header_structure['row_5'])  # Row 5: Formulas
        ws.append(header_structure['row_6'])  # Row 6: Category descriptions
        ws.append(header_structure['row_7'])  # Row 7: Sub-category details
        
        # Row 8: Build actual column headers with proper structure
        row8_headers = [''] * 426
        
        # Basic columns (1-13)
        basic_cols = column_mapping['basic_columns']
        for i, header in enumerate(basic_cols):
            row8_headers[i] = header
        
        # Market Cap dates (15-42)
        market_cap_dates = column_mapping['market_cap_dates']
        for i, date in enumerate(market_cap_dates):
            row8_headers[14 + i] = date  # Start at column 15 (index 14)
            
        # Market Cap Free Float dates (44-71) - same dates, different section
        for i, date in enumerate(market_cap_dates):
            row8_headers[43 + i] = date  # Start at column 44 (index 43)
            
        # TTM periods for various sections
        ttm_periods = column_mapping['ttm_periods']
        
        # TTM Revenue (73-108)
        for i, period in enumerate(ttm_periods):
            row8_headers[72 + i] = period
            
        # TTM Revenue Free Float (109-144)
        for i, period in enumerate(ttm_periods):
            row8_headers[108 + i] = period
            
        # TTM PAT (145-180)
        for i, period in enumerate(ttm_periods):
            row8_headers[144 + i] = period
            
        # TTM PAT Free Float (181-216)
        for i, period in enumerate(ttm_periods):
            row8_headers[180 + i] = period
            
        # Quarterly Revenue (217-252)
        for i, period in enumerate(ttm_periods):
            row8_headers[216 + i] = period
            
        # Quarterly Revenue Free Float (253-288)
        for i, period in enumerate(ttm_periods):
            row8_headers[252 + i] = period
            
        # Quarterly PAT (289-324)
        for i, period in enumerate(ttm_periods):
            row8_headers[288 + i] = period
            
        # Quarterly PAT Free Float (325-360)
        for i, period in enumerate(ttm_periods):
            row8_headers[324 + i] = period
        
        # Annual ratios
        annual_years = column_mapping['annual_years']
        
        # ROCE (361-372)
        for i, year in enumerate(annual_years):
            row8_headers[360 + i] = year
            
        # ROE (374-385)
        for i, year in enumerate(annual_years):
            row8_headers[373 + i] = year
            
        # Retention (387-398)
        for i, year in enumerate(annual_years):
            row8_headers[386 + i] = year
        
        # Price dates
        price_dates = column_mapping['price_dates']
        
        # Share Price (400-409)
        for i, date in enumerate(price_dates):
            row8_headers[399 + i] = date
            
        # PE Ratio (411-420)
        for i, date in enumerate(price_dates):
            row8_headers[410 + i] = date
        
        # Identifiers (422-425)
        row8_headers[421] = ''  # Separator
        row8_headers[422] = 'BSE Code'
        row8_headers[423] = 'NSE Code'
        row8_headers[424] = 'ISIN'
        
        ws.append(row8_headers)
        
        # Add stock data using FIXED COLUMN POSITIONING to match corrected import logic
        for i, stock in enumerate(queryset, 1):
            # Initialize row with 426 empty columns
            row_data = [''] * 426
            
            # Basic info columns (0-12) - matches corrected import logic
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
            
            # Market Cap data - columns 14-41 (matches corrected import logic)
            market_cap_data = {}
            for mc in stock.market_cap_data.all():
                market_cap_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap
            
            market_cap_dates = column_mapping['market_cap_dates']
            for idx, date_str in enumerate(market_cap_dates):
                col_position = 14 + idx  # Start at column 14
                if col_position <= 41:  # Within the range
                    try:
                        lookup_date = date_str
                        row_data[col_position] = market_cap_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # Market Cap Free Float data - columns 43-70 (matches corrected import logic)
            market_cap_ff_data = {}
            for mc in stock.market_cap_data.all():
                market_cap_ff_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap_free_float
            
            for idx, date_str in enumerate(market_cap_dates):
                col_position = 43 + idx  # Start at column 43 (corrected)
                if col_position <= 70:  # Within the range
                    try:
                        lookup_date = date_str
                        row_data[col_position] = market_cap_ff_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # TTM Revenue data - columns 72-107 (matches corrected import logic)
            ttm_data = {}
            for ttm in stock.ttm_data.all():
                ttm_data[ttm.period] = ttm
            
            ttm_periods = column_mapping['ttm_periods']
            for idx, period in enumerate(ttm_periods):
                col_position = 72 + idx  # Start at column 72 (corrected)
                if col_position <= 107:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_revenue if ttm_obj else None
            
            # TTM Revenue Free Float data - columns 108-143 (matches corrected import logic)  
            for idx, period in enumerate(ttm_periods):
                col_position = 108 + idx  # Start at column 108 (corrected)
                if col_position <= 143:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_revenue_free_float if ttm_obj else None
            
            # TTM PAT data - columns 144-179 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 144 + idx  # Start at column 144
                if col_position <= 179:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_pat if ttm_obj else None
            
            # TTM PAT Free Float data - columns 180-215 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 180 + idx  # Start at column 180 (corrected)
                if col_position <= 215:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_pat_free_float if ttm_obj else None
            
            # Quarterly Revenue data - columns 216-251 (matches corrected import logic)
            quarterly_data = {}
            for qtr in stock.quarterly_data.all():
                quarterly_data[qtr.period] = qtr
            
            for idx, period in enumerate(ttm_periods):
                col_position = 216 + idx  # Start at column 216 (corrected)
                if col_position <= 251:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_revenue if qtr_obj else None
            
            # Quarterly Revenue Free Float data - columns 252-287 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 252 + idx  # Start at column 252 (corrected)
                if col_position <= 287:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_revenue_free_float if qtr_obj else None
            
            # Quarterly PAT data - columns 288-323 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 288 + idx  # Start at column 288 (corrected)
                if col_position <= 323:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_pat if qtr_obj else None
            
            # Quarterly PAT Free Float data - columns 324-359 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 324 + idx  # Start at column 324 (corrected)
                if col_position <= 359:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_pat_free_float if qtr_obj else None
            
            # ROCE data - columns 360-372 (matches corrected import logic)
            annual_data = {}
            for ratio in stock.annual_ratios.all():
                annual_data[ratio.financial_year] = ratio
            
            annual_years = column_mapping['annual_years']
            for idx, year in enumerate(annual_years):
                col_position = 360 + idx  # Start at column 360 (corrected)
                if col_position <= 372:  # Within the range
                    ratio_obj = annual_data.get(year)
                    row_data[col_position] = ratio_obj.roce_percentage if ratio_obj else None
            
            # ROE data - columns 373-385 (matches corrected import logic)
            for idx, year in enumerate(annual_years):
                col_position = 373 + idx  # Start at column 373 (corrected)
                if col_position <= 385:  # Within the range
                    ratio_obj = annual_data.get(year)
                    row_data[col_position] = ratio_obj.roe_percentage if ratio_obj else None
            
            # Retention data - columns 386-398 (matches corrected import logic)
            for idx, year in enumerate(annual_years):
                col_position = 386 + idx  # Start at column 386 (corrected)
                if col_position <= 398:  # Within the range
                    ratio_obj = annual_data.get(year)
                    row_data[col_position] = ratio_obj.retention_percentage if ratio_obj else None
            
            # Share Price data - columns 399-409 (matches corrected import logic)
            price_data = {}
            pe_data = {}
            for price in stock.price_data.all():
                price_data[price.price_date.strftime('%Y-%m-%d')] = price.share_price
                pe_data[price.price_date.strftime('%Y-%m-%d')] = price.pe_ratio
            
            price_dates = column_mapping['price_dates']
            for idx, date_str in enumerate(price_dates):
                col_position = 399 + idx  # Start at column 399 (corrected)
                if col_position <= 409:  # Within the range
                    try:
                        lookup_date = date_str  # Date already in correct format
                        row_data[col_position] = price_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # PE Ratio data - columns 410-420 (matches corrected import logic)
            for idx, date_str in enumerate(price_dates):
                col_position = 410 + idx  # Start at column 410 (corrected)
                if col_position <= 420:  # Within the range
                    try:
                        lookup_date = date_str  # Date already in correct format
                        row_data[col_position] = pe_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # Identifiers - fixed positions (matches corrected import logic)
            row_data[422] = stock.bse_code
            row_data[423] = stock.nse_symbol
            row_data[424] = stock.isin
            
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
        response['Content-Disposition'] = f'attachment; filename="stocks_base_sheet_selected_{timezone.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
        
        self.message_user(request, f"Successfully exported {len(queryset)} stocks to Excel file.")
        return response
    
    export_stocks_base_sheet.short_description = "Export selected stocks as Base Sheet Excel"
    
    def export_all_stocks_base_sheet(self, request, queryset):
        """Export all stocks data in Base Sheet format"""
        from django.utils import timezone
        
        # Get all stocks instead of just the queryset
        all_stocks_queryset = Stock.objects.all()
        
        # Use the same logic but with all stocks
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
        
        # Import the corrected header structure
        import sys
        import os
        sys.path.append(os.path.dirname(os.path.abspath(__file__)))
        from header_mapping import get_complete_header_structure, get_data_column_mapping
        
        # Create workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Stocks Base Sheet"
        
        # Get the complete header structure
        header_structure = get_complete_header_structure()
        column_mapping = get_data_column_mapping()
        
        # Add all 8 header rows with correct structure
        ws.append(header_structure['row_1'])  # Row 1: Empty
        ws.append(header_structure['row_2'])  # Row 2: Column numbers
        
        ws.append(header_structure['row_3'])  # Row 3: Field descriptions
        ws.append(header_structure['row_4'])  # Row 4: Shareholding info
        ws.append(header_structure['row_5'])  # Row 5: Formulas
        ws.append(header_structure['row_6'])  # Row 6: Category descriptions
        ws.append(header_structure['row_7'])  # Row 7: Sub-category details
        
        # Row 8: Build actual column headers with proper structure
        row8_headers = [''] * 426
        
        # Basic columns (1-13)
        basic_cols = column_mapping['basic_columns']
        for i, header in enumerate(basic_cols):
            row8_headers[i] = header
        
        # Market Cap dates (15-42)
        market_cap_dates = column_mapping['market_cap_dates']
        for i, date in enumerate(market_cap_dates):
            row8_headers[14 + i] = date  # Start at column 15 (index 14)
            
        # Market Cap Free Float dates (44-71) - same dates, different section
        for i, date in enumerate(market_cap_dates):
            row8_headers[43 + i] = date  # Start at column 44 (index 43)
            
        # TTM periods for various sections
        ttm_periods = column_mapping['ttm_periods']
        
        # TTM Revenue (73-108)
        for i, period in enumerate(ttm_periods):
            row8_headers[72 + i] = period
            
        # TTM Revenue Free Float (109-144)
        for i, period in enumerate(ttm_periods):
            row8_headers[108 + i] = period
            
        # TTM PAT (145-180)
        for i, period in enumerate(ttm_periods):
            row8_headers[144 + i] = period
            
        # TTM PAT Free Float (181-216)
        for i, period in enumerate(ttm_periods):
            row8_headers[180 + i] = period
            
        # Quarterly Revenue (217-252)
        for i, period in enumerate(ttm_periods):
            row8_headers[216 + i] = period
            
        # Quarterly Revenue Free Float (253-288)
        for i, period in enumerate(ttm_periods):
            row8_headers[252 + i] = period
            
        # Quarterly PAT (289-324)
        for i, period in enumerate(ttm_periods):
            row8_headers[288 + i] = period
            
        # Quarterly PAT Free Float (325-360)
        for i, period in enumerate(ttm_periods):
            row8_headers[324 + i] = period
        
        # Annual ratios
        annual_years = column_mapping['annual_years']
        
        # ROCE (361-372)
        for i, year in enumerate(annual_years):
            row8_headers[360 + i] = year
            
        # ROE (374-385)
        for i, year in enumerate(annual_years):
            row8_headers[373 + i] = year
            
        # Retention (387-398)
        for i, year in enumerate(annual_years):
            row8_headers[386 + i] = year
        
        # Price dates
        price_dates = column_mapping['price_dates']
        
        # Share Price (400-409)
        for i, date in enumerate(price_dates):
            row8_headers[399 + i] = date
            
        # PE Ratio (411-420)
        for i, date in enumerate(price_dates):
            row8_headers[410 + i] = date
        
        # Identifiers (422-425)
        row8_headers[421] = ''  # Separator
        row8_headers[422] = 'BSE Code'
        row8_headers[423] = 'NSE Code'
        row8_headers[424] = 'ISIN'
        
        ws.append(row8_headers)
        
        # Add data rows for all stocks using FIXED COLUMN POSITIONING
        for idx, stock in enumerate(all_stocks_queryset, 1):
            # Initialize row with 426 empty columns
            row_data = [''] * 426
            
            # Basic info columns (0-12) - matches corrected import logic
            row_data[0] = idx  # S.No
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
            
            # Market Cap data - columns 14-41 (matches corrected import logic)
            market_cap_data = {}
            for mc in stock.market_cap_data.all():
                market_cap_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap
            
            market_cap_dates = column_mapping['market_cap_dates']
            for idx, date_str in enumerate(market_cap_dates):
                col_position = 14 + idx  # Start at column 14
                if col_position <= 41:  # Within the range
                    try:
                        lookup_date = date_str
                        row_data[col_position] = market_cap_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # Market Cap Free Float data - columns 43-70 (matches corrected import logic)
            market_cap_ff_data = {}
            for mc in stock.market_cap_data.all():
                market_cap_ff_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap_free_float
            
            for idx, date_str in enumerate(market_cap_dates):
                col_position = 43 + idx  # Start at column 43 (corrected)
                if col_position <= 70:  # Within the range
                    try:
                        lookup_date = date_str
                        row_data[col_position] = market_cap_ff_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # TTM Revenue data - columns 72-107 (matches corrected import logic)
            ttm_data = {}
            for ttm in stock.ttm_data.all():
                ttm_data[ttm.period] = ttm
            
            ttm_periods = column_mapping['ttm_periods']
            for idx, period in enumerate(ttm_periods):
                col_position = 72 + idx  # Start at column 72 (corrected)
                if col_position <= 107:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_revenue if ttm_obj else None
            
            # TTM Revenue Free Float data - columns 108-143 (matches corrected import logic)  
            for idx, period in enumerate(ttm_periods):
                col_position = 108 + idx  # Start at column 108 (corrected)
                if col_position <= 143:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_revenue_free_float if ttm_obj else None
            
            # TTM PAT data - columns 144-179 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 144 + idx  # Start at column 144
                if col_position <= 179:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_pat if ttm_obj else None
            
            # TTM PAT Free Float data - columns 180-215 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 180 + idx  # Start at column 180 (corrected)
                if col_position <= 215:  # Within the range
                    ttm_obj = ttm_data.get(period)
                    row_data[col_position] = ttm_obj.ttm_pat_free_float if ttm_obj else None
            
            # Quarterly Revenue data - columns 216-251 (matches corrected import logic)
            quarterly_data = {}
            for qtr in stock.quarterly_data.all():
                quarterly_data[qtr.period] = qtr
            
            for idx, period in enumerate(ttm_periods):
                col_position = 216 + idx  # Start at column 216 (corrected)
                if col_position <= 251:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_revenue if qtr_obj else None
            
            # Quarterly Revenue Free Float data - columns 252-287 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 252 + idx  # Start at column 252 (corrected)
                if col_position <= 287:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_revenue_free_float if qtr_obj else None
            
            # Quarterly PAT data - columns 288-323 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 288 + idx  # Start at column 288 (corrected)
                if col_position <= 323:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_pat if qtr_obj else None
            
            # Quarterly PAT Free Float data - columns 324-359 (matches corrected import logic)
            for idx, period in enumerate(ttm_periods):
                col_position = 324 + idx  # Start at column 324 (corrected)
                if col_position <= 359:  # Within the range
                    qtr_obj = quarterly_data.get(period)
                    row_data[col_position] = qtr_obj.quarterly_pat_free_float if qtr_obj else None
            
            # ROCE data - columns 360-372 (matches corrected import logic)
            annual_data = {}
            for ratio in stock.annual_ratios.all():
                annual_data[ratio.financial_year] = ratio
            
            annual_years = column_mapping['annual_years']
            for idx, year in enumerate(annual_years):
                col_position = 360 + idx  # Start at column 360 (corrected)
                if col_position <= 372:  # Within the range
                    ratio_obj = annual_data.get(year)
                    row_data[col_position] = ratio_obj.roce_percentage if ratio_obj else None
            
            # ROE data - columns 373-385 (matches corrected import logic)
            for idx, year in enumerate(annual_years):
                col_position = 373 + idx  # Start at column 373 (corrected)
                if col_position <= 385:  # Within the range
                    ratio_obj = annual_data.get(year)
                    row_data[col_position] = ratio_obj.roe_percentage if ratio_obj else None
            
            # Retention data - columns 386-398 (matches corrected import logic)
            for idx, year in enumerate(annual_years):
                col_position = 386 + idx  # Start at column 386 (corrected)
                if col_position <= 398:  # Within the range
                    ratio_obj = annual_data.get(year)
                    row_data[col_position] = ratio_obj.retention_percentage if ratio_obj else None
            
            # Share Price data - columns 399-409 (matches corrected import logic)
            price_data = {}
            pe_data = {}
            for price in stock.price_data.all():
                price_data[price.price_date.strftime('%Y-%m-%d')] = price.share_price
                pe_data[price.price_date.strftime('%Y-%m-%d')] = price.pe_ratio
            
            price_dates = column_mapping['price_dates']
            for idx, date_str in enumerate(price_dates):
                col_position = 399 + idx  # Start at column 399 (corrected)
                if col_position <= 409:  # Within the range
                    try:
                        lookup_date = date_str  # Date already in correct format
                        row_data[col_position] = price_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # PE Ratio data - columns 410-420 (matches corrected import logic)
            for idx, date_str in enumerate(price_dates):
                col_position = 410 + idx  # Start at column 410 (corrected)
                if col_position <= 420:  # Within the range
                    try:
                        lookup_date = date_str  # Date already in correct format
                        row_data[col_position] = pe_data.get(lookup_date)
                    except:
                        row_data[col_position] = None
            
            # Identifiers - fixed positions (matches corrected import logic)
            row_data[422] = stock.bse_code
            row_data[423] = stock.nse_symbol
            row_data[424] = stock.isin
            
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
        response['Content-Disposition'] = f'attachment; filename="stocks_base_sheet_all_{timezone.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
        
        self.message_user(request, f"Successfully exported all {len(all_stocks_queryset)} stocks to Excel file.")
        return response
    
    export_all_stocks_base_sheet.short_description = "Export all stocks as Base Sheet Excel"

class StockMarketCapAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['stock__company_name', 'stock__accord_code']
    list_display = ('stock', 'date', 'market_cap', 'market_cap_free_float')
    list_filter = ['date', 'created']
    raw_id_fields = ['stock']
    readonly_fields = ['created', 'modified']

class StockTTMDataAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['stock__company_name', 'stock__accord_code']
    list_display = ('stock', 'period', 'ttm_revenue', 'ttm_pat')
    list_filter = ['period', 'created']
    raw_id_fields = ['stock']
    readonly_fields = ['created', 'modified']

class StockQuarterlyDataAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['stock__company_name', 'stock__accord_code']
    list_display = ('stock', 'period', 'quarterly_revenue', 'quarterly_pat')
    list_filter = ['period', 'created']
    raw_id_fields = ['stock']
    readonly_fields = ['created', 'modified']

class StockAnnualRatiosAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['stock__company_name', 'stock__accord_code']
    list_display = ('stock', 'financial_year', 'roce_percentage', 'roe_percentage', 'retention_percentage')
    list_filter = ['financial_year', 'created']
    raw_id_fields = ['stock']
    readonly_fields = ['created', 'modified']

class StockPriceAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['stock__company_name', 'stock__accord_code']
    list_display = ('stock', 'price_date', 'share_price', 'pe_ratio')
    list_filter = ['price_date', 'created']
    raw_id_fields = ['stock']
    readonly_fields = ['created', 'modified']

class FundHoldingAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['scheme__name', 'stock__company_name', 'stock__accord_code']
    list_display = ('scheme', 'stock', 'holding_date', 'holding_percentage', 'market_value', 'number_of_shares')
    list_filter = ['holding_date', 'created']
    raw_id_fields = ['scheme', 'stock']
    readonly_fields = ['fund_holding_id', 'created', 'modified']

# Register the model with the custom admin class
admin.site.register(Customer, CustomerAdmin)
admin.site.register(AMCFundScheme, AMCFundSchemeAdmin)
admin.site.register(AMCFundSchemeNavLog, AMCFundSchemeNavLogAdmin)

# Register Stock models
admin.site.register(Stock, StockAdmin)
admin.site.register(StockMarketCap, StockMarketCapAdmin)
admin.site.register(StockTTMData, StockTTMDataAdmin)
admin.site.register(StockQuarterlyData, StockQuarterlyDataAdmin)
admin.site.register(StockAnnualRatios, StockAnnualRatiosAdmin)
admin.site.register(StockPrice, StockPriceAdmin)
admin.site.register(FundHolding, FundHoldingAdmin)


