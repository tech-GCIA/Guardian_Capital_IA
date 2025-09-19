# Register your models here.
from django.contrib import admin, messages
from django.http import HttpResponse
from gcia_app.models import *
from .models import Stock
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

    def export_all_stocks_base_sheet(self, request, queryset):
    """Export all stocks data in Base Sheet format with DYNAMIC periods"""
    from django.utils import timezone

    # Get all stocks instead of just the queryset
    all_stocks_queryset = Stock.objects.all()

    # Use the same dynamic logic but with all stocks
    return self.export_stocks_base_sheet(request, all_stocks_queryset)

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


