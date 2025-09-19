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

    def first_name(self, obj):
        return obj.first_name
    first_name.short_description = 'first_name'

    def generate_random_password(modeladmin, request, queryset):
        for customer in queryset:
            password = get_random_string(8)
            customer.password = make_password(password)
            customer.save()
            messages.success(request, f'New Password for {customer.first_name}: {password}')

    generate_random_password.short_description = "Generate Random Password"

    # Add the action to the list of actions
    # actions = [generate_random_password]

class AMCFundSchemeAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['name', 'fund_name', 'amfi_scheme_code', 'isin_number', 'accord_mf_name']
    list_display = ('amcfundscheme_id', 'name', 'fund_name', 'fund_class', 'latest_nav', 'is_active')
    list_filter = ['fund_class', 'is_direct_fund', 'is_active', 'launch_date']
    readonly_fields = ['amcfundscheme_id']


class AMCFundSchemeNavLogAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['amcfundscheme__name', 'amcfundscheme__fund_name']
    list_display = ('amcfundscheme', 'as_on_date', 'nav')
    list_filter = ['as_on_date']
    raw_id_fields = ['amcfundscheme']

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

        # Import the block-based export generator
        from .dynamic_admin_export import BlockBasedExportGenerator

        # Create workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Stocks Base Sheet"

        # Generate BLOCK-BASED export structure for scalable period handling
        block_generator = BlockBasedExportGenerator()
        export_structure = block_generator.get_complete_block_based_export_structure()

        header_structure = export_structure['header_structure']
        block_mapping = export_structure['block_mapping']
        periods = export_structure['periods']
        total_columns = export_structure['total_columns']
        generator_instance = export_structure['generator_instance']

        print(f"=== BLOCK-BASED ADMIN EXPORT (Scalable Template System) ===")
        print(f"Total columns: {total_columns} (block-based calculation)")
        print(f"Market cap dates: {len(periods['market_cap_dates'])}")
        print(f"TTM periods: {len(periods['ttm_periods'])}")
        print(f"Quarterly periods: {len(periods['quarterly_periods'])}")
        print(f"Annual years: {len(periods['annual_years'])}")
        print(f"Price dates: {len(periods['price_dates'])}")
        print(f"Total blocks: {len(block_mapping)}")

        # Add all 8 header rows with DYNAMIC structure
        ws.append(header_structure['row_1'])  # Row 1: Empty
        ws.append(header_structure['row_2'])  # Row 2: Column numbers
        ws.append(header_structure['row_3'])  # Row 3: Field descriptions
        ws.append(header_structure['row_4'])  # Row 4: Shareholding info
        ws.append(header_structure['row_5'])  # Row 5: Formulas
        ws.append(header_structure['row_6'])  # Row 6: Category descriptions
        ws.append(header_structure['row_7'])  # Row 7: Sub-category details
        ws.append(header_structure['row_8'])  # Row 8: DYNAMIC column headers with actual periods

        # Add stock data using BLOCK-BASED POSITIONING
        for i, stock in enumerate(queryset, 1):
            # Get organized stock data by blocks
            stock_data = generator_instance.get_stock_data_by_blocks(stock, block_mapping)

            # Populate row using block-based positioning (scalable approach)
            row_data = generator_instance.populate_stock_row_by_blocks(
                stock, stock_data, block_mapping, i
            )

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
        response['Content-Disposition'] = f'attachment; filename="stocks_base_sheet_block_based_{timezone.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'

        self.message_user(request, f"Successfully exported {len(queryset)} stocks using BLOCK-BASED scalable system to Excel file.")
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
    list_display = ('stock', 'price_date', 'share_price', 'pr_ratio', 'pe_ratio')
    list_filter = ['price_date', 'created']
    raw_id_fields = ['stock']
    readonly_fields = ['created', 'modified']

class FundHoldingAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['scheme__name', 'stock__company_name', 'stock__accord_code']
    list_display = ('scheme', 'stock', 'holding_date', 'holding_percentage', 'market_value', 'number_of_shares')
    list_filter = ['holding_date', 'created']
    raw_id_fields = ['scheme', 'stock']
    readonly_fields = ['fund_holding_id', 'created', 'modified']

class FileStructureMetadataAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['upload_session_id', 'original_filename']
    list_display = ('original_filename', 'upload_session_id', 'total_columns', 'import_status', 'records_imported', 'created')
    list_filter = ['import_status', 'created']
    readonly_fields = ['file_structure_id', 'created', 'modified']

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
admin.site.register(FileStructureMetadata, FileStructureMetadataAdmin)