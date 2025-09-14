# Register your models here.
from django.contrib import admin, messages
from gcia_app.models import *
from import_export.admin import ImportExportModelAdmin
import string
import random
from django.utils.crypto import get_random_string
from django.contrib.auth.hashers import make_password
from django.http import HttpResponse
from datetime import datetime
import pandas as pd
import tempfile
import os

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
        
class SchemeUnderlyingHoldingsAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['amcfundscheme__name', 'amcfundscheme__accord_scheme_name']
    list_display = ('schemeunderlyingholding_id', 'amcfundscheme__name', 'holding__name', 'weightage')

    raw_id_fields = ['amcfundscheme', 'holding']

    def get_queryset(self, request):
        if 'change' not in request.get_full_path():
            return super(SchemeUnderlyingHoldingsAdmin, self).get_queryset(request).select_related('amcfundscheme', 'holding')
        else:
            return super(SchemeUnderlyingHoldingsAdmin, self).get_queryset(request)

# Register the model with the custom admin class
admin.site.register(Customer, CustomerAdmin)
admin.site.register(AMCFundScheme, AMCFundSchemeAdmin)
admin.site.register(AMCFundSchemeNavLog, AMCFundSchemeNavLogAdmin)
admin.site.register(SchemeUnderlyingHoldings, SchemeUnderlyingHoldingsAdmin)

# Add this to gcia_app/admin.py

from django.contrib import admin
from gcia_app.models import Stock, StockQuarterlyData, StockUploadLog

@admin.register(Stock)
class StockAdmin(admin.ModelAdmin):
    list_display = ('stock_id', 'name', 'symbol', 'sector', 'industry', 'is_active', 'created')
    list_filter = ('sector', 'industry', 'market_cap_category', 'is_active', 'created')
    search_fields = ('name', 'symbol', 'sector', 'industry', 'isin')
    readonly_fields = ('stock_id', 'created', 'modified')
    list_per_page = 100
    ordering = ('name',)
    
    fieldsets = (
        ('Basic Information', {
            'fields': ('name', 'symbol', 'isin', 'is_active')
        }),
        ('Classification', {
            'fields': ('sector', 'industry', 'market_cap_category')
        }),
        ('Dates', {
            'fields': ('listing_date',)
        }),
        ('System Fields', {
            'fields': ('stock_id', 'created', 'modified'),
            'classes': ('collapse',)
        })
    )


@admin.register(StockQuarterlyData)
class StockQuarterlyDataAdmin(admin.ModelAdmin):
    list_display = (
        'quarterly_data_id', 'stock_symbol', 'quarter_label', 
        'mcap', 'pat', 'pe_ratio', 'created'
    )
    list_filter = (
        'quarter_year', 'quarter_number', 'stock__sector', 
        'stock__industry', 'created'
    )
    search_fields = ('stock__name', 'stock__symbol')
    readonly_fields = ('quarterly_data_id', 'created', 'modified')
    list_per_page = 100
    ordering = ('-quarter_year', '-quarter_number', 'stock__name')
    raw_id_fields = ('stock',)
    
    fieldsets = (
        ('Stock & Quarter Info', {
            'fields': ('stock', 'quarter_year', 'quarter_number', 'quarter_date')
        }),
        ('Market Metrics', {
            'fields': ('mcap', 'price', 'pe_ratio', 'pb_ratio')
        }),
        ('Financial Metrics', {
            'fields': ('revenue', 'ebitda', 'net_profit', 'pat', 'ttm')
        }),
        ('Ratios & Returns', {
            'fields': ('book_value', 'dividend_yield', 'roe', 'roa', 'debt_to_equity')
        }),
        ('System Fields', {
            'fields': ('quarterly_data_id', 'created', 'modified'),
            'classes': ('collapse',)
        })
    )
    
    def stock_symbol(self, obj):
        return obj.stock.symbol
    stock_symbol.short_description = 'Symbol'
    stock_symbol.admin_order_field = 'stock__symbol'
    
    def quarter_label(self, obj):
        return f"Q{obj.quarter_number}-{obj.quarter_year}"
    quarter_label.short_description = 'Quarter'
    quarter_label.admin_order_field = 'quarter_year'
    
    actions = ['export_stocks_data_to_excel']
    
    def export_stocks_data_to_excel(self, request, queryset):
        """
        Admin action to export all stocks and quarterly data to Excel
        in the same format as the Base Sheet upload
        """
        try:
            # Get all stocks and their data
            all_stocks = Stock.objects.filter(is_active=True).order_by('name')
            
            if not all_stocks.exists():
                self.message_user(request, "No active stocks found to export.", level=messages.WARNING)
                return
            
            # Create a temporary file for the Excel export
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
                temp_path = tmp_file.name
            
            try:
                # Create Excel writer
                with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                    # Prepare data structure similar to Base Sheet
                    export_data = []
                    
                    for idx, stock in enumerate(all_stocks, 1):
                        # Get all quarterly data for this stock
                        quarterly_data = StockQuarterlyData.objects.filter(
                            stock=stock
                        ).order_by('-quarter_date')
                        
                        # Create base row with stock information
                        row_data = {
                            'S. No.': idx,
                            'Company Name': stock.name,
                            'Accord Code': stock.accord_code or '',
                            'Sector': stock.sector or '',
                            'Cap': stock.cap or '',
                            'Free Float': float(stock.free_float) if stock.free_float else '',
                            '6 Year CAGR': float(stock.revenue_6yr_cagr) if stock.revenue_6yr_cagr else '',
                            'TTM': float(stock.revenue_ttm) if stock.revenue_ttm else '',
                            '6 Year CAGR.1': float(stock.pat_6yr_cagr) if stock.pat_6yr_cagr else '',
                            'TTM.1': float(stock.pat_ttm) if stock.pat_ttm else '',
                            'Current': float(stock.current_pr) if stock.current_pr else '',
                            '2 Yr Avg': float(stock.pr_2yr_avg) if stock.pr_2yr_avg else '',
                            'Reval/deval': float(stock.reval_deval) if stock.reval_deval else '',
                            'BSE Code': stock.bse_code or '',
                            'NSE Code': stock.nse_code or '',
                            'ISIN': stock.isin or ''
                        }
                        
                        # Add quarterly data columns
                        for qdata in quarterly_data:
                            date_key = qdata.quarter_date.strftime('%Y-%m-%d')
                            
                            # Add different metrics based on what's available
                            if qdata.mcap:
                                row_data[f'Market_Cap_{date_key}'] = float(qdata.mcap)
                            if qdata.free_float_mcap:
                                row_data[f'Free_Float_MCap_{date_key}'] = float(qdata.free_float_mcap)
                            if qdata.ttm_revenue:
                                row_data[f'TTM_Revenue_{date_key}'] = float(qdata.ttm_revenue)
                            if qdata.pat:
                                row_data[f'PAT_{date_key}'] = float(qdata.pat)
                            if qdata.quarterly_revenue:
                                row_data[f'Quarterly_Revenue_{date_key}'] = float(qdata.quarterly_revenue)
                            if qdata.quarterly_pat:
                                row_data[f'Quarterly_PAT_{date_key}'] = float(qdata.quarterly_pat)
                            if qdata.roce:
                                row_data[f'ROCE_{date_key}'] = float(qdata.roce)
                            if qdata.roe:
                                row_data[f'ROE_{date_key}'] = float(qdata.roe)
                            if qdata.retention:
                                row_data[f'Retention_{date_key}'] = float(qdata.retention)
                            if qdata.share_price:
                                row_data[f'Share_Price_{date_key}'] = float(qdata.share_price)
                            if qdata.pr_quarterly:
                                row_data[f'PR_{date_key}'] = float(qdata.pr_quarterly)
                            if qdata.pe_quarterly:
                                row_data[f'PE_{date_key}'] = float(qdata.pe_quarterly)
                        
                        export_data.append(row_data)
                    
                    # Create DataFrame
                    df = pd.DataFrame(export_data)
                    
                    # Write to Excel with proper sheet name
                    df.to_excel(writer, sheet_name='App-Base Sheet', index=False)
                    
                    # Get the worksheet to add headers
                    worksheet = writer.sheets['App-Base Sheet']
                    
                    # Add multi-level headers similar to original format
                    # Insert rows at the top for headers
                    worksheet.insert_rows(1, 7)
                    
                    # Add header information
                    worksheet['A6'] = 'Stock wise Fundamentals and Valuations'
                    worksheet['A7'] = 'Revenue (Q1 FY-26)'
                    worksheet['A8'] = 'S. No.'
                    
                # Read the file and create response
                with open(temp_path, 'rb') as excel_file:
                    response = HttpResponse(
                        excel_file.read(),
                        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                    
                    # Set filename for download
                    filename = f"Stocks_Base_Sheet_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    response['Content-Disposition'] = f'attachment; filename="{filename}"'
                    
                # Clean up temporary file
                os.unlink(temp_path)
                
                # Success message
                self.message_user(
                    request,
                    f"Successfully exported {len(export_data)} stocks with their quarterly data.",
                    level=messages.SUCCESS
                )
                
                return response
                
            except Exception as e:
                # Clean up temp file on error
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
                raise e
                
        except Exception as e:
            self.message_user(
                request,
                f"Error exporting data: {str(e)}",
                level=messages.ERROR
            )
            return
    
    export_stocks_data_to_excel.short_description = "Export all stocks data to Excel (Base Sheet format)"


@admin.register(StockUploadLog)
class StockUploadLogAdmin(admin.ModelAdmin):
    list_display = (
        'upload_id', 'filename', 'uploaded_by', 'status', 
        'stocks_added', 'quarterly_records_added', 'uploaded_at'
    )
    list_filter = ('status', 'uploaded_at', 'uploaded_by')
    search_fields = ('filename', 'uploaded_by__username', 'uploaded_by__email')
    readonly_fields = (
        'upload_id', 'uploaded_at', 'processing_started_at', 
        'processing_completed_at', 'processing_time'
    )
    list_per_page = 50
    ordering = ('-uploaded_at',)
    
    fieldsets = (
        ('Upload Info', {
            'fields': ('filename', 'uploaded_by', 'file_size', 'status')
        }),
        ('Processing Results', {
            'fields': (
                'stocks_added', 'stocks_updated', 
                'quarterly_records_added', 'quarterly_records_updated'
            )
        }),
        ('Timing', {
            'fields': (
                'uploaded_at', 'processing_started_at', 
                'processing_completed_at', 'processing_time'
            )
        }),
        ('Error Details', {
            'fields': ('error_message',),
            'classes': ('collapse',)
        })
    )
    
    def has_add_permission(self, request):
        # Don't allow manual creation of upload logs
        return False
