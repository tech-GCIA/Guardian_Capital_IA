# Register your models here.
from django.contrib import admin, messages
from gcia_app.models import *
from import_export.admin import ImportExportModelAdmin
import string
import random
from django.utils.crypto import get_random_string
from django.contrib.auth.hashers import make_password

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
    list_display = ('amcfundscheme_id', 'name', 'isin_div_or_growth_code', 'scheme_benchmark', 'is_active', 'latest_nav', 'latest_nav_as_on_date')
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

# Register the model with the custom admin class
admin.site.register(Customer, CustomerAdmin)
admin.site.register(AMCFundScheme, AMCFundSchemeAdmin)
admin.site.register(AMCFundSchemeNavLog, AMCFundSchemeNavLogAdmin)

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
