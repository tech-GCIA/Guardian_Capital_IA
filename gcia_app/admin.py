from django.contrib import admin

# Register your models here.
from django.contrib import admin
from gcia_app.models import *
from import_export.admin import ImportExportModelAdmin

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

    # Custom queryset for displaying records
    def get_queryset(self, request):
        queryset = super().get_queryset(request)
        return queryset

class AMCAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['name']
    list_display = ('amc_id', 'name', 'assets_under_management', 'is_active', 'registrartransferagency')
    readonly_fields = ['amc_id', 'created', 'modified','assets_under_management']

    list_filter = ['is_active']

    def get_queryset(self, request):
        if 'change' not in request.get_full_path():
            return super(AMCAdmin, self).get_queryset(request).only('name','is_active','assets_under_management', 'registrartransferagency', 'created', 'modified')
        else:
            return super(AMCAdmin, self).get_queryset(request)
        
class AMCFundAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['name']
    list_display = ('amcfund_id', 'name', 'fund_class', 'assets_under_management', 'is_active', 'launch_date')
    readonly_fields = ['amcfund_id', 'created', 'modified','assets_under_management']

    list_filter = ['is_active']
    raw_id_fields = ['AMC']

    def get_queryset(self, request):
        if 'change' not in request.get_full_path():
            return super(AMCFundAdmin, self).get_queryset(request).select_related('AMC')
        else:
            return super(AMCFundAdmin, self).get_queryset(request)
        
class AMCFundSchemeAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
    search_fields = ['name']
    list_display = ('amcfundscheme_id', 'name', 'isin_div_or_growth_code', 'scheme_benchmark', 'is_active', 'latest_nav', 'latest_nav_as_on_date')
    readonly_fields = ['amcfundscheme_id', 'created', 'modified','assets_under_management']

    list_filter = ['is_active', 'is_direct_fund', 'is_scheme_benchmark']
    raw_id_fields = ['AMCFund']

    def get_queryset(self, request):
        if 'change' not in request.get_full_path():
            return super(AMCFundSchemeAdmin, self).get_queryset(request).select_related('AMCFund', 'AMCFund__AMC')
        else:
            return super(AMCFundSchemeAdmin, self).get_queryset(request)
        
# class AMCFundSchemeNavLogAdmin(ImportExportModelAdmin, ReadOnlyModelAdmin):
#     search_fields = ['amcfundscheme__name']
#     list_display = ('amcfundschemenavlog_id', 'amcfundscheme__name', 'as_on_date', 'nav')

#     raw_id_fields = ['amcfundscheme']

#     def get_queryset(self, request):
#         if 'change' not in request.get_full_path():
#             return super(AMCFundSchemeNavLogAdmin, self).get_queryset(request).select_related('amcfundscheme')
#         else:
#             return super(AMCFundSchemeNavLogAdmin, self).get_queryset(request)

# Register the model with the custom admin class
admin.site.register(Customer, CustomerAdmin)
admin.site.register(AMC, AMCAdmin)
admin.site.register(AMCFund, AMCFundAdmin)
admin.site.register(AMCFundScheme, AMCFundSchemeAdmin)
# admin.site.register(AMCFundSchemeNavLog, AMCFundSchemeNavLogAdmin)


