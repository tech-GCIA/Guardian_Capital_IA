import os
from django.core.management.base import BaseCommand
from gcia_app.views import validate_excel_structure, process_stocks_base_sheet
from gcia_app.models import Stock

class Command(BaseCommand):
    help = 'Test stock data processing and show diagnostics'
    
    def add_arguments(self, parser):
        parser.add_argument('--validate-only', action='store_true', help='Only validate structure, do not process data')
        parser.add_argument('--file', type=str, default='Base Sheet.xlsx', help='Excel file to process')
        
    def handle(self, *args, **options):
        excel_file = options['file']
        
        if not os.path.exists(excel_file):
            self.stdout.write(self.style.ERROR(f'File not found: {excel_file}'))
            return
        
        self.stdout.write(self.style.SUCCESS(f'Testing stock processing with file: {excel_file}'))
        
        # Validate structure
        self.stdout.write('\n=== VALIDATING EXCEL STRUCTURE ===')
        diagnostics = validate_excel_structure(excel_file)
        
        for key, value in diagnostics.items():
            if key == 'header_row_8':
                self.stdout.write(f'{key}: {value[:10]}... (first 10 items)' if len(value) > 10 else f'{key}: {value}')
            elif key == 'sample_data_row':
                self.stdout.write(f'{key}: {value[:10]}... (first 10 items)' if len(value) > 10 else f'{key}: {value}')
            else:
                self.stdout.write(f'{key}: {value}')
        
        if options['validate_only']:
            return
            
        # Show current stock count
        initial_count = Stock.objects.count()
        self.stdout.write(f'\nInitial stock count: {initial_count}')
        
        # Process data
        self.stdout.write('\n=== PROCESSING STOCK DATA ===')
        try:
            stats = process_stocks_base_sheet(excel_file)
            self.stdout.write(self.style.SUCCESS('\nProcessing completed!'))
            
            for key, value in stats.items():
                self.stdout.write(f'{key}: {value}')
                
            # Show final stock count
            final_count = Stock.objects.count()
            self.stdout.write(f'\nFinal stock count: {final_count}')
            
            # Show sample data from first few stocks
            self.stdout.write('\n=== SAMPLE STOCK DATA ===')
            for i, stock in enumerate(Stock.objects.all()[:3], 1):
                self.stdout.write(f'Stock {i}: {stock.company_name} ({stock.accord_code})')
                self.stdout.write(f'  Free Float: {stock.free_float}')
                self.stdout.write(f'  Revenue TTM: {stock.revenue_ttm}')
                self.stdout.write(f'  Market Cap Records: {stock.market_cap_data.count()}')
                self.stdout.write(f'  TTM Records: {stock.ttm_data.count()}')
                self.stdout.write('')
                
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Error processing file: {str(e)}'))
            import traceback
            self.stdout.write(traceback.format_exc())