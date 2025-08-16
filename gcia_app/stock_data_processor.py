# Create a new file: gcia_app/stock_data_processor.py

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
import re
from django.db import transaction
from django.utils import timezone
from gcia_app.models import Stock, StockQuarterlyData, StockUploadLog
import logging

logger = logging.getLogger(__name__)

class StockDataProcessor:
    """
    Processes Excel files containing stock data and updates the database
    """
    
    def __init__(self, user):
        self.user = user
        self.stats = {
            'stocks_added': 0,
            'stocks_updated': 0,
            'quarterly_records_added': 0,
            'quarterly_records_updated': 0,
            'errors': []
        }
    
    def process_excel_file(self, excel_file, upload_log):
        """
        Main method to process the uploaded Excel file
        """
        start_time = timezone.now()
        
        try:
            upload_log.status = 'processing'
            upload_log.processing_started_at = start_time
            upload_log.save()
            
            # Load the workbook
            workbook = load_workbook(excel_file, data_only=True)
            
            # Look for the App-Base Sheet
            app_base_sheet = self._find_app_base_sheet(workbook)
            if not app_base_sheet:
                raise ValueError("App-Base Sheet not found in the uploaded file")
            
            # Process the data
            with transaction.atomic():
                self._process_app_base_sheet(app_base_sheet)
            
            # Update upload log with success
            upload_log.status = 'completed'
            upload_log.stocks_added = self.stats['stocks_added']
            upload_log.stocks_updated = self.stats['stocks_updated']
            upload_log.quarterly_records_added = self.stats['quarterly_records_added']
            upload_log.quarterly_records_updated = self.stats['quarterly_records_updated']
            
        except Exception as e:
            logger.error(f"Error processing Excel file: {str(e)}")
            upload_log.status = 'failed'
            upload_log.error_message = str(e)
            self.stats['errors'].append(str(e))
            raise
        
        finally:
            end_time = timezone.now()
            upload_log.processing_completed_at = end_time
            upload_log.processing_time = end_time - start_time
            upload_log.save()
        
        return self.stats
    
    def _find_app_base_sheet(self, workbook):
        """
        Find the App-Base Sheet in the workbook
        """
        # Look for exact match first
        if "App-Base Sheet" in workbook.sheetnames:
            return workbook["App-Base Sheet"]
        
        # Common names for the app-base sheet
        possible_names = [
            'App-Base Sheet', 'App Base Sheet', 'AppBase Sheet', 
            'App-Base', 'AppBase', 'Base Sheet', 'Stock Data', 'Data'
        ]
        
        for sheet_name in workbook.sheetnames:
            if any(name.lower() in sheet_name.lower() for name in possible_names):
                return workbook[sheet_name]
        
        # If not found, use the first sheet
        if workbook.sheetnames:
            return workbook[workbook.sheetnames[0]]
        
        return None
    
    def _process_app_base_sheet(self, sheet):
        """
        Process the App-Base Sheet data based on the actual structure
        """
        # Convert sheet to list of lists for easier processing
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(list(row))
        
        if len(data) < 10:
            raise ValueError("Sheet appears to be empty or has insufficient data")
        
        # Based on analysis, data starts from row 8 (index 7) for headers
        # and row 9 (index 8) for actual data
        header_row_index = 7  # Row 8 in Excel
        data_start_index = 9   # Row 10 in Excel (skip the sample row)
        
        # Process each data row
        for row_index in range(data_start_index, len(data)):
            try:
                row_data = data[row_index]
                if not row_data or not any(cell for cell in row_data[:10] if cell):
                    continue  # Skip empty rows
                
                self._process_stock_row(row_data, row_index + 1)  # +1 for Excel row number
                
            except Exception as e:
                error_msg = f"Error processing row {row_index + 1}: {str(e)}"
                logger.warning(error_msg)
                self.stats['errors'].append(error_msg)
    
    def _process_stock_row(self, row_data, excel_row_num):
        """
        Process a single stock row from the Excel data
        Based on the Excel structure analysis:
        Column 1: Company Name
        Column 2: Accord Code (will use as symbol)
        Column 3: Sector
        Column 4: Cap (market cap category)
        Column 5: Free Float
        """
        # Extract basic stock information
        company_name = str(row_data[1]).strip() if row_data[1] else None
        accord_code = str(row_data[2]).strip() if row_data[2] else None
        sector = str(row_data[3]).strip() if row_data[3] else None
        cap_category = str(row_data[4]).strip() if row_data[4] else None
        free_float = self._convert_to_decimal(row_data[5]) if len(row_data) > 5 else None
        
        # Skip if essential data is missing
        if not company_name or not accord_code or company_name.startswith("Sample"):
            return
        
        # Create stock symbol from accord code
        symbol = f"STOCK_{accord_code}"
        
        # Prepare stock data
        stock_data = {
            'name': company_name,
            'symbol': symbol,
            'sector': sector if sector and sector != 'Sample' else None,
            'market_cap_category': cap_category if cap_category and cap_category != 'Sample' else None,
            'is_active': True
        }
        
        # Get or create stock
        stock, created = Stock.objects.get_or_create(
            symbol=symbol,
            defaults=stock_data
        )
        
        if created:
            self.stats['stocks_added'] += 1
            logger.info(f"Added new stock: {company_name} ({symbol})")
        else:
            # Update existing stock data
            updated = False
            for field, value in stock_data.items():
                if field != 'symbol' and hasattr(stock, field):
                    current_value = getattr(stock, field)
                    if current_value != value and value is not None:
                        setattr(stock, field, value)
                        updated = True
            
            if updated:
                stock.save()
                self.stats['stocks_updated'] += 1
                logger.info(f"Updated stock: {company_name} ({symbol})")
        
        # Process quarterly financial data
        self._process_quarterly_financial_data(stock, row_data)
    
    def _process_quarterly_financial_data(self, stock, row_data):
        """
        Process quarterly financial data for the stock
        Based on Excel structure:
        - Market Cap data: columns 14-41 (dates), with corresponding values
        - TTM Revenue: columns 74-108 
        - TTM PAT: columns 142-176
        - Quarter codes: columns 74-108 contain YYYYMM format (202503, 202412, etc.)
        """
        
        # Define the quarter mappings based on Excel analysis
        quarter_mappings = [
            # Recent quarters first (Latest data)
            {'year': 2025, 'quarter': 1, 'mcap_col': 14, 'revenue_col': 74, 'pat_col': 142},  # Q1 2025 (Jan-Mar 2025)
            {'year': 2024, 'quarter': 4, 'mcap_col': 15, 'revenue_col': 75, 'pat_col': 143},  # Q4 2024 (Oct-Dec 2024)
            {'year': 2024, 'quarter': 3, 'mcap_col': 16, 'revenue_col': 76, 'pat_col': 144},  # Q3 2024 (Jul-Sep 2024)
            {'year': 2024, 'quarter': 2, 'mcap_col': 17, 'revenue_col': 77, 'pat_col': 145},  # Q2 2024 (Apr-Jun 2024)
            {'year': 2024, 'quarter': 1, 'mcap_col': 18, 'revenue_col': 78, 'pat_col': 146},  # Q1 2024 (Jan-Mar 2024)
            {'year': 2023, 'quarter': 4, 'mcap_col': 19, 'revenue_col': 79, 'pat_col': 147},  # Q4 2023
            {'year': 2023, 'quarter': 3, 'mcap_col': 20, 'revenue_col': 80, 'pat_col': 148},  # Q3 2023
            {'year': 2023, 'quarter': 2, 'mcap_col': 21, 'revenue_col': 81, 'pat_col': 149},  # Q2 2023
            {'year': 2023, 'quarter': 1, 'mcap_col': 22, 'revenue_col': 82, 'pat_col': 150},  # Q1 2023
            {'year': 2022, 'quarter': 4, 'mcap_col': 23, 'revenue_col': 83, 'pat_col': 151},  # Q4 2022
        ]
        
        # Process each quarter's data
        for quarter_info in quarter_mappings:
            try:
                year = quarter_info['year']
                quarter_num = quarter_info['quarter']
                
                # Extract financial data for this quarter
                mcap = None
                revenue = None
                pat = None
                
                # Get Market Cap (in crores)
                if quarter_info['mcap_col'] < len(row_data):
                    mcap = self._convert_to_decimal(row_data[quarter_info['mcap_col']])
                
                # Get Revenue (TTM)
                if quarter_info['revenue_col'] < len(row_data):
                    revenue = self._convert_to_decimal(row_data[quarter_info['revenue_col']])
                
                # Get PAT (TTM)
                if quarter_info['pat_col'] < len(row_data):
                    pat = self._convert_to_decimal(row_data[quarter_info['pat_col']])
                
                # Only create record if we have at least one piece of financial data
                if mcap is not None or revenue is not None or pat is not None:
                    quarter_date = self._get_quarter_end_date(year, quarter_num)
                    
                    # Create or update quarterly data
                    quarterly_data, created = StockQuarterlyData.objects.get_or_create(
                        stock=stock,
                        quarter_year=year,
                        quarter_number=quarter_num,
                        defaults={
                            'quarter_date': quarter_date,
                            'mcap': mcap,
                            'revenue': revenue,
                            'pat': pat,
                            'ttm': revenue,  # TTM is essentially the revenue for this context
                        }
                    )
                    
                    if created:
                        self.stats['quarterly_records_added'] += 1
                    else:
                        # Update existing record if new data is different
                        updated = False
                        
                        if mcap is not None and quarterly_data.mcap != mcap:
                            quarterly_data.mcap = mcap
                            updated = True
                        
                        if revenue is not None and quarterly_data.revenue != revenue:
                            quarterly_data.revenue = revenue
                            quarterly_data.ttm = revenue
                            updated = True
                        
                        if pat is not None and quarterly_data.pat != pat:
                            quarterly_data.pat = pat
                            updated = True
                        
                        if updated:
                            quarterly_data.save()
                            self.stats['quarterly_records_updated'] += 1
                            
            except Exception as e:
                error_msg = f"Error processing Q{quarter_num} {year} for {stock.name}: {str(e)}"
                logger.warning(error_msg)
                self.stats['errors'].append(error_msg)
    
    def _convert_to_decimal(self, value):
        """
        Convert various number formats to Decimal
        """
        if value is None or value == '' or value == 0:
            return None
        
        try:
            # Handle string values that might contain commas, currency symbols, etc.
            if isinstance(value, str):
                # Remove common currency symbols and formatting
                cleaned = re.sub(r'[₹$,\s]', '', value.strip())
                # Handle parentheses for negative numbers
                if cleaned.startswith('(') and cleaned.endswith(')'):
                    cleaned = '-' + cleaned[1:-1]
                
                if not cleaned or cleaned == '-' or cleaned.lower() == 'sample':
                    return None
                
                return Decimal(cleaned)
            
            # Handle numeric values
            return Decimal(str(value))
            
        except (ValueError, InvalidOperation, TypeError):
            return None
    
    def _get_quarter_end_date(self, year, quarter):
        """
        Get the end date for a given quarter and year
        """
        quarter_end_months = {1: 3, 2: 6, 3: 9, 4: 12}  # March, June, Sep, Dec
        month = quarter_end_months[quarter]
        
        # Get last day of the month
        if month == 12:
            next_month = datetime(year + 1, 1, 1)
        else:
            next_month = datetime(year, month + 1, 1)
        
        last_day = next_month - timedelta(days=1)
        return last_day.date()


def process_stock_data_file(excel_file, user):
    """
    Main function to process stock data Excel file
    """
    # Create upload log entry
    upload_log = StockUploadLog.objects.create(
        uploaded_by=user,
        filename=excel_file.name,
        file_size=excel_file.size,
        status='pending'
    )
    
    try:
        processor = StockDataProcessor(user)
        stats = processor.process_excel_file(excel_file, upload_log)
        return stats, upload_log
    except Exception as e:
        logger.error(f"Failed to process stock data file: {str(e)}")
        raise