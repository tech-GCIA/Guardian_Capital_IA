import os
import tempfile
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from django.conf import settings
from django.contrib.staticfiles.finders import find
from gcia_app.models import Stock, StockQuarterlyData
import logging

logger = logging.getLogger(__name__)

class FundReportGenerator:
    """
    Generate Excel report based on selected stocks and template file
    """
    
    def __init__(self):
        self.template_path = self._get_template_path()
        self.workbook = None
        self.sheets = {}
        
    def _get_template_path(self):
        """Get the path to the template Excel file"""
        # List of possible template filenames (with and without extra spaces)
        template_filenames = [
            "Format  Mutual Fund Review Q4FY25.xlsm",  # Original with double space
            "Format Mutual Fund Review Q4FY25.xlsm",   # Single space
            "Format - Mutual Fund Review (Q4-FY-25).xlsm",  # Alternative format
        ]
        
        for template_filename in template_filenames:
            # Try base directory first (same as manage.py)
            base_template_path = os.path.join(settings.BASE_DIR, template_filename)
            if os.path.exists(base_template_path):
                print(f"Found template file at: {base_template_path}")
                return base_template_path
            
            # Try static files
            template_path = find(f'gcia_app/templates/{template_filename}')
            if template_path and os.path.exists(template_path):
                print(f"Found template file at: {template_path}")
                return template_path
                
            # Try media root
            media_template_path = os.path.join(settings.MEDIA_ROOT, template_filename)
            if os.path.exists(media_template_path):
                print(f"Found template file at: {media_template_path}")
                return media_template_path
        
        # List all files in base directory for debugging
        base_files = []
        try:
            base_files = [f for f in os.listdir(settings.BASE_DIR) if f.endswith(('.xlsx', '.xlsm'))]
        except:
            pass
            
        error_msg = f"Template file not found. Searched for: {template_filenames}\n"
        error_msg += f"Base directory: {settings.BASE_DIR}\n"
        error_msg += f"Excel files in base directory: {base_files}\n"
        error_msg += f"Media root: {getattr(settings, 'MEDIA_ROOT', 'Not set')}"
        
        raise FileNotFoundError(error_msg)
    
    def generate_report(self, fund_name, selected_stocks):
        """
        Generate the complete Excel report
        
        Args:
            fund_name (str): Name of the fund
            selected_stocks (list): List of stock data with weightage and shares
            
        Returns:
            str: Path to the generated Excel file
        """
        try:
            # Load the template
            self.workbook = load_workbook(self.template_path, keep_vba=True)
            
            # Get all required sheets
            self._load_sheets()
            
            # Remove App-Base Sheet from the output as it's not needed in the final report
            if 'App-Base Sheet' in self.workbook.sheetnames:
                self.workbook.remove(self.workbook['App-Base Sheet'])
                logger.info("Removed App-Base Sheet from output - using database data instead")
            
            # Process selected stocks and get database data from Step 5.1
            stock_data = self._process_selected_stocks(selected_stocks)
            
            # Fill Portfolio Analysis sheet with DATABASE data (not template data)
            self._fill_portfolio_analysis(stock_data)
            
            # Fill Stock Weights sheet with DATABASE data
            self._fill_stock_weights(stock_data)
            
            # Fill Summary sheet with calculated statistics
            self._fill_summary(fund_name, stock_data)
            
            # Save the file
            output_path = self._save_report(fund_name)
            
            logger.info(f"Report generated successfully using database data from {len(stock_data)} selected stocks")
            
            return output_path
            
        except Exception as e:
            logger.error(f"Error generating report: {str(e)}")
            raise
    
    def _load_sheets(self):
        """Load all required sheets from the workbook"""
        required_sheets = ['App-Base Sheet', 'Portfolio Analysis', '2. Stock Weights', '1. Summary']
        
        for sheet_name in required_sheets:
            if sheet_name in self.workbook.sheetnames:
                self.sheets[sheet_name] = self.workbook[sheet_name]
            else:
                # Try alternative names
                for ws_name in self.workbook.sheetnames:
                    if 'portfolio' in ws_name.lower() and 'analysis' in ws_name.lower():
                        self.sheets['Portfolio Analysis'] = self.workbook[ws_name]
                    elif 'stock' in ws_name.lower() and 'weight' in ws_name.lower():
                        self.sheets['2. Stock Weights'] = self.workbook[ws_name]
                    elif 'summary' in ws_name.lower():
                        self.sheets['1. Summary'] = self.workbook[ws_name]
                    elif 'app' in ws_name.lower() and 'base' in ws_name.lower():
                        self.sheets['App-Base Sheet'] = self.workbook[ws_name]
        
        logger.info(f"Loaded sheets: {list(self.sheets.keys())}")
    
    def _process_selected_stocks(self, selected_stocks):
        """
        Process selected stocks and get all required data from DATABASE (Step 5.1 data)
        
        Args:
            selected_stocks (list): List of stock selections
            
        Returns:
            list: Enhanced stock data with database information from Step 5.1 upload
        """
        enhanced_data = []
        
        logger.info(f"Processing {len(selected_stocks)} selected stocks using database data from Step 5.1")
        
        for stock_selection in selected_stocks:
            stock_id = stock_selection['id']
            weightage = float(stock_selection['weightage'])
            shares = int(stock_selection['shares'])
            
            try:
                # Get stock from database (uploaded in Step 5.1)
                stock = Stock.objects.get(stock_id=stock_id)
                
                # Get quarterly data for this stock (uploaded in Step 5.1)
                quarterly_data = StockQuarterlyData.objects.filter(
                    stock=stock
                ).order_by('-quarter_year', '-quarter_number')[:20]  # Get more quarters for better analysis
                
                logger.info(f"Found {len(quarterly_data)} quarterly records for {stock.name} from database")
                
                # Calculate required metrics using DATABASE data
                metrics = self._calculate_stock_metrics(stock, quarterly_data)
                
                enhanced_data.append({
                    'stock': stock,
                    'weightage': weightage,
                    'shares': shares,
                    'quarterly_data': quarterly_data,
                    'metrics': metrics,
                    'data_source': 'database'  # Mark that this comes from Step 5.1 upload
                })
                
                logger.info(f"Successfully processed {stock.name} with {len(quarterly_data)} quarters of data")
                
            except Stock.DoesNotExist:
                logger.warning(f"Stock with ID {stock_id} not found in database - may need to upload stock data in Step 5.1")
                continue
        
        logger.info(f"Successfully processed {len(enhanced_data)} stocks using database data")
        return enhanced_data
    
    def _calculate_stock_metrics(self, stock, quarterly_data):
        """
        Calculate all required metrics for a stock
        """
        metrics = {}
        
        if not quarterly_data:
            return self._get_empty_metrics()
        
        # Convert quarterly data to list for easier processing
        quarters = list(quarterly_data)
        
        # Get current quarter data
        current_quarter = quarters[0] if quarters else None
        
        if current_quarter:
            metrics['current_mcap'] = float(current_quarter.mcap or 0)
            metrics['current_revenue'] = float(current_quarter.revenue or 0)
            metrics['current_pat'] = float(current_quarter.pat or 0)
            metrics['current_pe'] = float(current_quarter.pe_ratio or 0)
            metrics['current_pb'] = float(current_quarter.pb_ratio or 0)
        
        # Calculate QoQ (Quarter over Quarter) growth
        if len(quarters) >= 2:
            current_pat = float(quarters[0].pat or 0)
            previous_pat = float(quarters[1].pat or 0)
            metrics['qoq_growth'] = ((current_pat - previous_pat) / previous_pat * 100) if previous_pat > 0 else 0
        else:
            metrics['qoq_growth'] = 0
        
        # Calculate YoY (Year over Year) growth
        if len(quarters) >= 4:
            current_pat = float(quarters[0].pat or 0)
            year_ago_pat = float(quarters[3].pat or 0)
            metrics['yoy_growth'] = ((current_pat - year_ago_pat) / year_ago_pat * 100) if year_ago_pat > 0 else 0
        else:
            metrics['yoy_growth'] = 0
        
        # Calculate CAGR (6-year if available)
        metrics['cagr_6yr'] = self._calculate_cagr(quarters, 6)
        
        # Calculate PE statistics
        metrics.update(self._calculate_pe_statistics(quarters))
        
        # Calculate Price-to-Book statistics
        metrics.update(self._calculate_pb_statistics(quarters))
        
        # Calculate Alpha metrics
        metrics.update(self._calculate_alpha_metrics(quarters))
        
        return metrics
    
    def _calculate_cagr(self, quarters, years):
        """Calculate Compound Annual Growth Rate"""
        if len(quarters) < years * 4:  # Need at least years*4 quarters
            return 0
        
        current_value = float(quarters[0].revenue or 0)
        past_value = float(quarters[years * 4 - 1].revenue or 0)
        
        if past_value <= 0:
            return 0
        
        cagr = ((current_value / past_value) ** (1/years) - 1) * 100
        return round(cagr, 2)
    
    def _calculate_pe_statistics(self, quarters):
        """Calculate PE-related statistics"""
        pe_metrics = {}
        
        # Get PE ratios for calculations
        pe_ratios = [float(q.pe_ratio or 0) for q in quarters if q.pe_ratio]
        
        if pe_ratios:
            # Current PE
            pe_metrics['current_pe'] = pe_ratios[0] if pe_ratios else 0
            
            # 2-year average PE (8 quarters)
            two_year_pes = pe_ratios[:8] if len(pe_ratios) >= 8 else pe_ratios
            pe_metrics['pe_2yr_avg'] = sum(two_year_pes) / len(two_year_pes) if two_year_pes else 0
            
            # 5-year average PE (20 quarters)
            five_year_pes = pe_ratios[:20] if len(pe_ratios) >= 20 else pe_ratios
            pe_metrics['pe_5yr_avg'] = sum(five_year_pes) / len(five_year_pes) if five_year_pes else 0
            
            # Revaluation/Devaluation percentages
            if pe_metrics['pe_2yr_avg'] > 0:
                pe_metrics['pe_reval_2yr'] = ((pe_metrics['current_pe'] - pe_metrics['pe_2yr_avg']) / pe_metrics['pe_2yr_avg'] * 100)
            else:
                pe_metrics['pe_reval_2yr'] = 0
                
            if pe_metrics['pe_5yr_avg'] > 0:
                pe_metrics['pe_reval_5yr'] = ((pe_metrics['current_pe'] - pe_metrics['pe_5yr_avg']) / pe_metrics['pe_5yr_avg'] * 100)
            else:
                pe_metrics['pe_reval_5yr'] = 0
        else:
            pe_metrics = {
                'current_pe': 0,
                'pe_2yr_avg': 0,
                'pe_5yr_avg': 0,
                'pe_reval_2yr': 0,
                'pe_reval_5yr': 0
            }
        
        return pe_metrics
    
    def _calculate_pb_statistics(self, quarters):
        """Calculate Price-to-Book related statistics"""
        pb_metrics = {}
        
        # Get PB ratios for calculations
        pb_ratios = [float(q.pb_ratio or 0) for q in quarters if q.pb_ratio]
        
        if pb_ratios:
            # Current PB
            pb_metrics['current_pb'] = pb_ratios[0] if pb_ratios else 0
            
            # 2-year average PB (8 quarters)
            two_year_pbs = pb_ratios[:8] if len(pb_ratios) >= 8 else pb_ratios
            pb_metrics['pb_2yr_avg'] = sum(two_year_pbs) / len(two_year_pbs) if two_year_pbs else 0
            
            # 5-year average PB (20 quarters)
            five_year_pbs = pb_ratios[:20] if len(pb_ratios) >= 20 else pb_ratios
            pb_metrics['pb_5yr_avg'] = sum(five_year_pbs) / len(five_year_pbs) if five_year_pbs else 0
            
            # Revaluation/Devaluation percentages
            if pb_metrics['pb_2yr_avg'] > 0:
                pb_metrics['pb_reval_2yr'] = ((pb_metrics['current_pb'] - pb_metrics['pb_2yr_avg']) / pb_metrics['pb_2yr_avg'] * 100)
            else:
                pb_metrics['pb_reval_2yr'] = 0
                
            if pb_metrics['pb_5yr_avg'] > 0:
                pb_metrics['pb_reval_5yr'] = ((pb_metrics['current_pb'] - pb_metrics['pb_5yr_avg']) / pb_metrics['pb_5yr_avg'] * 100)
            else:
                pb_metrics['pb_reval_5yr'] = 0
                
            # 10 Quarter Low/High
            ten_quarter_pbs = pb_ratios[:10] if len(pb_ratios) >= 10 else pb_ratios
            pb_metrics['pb_10q_low'] = min(ten_quarter_pbs) if ten_quarter_pbs else 0
            pb_metrics['pb_10q_high'] = max(ten_quarter_pbs) if ten_quarter_pbs else 0
        else:
            pb_metrics = {
                'current_pb': 0,
                'pb_2yr_avg': 0,
                'pb_5yr_avg': 0,
                'pb_reval_2yr': 0,
                'pb_reval_5yr': 0,
                'pb_10q_low': 0,
                'pb_10q_high': 0
            }
        
        return pb_metrics
    
    def _calculate_alpha_metrics(self, quarters):
        """Calculate Alpha-related metrics"""
        alpha_metrics = {}
        
        # For simplicity, using basic calculations
        # In real implementation, you might want to compare against bond yields and market returns
        
        if quarters:
            # Assuming 6% bond rate for calculations
            bond_rate = 6.0
            
            # Calculate basic alpha (simplified)
            returns = []
            for i in range(len(quarters) - 1):
                current_pat = float(quarters[i].pat or 0)
                previous_pat = float(quarters[i+1].pat or 0)
                if previous_pat > 0:
                    returns.append((current_pat - previous_pat) / previous_pat * 100)
            
            avg_return = sum(returns) / len(returns) if returns else 0
            alpha_metrics['alpha_over_bond'] = avg_return - bond_rate
            alpha_metrics['absolute_alpha'] = avg_return
            
            # PE Yield (simplified as 1/PE * 100)
            current_pe = float(quarters[0].pe_ratio or 0)
            alpha_metrics['pe_yield'] = (1 / current_pe * 100) if current_pe > 0 else 0
            
        else:
            alpha_metrics = {
                'alpha_over_bond': 0,
                'absolute_alpha': 0,
                'pe_yield': 0
            }
        
        alpha_metrics['growth_rate'] = alpha_metrics.get('absolute_alpha', 0)
        alpha_metrics['bond_rate'] = 6.0  # Assumed bond rate
        
        return alpha_metrics
    
    def _get_empty_metrics(self):
        """Return empty metrics structure"""
        return {
            'current_mcap': 0,
            'current_revenue': 0,
            'current_pat': 0,
            'current_pe': 0,
            'current_pb': 0,
            'qoq_growth': 0,
            'yoy_growth': 0,
            'cagr_6yr': 0,
            'pe_2yr_avg': 0,
            'pe_5yr_avg': 0,
            'pe_reval_2yr': 0,
            'pe_reval_5yr': 0,
            'pb_2yr_avg': 0,
            'pb_5yr_avg': 0,
            'pb_reval_2yr': 0,
            'pb_reval_5yr': 0,
            'pb_10q_low': 0,
            'pb_10q_high': 0,
            'alpha_over_bond': 0,
            'absolute_alpha': 0,
            'pe_yield': 0,
            'growth_rate': 0,
            'bond_rate': 6.0
        }
    
    def _safe_set_cell_value(self, sheet, cell_address, value):
        """Safely set cell value, handling merged cells"""
        try:
            cell = sheet[cell_address]
            
            # Check if cell is part of a merged range
            if hasattr(cell, 'coordinate') and sheet.merged_cells:
                for merged_range in sheet.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        # This cell is part of a merged range, skip it
                        logger.warning(f"Skipping merged cell {cell_address}")
                        return False
            
            # Set the value
            cell.value = value
            return True
            
        except AttributeError as e:
            if "'MergedCell' object attribute 'value' is read-only" in str(e):
                logger.warning(f"Cannot write to merged cell {cell_address}, skipping")
                return False
            else:
                logger.error(f"Error setting cell {cell_address}: {str(e)}")
                return False
        except Exception as e:
            logger.error(f"Unexpected error setting cell {cell_address}: {str(e)}")
            return False
    
    def _find_data_start_row(self, sheet):
        """Find the appropriate starting row for data entry"""
        # Look for common header patterns to determine where data should start
        for row_num in range(1, 20):  # Check first 20 rows
            try:
                cell_value = sheet[f'A{row_num}'].value
                if cell_value and isinstance(cell_value, str):
                    cell_value_lower = cell_value.lower()
                    # Look for header-like content
                    if any(keyword in cell_value_lower for keyword in ['stock', 'company', 'name', 'symbol', 'security']):
                        return row_num + 1  # Return next row after header
            except:
                continue
        
        # Default fallback - start from row 5 if no header found
        return 5
    
    def _fill_portfolio_analysis(self, stock_data):
        """Fill the Portfolio Analysis sheet with safe cell writing"""
        if 'Portfolio Analysis' not in self.sheets:
            logger.warning("Portfolio Analysis sheet not found")
            return
        
        sheet = self.sheets['Portfolio Analysis']
        
        # Find the starting row for data by looking for headers or existing data
        start_row = self._find_data_start_row(sheet)
        
        logger.info(f"Starting to fill portfolio analysis from row {start_row}")
        
        for idx, stock_info in enumerate(stock_data):
            row = start_row + idx
            stock = stock_info['stock']
            metrics = stock_info['metrics']
            
            logger.info(f"Processing stock {idx + 1}: {stock.name} at row {row}")
            
            try:
                # Fill stock information - use safe cell writing
                self._safe_set_cell_value(sheet, f'A{row}', stock.name)  # Stock Name
                self._safe_set_cell_value(sheet, f'B{row}', stock.symbol)  # Stock Symbol
                self._safe_set_cell_value(sheet, f'C{row}', stock_info['weightage'])  # Weightage
                self._safe_set_cell_value(sheet, f'D{row}', stock_info['shares'])  # Number of Shares
                
                # Fill calculated metrics
                self._safe_set_cell_value(sheet, f'E{row}', metrics['current_mcap'])  # Market Cap
                self._safe_set_cell_value(sheet, f'F{row}', metrics['current_pat'])  # PAT
                self._safe_set_cell_value(sheet, f'G{row}', metrics['qoq_growth'])  # QoQ Growth
                self._safe_set_cell_value(sheet, f'H{row}', metrics['yoy_growth'])  # YoY Growth
                self._safe_set_cell_value(sheet, f'I{row}', metrics['cagr_6yr'])  # 6-Year CAGR
                
                # PE Statistics
                self._safe_set_cell_value(sheet, f'J{row}', metrics['current_pe'])  # Current PE
                self._safe_set_cell_value(sheet, f'K{row}', metrics['pe_2yr_avg'])  # 2-Year Avg PE
                self._safe_set_cell_value(sheet, f'L{row}', metrics['pe_5yr_avg'])  # 5-Year Avg PE
                self._safe_set_cell_value(sheet, f'M{row}', metrics['pe_reval_2yr'])  # PE Reval 2-Year
                self._safe_set_cell_value(sheet, f'N{row}', metrics['pe_reval_5yr'])  # PE Reval 5-Year
                
                # PB Statistics
                self._safe_set_cell_value(sheet, f'O{row}', metrics['current_pb'])  # Current PB
                self._safe_set_cell_value(sheet, f'P{row}', metrics['pb_2yr_avg'])  # 2-Year Avg PB
                self._safe_set_cell_value(sheet, f'Q{row}', metrics['pb_5yr_avg'])  # 5-Year Avg PB
                self._safe_set_cell_value(sheet, f'R{row}', metrics['pb_reval_2yr'])  # PB Reval 2-Year
                self._safe_set_cell_value(sheet, f'S{row}', metrics['pb_reval_5yr'])  # PB Reval 5-Year
                self._safe_set_cell_value(sheet, f'T{row}', metrics['pb_10q_low'])  # 10Q PB Low
                self._safe_set_cell_value(sheet, f'U{row}', metrics['pb_10q_high'])  # 10Q PB High
                
                # Alpha Metrics
                self._safe_set_cell_value(sheet, f'V{row}', metrics['alpha_over_bond'])  # Alpha over Bond
                self._safe_set_cell_value(sheet, f'W{row}', metrics['absolute_alpha'])  # Absolute Alpha
                self._safe_set_cell_value(sheet, f'X{row}', metrics['pe_yield'])  # PE Yield
                self._safe_set_cell_value(sheet, f'Y{row}', metrics['growth_rate'])  # Growth Rate
                self._safe_set_cell_value(sheet, f'Z{row}', metrics['bond_rate'])  # Bond Rate
                
                logger.info(f"Successfully filled data for {stock.name}")
                
            except Exception as e:
                logger.error(f"Error filling data for stock {stock.name} at row {row}: {str(e)}")
                continue
    
    def _fill_stock_weights(self, stock_data):
        """Fill the Stock Weights sheet"""
        if '2. Stock Weights' not in self.sheets:
            logger.warning("Stock Weights sheet not found")
            return
        
        sheet = self.sheets['2. Stock Weights']
        
        # This sheet typically contains quarter-wise weights calculation
        # Fill based on the quarterly data of selected stocks
        
        # Get all unique quarters from the stock data
        all_quarters = set()
        for stock_info in stock_data:
            for quarter_data in stock_info['quarterly_data']:
                quarter_key = f"Q{quarter_data.quarter_number}-{quarter_data.quarter_year}"
                all_quarters.add((quarter_data.quarter_year, quarter_data.quarter_number, quarter_key))
        
        # Sort quarters by year and quarter number (most recent first)
        sorted_quarters = sorted(all_quarters, key=lambda x: (x[0], x[1]), reverse=True)
        
        # Fill quarter headers (starting from column B)
        for idx, (year, quarter, quarter_key) in enumerate(sorted_quarters[:10]):  # Limit to 10 quarters
            col_idx = idx + 2  # Start from column B (index 2)
            if col_idx <= 26:  # Limit to Z column
                col_letter = chr(ord('A') + col_idx - 1)
                self._safe_set_cell_value(sheet, f'{col_letter}1', quarter_key)
        
        # Fill stock weights for each quarter
        start_row = 2
        for stock_idx, stock_info in enumerate(stock_data):
            row = start_row + stock_idx
            stock = stock_info['stock']
            
            self._safe_set_cell_value(sheet, f'A{row}', stock.name)  # Stock name in first column
            
            # Fill weights for each quarter
            for quarter_idx, (year, quarter, quarter_key) in enumerate(sorted_quarters[:10]):
                col_idx = quarter_idx + 2
                if col_idx <= 26:
                    col_letter = chr(ord('A') + col_idx - 1)
                    
                    # Find the quarter data for this stock
                    quarter_data = None
                    for qd in stock_info['quarterly_data']:
                        if qd.quarter_year == year and qd.quarter_number == quarter:
                            quarter_data = qd
                            break
                    
                    if quarter_data and quarter_data.mcap:
                        # Calculate weight as percentage of total market cap
                        weight = stock_info['weightage']  # Use entered weightage
                        self._safe_set_cell_value(sheet, f'{col_letter}{row}', weight)
                    else:
                        self._safe_set_cell_value(sheet, f'{col_letter}{row}', 0)
    
    def _fill_summary(self, fund_name, stock_data):
        """Fill the Summary sheet"""
        if '1. Summary' not in self.sheets:
            logger.warning("Summary sheet not found")
            return
        
        sheet = self.sheets['1. Summary']
        
        # Add fund name
        self._safe_set_cell_value(sheet, 'B1', fund_name)
        self._safe_set_cell_value(sheet, 'B2', f"Generated on {datetime.now().strftime('%Y-%m-%d')}")
        
        # Calculate summary statistics
        total_stocks = len(stock_data)
        total_weightage = sum(stock_info['weightage'] for stock_info in stock_data)
        
        # Calculate average metrics
        if stock_data:
            avg_pe = sum(stock_info['metrics']['current_pe'] for stock_info in stock_data) / total_stocks
            avg_pb = sum(stock_info['metrics']['current_pb'] for stock_info in stock_data) / total_stocks
            avg_mcap = sum(stock_info['metrics']['current_mcap'] for stock_info in stock_data) / total_stocks
            avg_cagr = sum(stock_info['metrics']['cagr_6yr'] for stock_info in stock_data) / total_stocks
            
            # Fill summary values (adjust cell references based on template)
            self._safe_set_cell_value(sheet, 'B5', total_stocks)  # Total number of stocks
            self._safe_set_cell_value(sheet, 'B6', total_weightage)  # Total weightage
            self._safe_set_cell_value(sheet, 'B7', round(avg_pe, 2))  # Average PE
            self._safe_set_cell_value(sheet, 'B8', round(avg_pb, 2))  # Average PB
            self._safe_set_cell_value(sheet, 'B9', round(avg_mcap, 2))  # Average Market Cap
            self._safe_set_cell_value(sheet, 'B10', round(avg_cagr, 2))  # Average CAGR
    
    def _save_report(self, fund_name):
        """Save the generated report"""
        # Create filename
        safe_fund_name = "".join(c for c in fund_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_fund_name = safe_fund_name.replace(' ', '_')
        current_date = datetime.now().strftime('%Y%m%d')
        filename = f"{safe_fund_name}_Analysis_file_{current_date}.xlsm"  # Keep .xlsm extension to preserve macros
        
        # Create temporary file
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, filename)
        
        # Save the workbook with proper format
        try:
            self.workbook.save(output_path)
            logger.info(f"Excel file saved successfully: {output_path}")
        except Exception as e:
            logger.error(f"Error saving Excel file: {str(e)}")
            raise
        
        return output_path