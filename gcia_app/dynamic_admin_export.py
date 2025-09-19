"""
Scalable Block-Based Stock Export Generator
==========================================

This module provides scalable block-based header and column generation for admin stock exports.
It uses a hybrid template-dynamic approach that handles any number of periods (3 to 100+ quarters).

Key Features:
- Fixed template structure with standardized block order
- Dynamic period expansion within each block
- Multi-upload compatible (accumulates periods across uploads)
- Scalable to any number of quarters/periods
- Consistent column positioning regardless of period count
"""

from django.db import models
from .models import StockMarketCap, StockTTMData, StockQuarterlyData, StockAnnualRatios, StockPrice
import logging

logger = logging.getLogger(__name__)

class BlockBasedExportGenerator:
    """
    Generates scalable block-based headers and column mappings for admin stock exports.
    Uses template structure with dynamic period expansion for unlimited scalability.
    """

    def __init__(self):
        self.template_structure = self._define_template_structure()
        self.detected_periods = {}
        self.block_mapping = {}
        self.total_columns = 0

    def _define_template_structure(self):
        """
        Define the fixed template structure with standardized block order.
        This ensures consistent positioning regardless of period count.
        """
        return {
            'basic_info': {
                'name': 'Basic Info',
                'columns': ['S. No.', 'Company Name', 'Accord Code', 'Sector', 'Cap',
                           'Free Float', '6 Year CAGR', 'TTM', '6 Year CAGR', 'TTM',
                           'Current', '2 Yr Avg', 'Reval/deval'],
                'fixed_size': 13
            },
            'market_cap': {
                'name': 'Market Cap',
                'field_name': 'market_cap',
                'model': 'StockMarketCap',
                'date_field': 'date',
                'description': 'Market Cap (in crores)',
                'row3_title': 'Market Cap'
            },
            'market_cap_ff': {
                'name': 'Market Cap Free Float',
                'field_name': 'market_cap_free_float',
                'model': 'StockMarketCap',
                'date_field': 'date',
                'description': 'Market Cap- Free Float (in crores)',
                'row3_title': ''
            },
            'ttm_revenue': {
                'name': 'TTM Revenue',
                'field_name': 'ttm_revenue',
                'model': 'StockTTMData',
                'date_field': 'period',
                'description': 'TTM Revenue',
                'row3_title': 'Net Sales'
            },
            'ttm_revenue_ff': {
                'name': 'TTM Revenue Free Float',
                'field_name': 'ttm_revenue_free_float',
                'model': 'StockTTMData',
                'date_field': 'period',
                'description': 'TTM Revenue- Free Float',
                'row3_title': ''
            },
            'ttm_pat': {
                'name': 'TTM PAT',
                'field_name': 'ttm_pat',
                'model': 'StockTTMData',
                'date_field': 'period',
                'description': 'TTM PAT',
                'row3_title': 'Profit After Tax'
            },
            'ttm_pat_ff': {
                'name': 'TTM PAT Free Float',
                'field_name': 'ttm_pat_free_float',
                'model': 'StockTTMData',
                'date_field': 'period',
                'description': 'TTM PAT- Free Float',
                'row3_title': ''
            },
            'quarterly_revenue': {
                'name': 'Quarterly Revenue',
                'field_name': 'quarterly_revenue',
                'model': 'StockQuarterlyData',
                'date_field': 'period',
                'description': 'Quarterly- Revenue',
                'row3_title': 'Net Sales & Other Operating Income'
            },
            'quarterly_revenue_ff': {
                'name': 'Quarterly Revenue Free Float',
                'field_name': 'quarterly_revenue_free_float',
                'model': 'StockQuarterlyData',
                'date_field': 'period',
                'description': 'Quarterly- Revenue- Free Float',
                'row3_title': ''
            },
            'quarterly_pat': {
                'name': 'Quarterly PAT',
                'field_name': 'quarterly_pat',
                'model': 'StockQuarterlyData',
                'date_field': 'period',
                'description': 'Quarterly- PAT',
                'row3_title': 'Profit after tax'
            },
            'quarterly_pat_ff': {
                'name': 'Quarterly PAT Free Float',
                'field_name': 'quarterly_pat_free_float',
                'model': 'StockQuarterlyData',
                'date_field': 'period',
                'description': 'Quarterly-PAT- Free Float',
                'row3_title': ''
            },
            'roce': {
                'name': 'ROCE',
                'field_name': 'roce_percentage',
                'model': 'StockAnnualRatios',
                'date_field': 'financial_year',
                'description': 'ROCE (%)',
                'row3_title': 'ROCE (%)'
            },
            'roe': {
                'name': 'ROE',
                'field_name': 'roe_percentage',
                'model': 'StockAnnualRatios',
                'date_field': 'financial_year',
                'description': 'ROE (%)',
                'row3_title': 'ROE (%)'
            },
            'retention': {
                'name': 'Retention',
                'field_name': 'retention_percentage',
                'model': 'StockAnnualRatios',
                'date_field': 'financial_year',
                'description': 'Retention (%)',
                'row3_title': 'Retention (%)'
            },
            'share_price': {
                'name': 'Share Price',
                'field_name': 'share_price',
                'model': 'StockPrice',
                'date_field': 'price_date',
                'description': 'Share Price',
                'row3_title': 'Share Price'
            },
            'pr_ratio': {
                'name': 'PR Ratio',
                'field_name': 'pr_ratio',
                'model': 'StockPrice',
                'date_field': 'price_date',
                'description': 'Price to Revenue Ratio',
                'row3_title': 'PR'
            },
            'pe_ratio': {
                'name': 'PE Ratio',
                'field_name': 'pe_ratio',
                'model': 'StockPrice',
                'date_field': 'price_date',
                'description': 'Price to Earnings Ratio',
                'row3_title': 'PE'
            },
            'identifiers': {
                'name': 'Identifiers',
                'columns': ['BSE Code', 'NSE Code', 'ISIN'],
                'fixed_size': 3
            }
        }

    def collect_all_periods_from_database(self):
        """
        Collect all unique periods from database for each data type.
        This provides the foundation for dynamic block sizing.

        Returns:
            dict: All unique periods/dates for each data type, sorted chronologically
        """
        try:
            periods = {}

            # Market Cap dates (shared by market_cap and market_cap_ff blocks)
            market_cap_dates = list(
                StockMarketCap.objects.values_list('date', flat=True)
                .distinct()
                .order_by('-date')  # Most recent first
            )
            periods['market_cap_dates'] = [date.strftime('%Y-%m-%d') for date in market_cap_dates]

            # TTM periods (shared by all TTM blocks)
            ttm_periods = list(
                StockTTMData.objects.values_list('period', flat=True)
                .distinct()
                .order_by('-period')  # Most recent first
            )
            periods['ttm_periods'] = ttm_periods

            # Quarterly periods (shared by all quarterly blocks)
            quarterly_periods = list(
                StockQuarterlyData.objects.values_list('period', flat=True)
                .distinct()
                .order_by('-period')  # Most recent first
            )
            periods['quarterly_periods'] = quarterly_periods

            # Annual years (shared by ROCE, ROE, Retention blocks)
            annual_years = list(
                StockAnnualRatios.objects.values_list('financial_year', flat=True)
                .distinct()
                .order_by('-financial_year')  # Most recent first
            )
            periods['annual_years'] = annual_years

            # Price dates (shared by share_price, pr_ratio, pe_ratio blocks)
            price_dates = list(
                StockPrice.objects.values_list('price_date', flat=True)
                .distinct()
                .order_by('-price_date')  # Most recent first
            )
            periods['price_dates'] = [date.strftime('%Y-%m-%d') for date in price_dates]

            logger.info(f"BLOCK-BASED period collection:")
            logger.info(f"  - Market cap dates: {len(periods['market_cap_dates'])}")
            logger.info(f"  - TTM periods: {len(periods['ttm_periods'])}")
            logger.info(f"  - Quarterly periods: {len(periods['quarterly_periods'])}")
            logger.info(f"  - Annual years: {len(periods['annual_years'])}")
            logger.info(f"  - Price dates: {len(periods['price_dates'])}")

            self.detected_periods = periods
            return periods

        except Exception as e:
            logger.error(f"Error collecting periods from database: {str(e)}")
            return {
                'market_cap_dates': [],
                'ttm_periods': [],
                'quarterly_periods': [],
                'annual_years': [],
                'price_dates': []
            }

    def calculate_block_sizes_and_positions(self, periods):
        """
        Calculate the size of each block and its column positions.
        Uses contiguous column ranges for scalable positioning.

        Args:
            periods: Dictionary of all periods for each data type

        Returns:
            dict: Block mapping with start/end positions and periods
        """
        try:
            block_order = [
                'basic_info',
                'market_cap', 'market_cap_ff',
                'ttm_revenue', 'ttm_revenue_ff', 'ttm_pat', 'ttm_pat_ff',
                'quarterly_revenue', 'quarterly_revenue_ff', 'quarterly_pat', 'quarterly_pat_ff',
                'roce', 'roe', 'retention',
                'share_price', 'pr_ratio', 'pe_ratio',
                'identifiers'
            ]

            mapping = {}
            current_col = 0
            separator_cols = 1  # Add 1 separator column between each block

            for block_key in block_order:
                block_def = self.template_structure[block_key]

                if block_key in ['basic_info', 'identifiers']:
                    # Fixed-size blocks
                    block_size = block_def['fixed_size']
                    periods_list = block_def['columns']
                else:
                    # Dynamic blocks - determine size based on periods
                    if block_key in ['market_cap', 'market_cap_ff']:
                        periods_list = periods['market_cap_dates']
                    elif block_key.startswith('ttm_'):
                        periods_list = periods['ttm_periods']
                    elif block_key.startswith('quarterly_'):
                        periods_list = periods['quarterly_periods']
                    elif block_key in ['roce', 'roe', 'retention']:
                        periods_list = periods['annual_years']
                    elif block_key in ['share_price', 'pr_ratio', 'pe_ratio']:
                        periods_list = periods['price_dates']
                    else:
                        periods_list = []

                    block_size = len(periods_list)

                # Store block mapping
                mapping[block_key] = {
                    'start_col': current_col,
                    'end_col': current_col + block_size - 1 if block_size > 0 else current_col,
                    'size': block_size,
                    'periods': periods_list,
                    'definition': block_def
                }

                current_col += block_size

                # Add separator except after last block
                if block_key != 'identifiers':
                    current_col += separator_cols

            # Calculate total columns
            # current_col now points to the position after the last block
            # Since identifiers is the last block and has no separator after it,
            # total_cols should be current_col (which equals the number of columns needed)
            total_cols = current_col

            logger.info(f"BLOCK-BASED mapping calculated:")
            logger.info(f"  - Total blocks: {len(mapping)}")
            logger.info(f"  - Total columns: {total_cols}")
            for block_key, block_info in mapping.items():
                logger.info(f"  - {block_key}: cols {block_info['start_col']}-{block_info['end_col']} (size: {block_info['size']})")

            self.block_mapping = mapping
            self.total_columns = total_cols
            return mapping

        except Exception as e:
            logger.error(f"Error calculating block sizes: {str(e)}")
            raise

    def generate_block_based_headers(self, block_mapping, total_columns):
        """
        Generate 8-row header structure using block-based positioning.
        Each block gets its periods placed in contiguous column ranges.

        Args:
            block_mapping: Dictionary of block positions and periods
            total_columns: Total number of columns

        Returns:
            dict: Complete 8-row header structure
        """
        try:
            # Initialize all rows with empty cells
            header_structure = {}
            for row in range(1, 9):
                header_structure[f'row_{row}'] = [''] * total_columns

            # Row 2: Column numbers
            for i in range(1, total_columns):
                header_structure['row_2'][i] = i

            # Row 3: Field descriptions (positioned at start of relevant blocks)
            for block_key, block_info in block_mapping.items():
                if block_key in ['basic_info', 'identifiers']:
                    # Basic info and identifiers
                    for i, header in enumerate(block_info['periods']):
                        col_pos = block_info['start_col'] + i
                        if col_pos < total_columns:
                            header_structure['row_8'][col_pos] = header
                else:
                    # Dynamic blocks - add row 3 title at start of block
                    row3_title = block_info['definition'].get('row3_title', '')
                    if row3_title and block_info['start_col'] < total_columns:
                        header_structure['row_3'][block_info['start_col']] = row3_title

            # Row 6: Category descriptions (span entire blocks)
            for block_key, block_info in block_mapping.items():
                if block_key not in ['basic_info', 'identifiers']:
                    description = block_info['definition'].get('description', '')
                    if description:
                        for col in range(block_info['start_col'], min(block_info['end_col'] + 1, total_columns)):
                            header_structure['row_6'][col] = description

            # Row 8: Actual column headers (periods/dates/basic columns)
            for block_key, block_info in block_mapping.items():
                for i, period in enumerate(block_info['periods']):
                    col_pos = block_info['start_col'] + i
                    if col_pos < total_columns:
                        header_structure['row_8'][col_pos] = period

            logger.info(f"BLOCK-BASED headers generated for {total_columns} columns")
            return header_structure

        except Exception as e:
            logger.error(f"Error generating block-based headers: {str(e)}")
            raise

    def get_stock_data_by_blocks(self, stock, block_mapping):
        """
        Retrieve stock data organized by blocks for efficient row population.

        Args:
            stock: Stock instance
            block_mapping: Dictionary of block positions and periods

        Returns:
            dict: Stock data organized by block and period
        """
        try:
            stock_data = {}

            # Market Cap data
            market_cap_data = {}
            for mc in stock.market_cap_data.all():
                date_str = mc.date.strftime('%Y-%m-%d')
                market_cap_data[date_str] = {
                    'market_cap': mc.market_cap,
                    'market_cap_free_float': mc.market_cap_free_float
                }
            stock_data['market_cap'] = market_cap_data

            # TTM data
            ttm_data = {}
            for ttm in stock.ttm_data.all():
                ttm_data[ttm.period] = {
                    'ttm_revenue': ttm.ttm_revenue,
                    'ttm_revenue_free_float': ttm.ttm_revenue_free_float,
                    'ttm_pat': ttm.ttm_pat,
                    'ttm_pat_free_float': ttm.ttm_pat_free_float
                }
            stock_data['ttm'] = ttm_data

            # Quarterly data
            quarterly_data = {}
            for qtr in stock.quarterly_data.all():
                quarterly_data[qtr.period] = {
                    'quarterly_revenue': qtr.quarterly_revenue,
                    'quarterly_revenue_free_float': qtr.quarterly_revenue_free_float,
                    'quarterly_pat': qtr.quarterly_pat,
                    'quarterly_pat_free_float': qtr.quarterly_pat_free_float
                }
            stock_data['quarterly'] = quarterly_data

            # Annual data
            annual_data = {}
            for ratio in stock.annual_ratios.all():
                annual_data[ratio.financial_year] = {
                    'roce_percentage': ratio.roce_percentage,
                    'roe_percentage': ratio.roe_percentage,
                    'retention_percentage': ratio.retention_percentage
                }
            stock_data['annual'] = annual_data

            # Price data
            price_data = {}
            for price in stock.price_data.all():
                date_str = price.price_date.strftime('%Y-%m-%d')
                price_data[date_str] = {
                    'share_price': price.share_price,
                    'pr_ratio': price.pr_ratio,
                    'pe_ratio': price.pe_ratio
                }
            stock_data['price'] = price_data

            return stock_data

        except Exception as e:
            logger.error(f"Error getting stock data by blocks: {str(e)}")
            return {}

    def populate_stock_row_by_blocks(self, stock, stock_data, block_mapping, row_index):
        """
        Populate a stock row using block-based positioning.
        Each block gets its data placed in the correct contiguous column range.

        Args:
            stock: Stock instance
            stock_data: Stock data organized by blocks
            block_mapping: Dictionary of block positions and periods
            row_index: Row number for basic info

        Returns:
            list: Complete row data with correct positioning
        """
        try:
            row_data = [''] * self.total_columns

            # Basic info block
            basic_block = block_mapping['basic_info']
            row_data[0] = row_index  # S.No
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

            # Market Cap blocks
            self._populate_block_data(row_data, block_mapping['market_cap'],
                                    stock_data.get('market_cap', {}), 'market_cap')
            self._populate_block_data(row_data, block_mapping['market_cap_ff'],
                                    stock_data.get('market_cap', {}), 'market_cap_free_float')

            # TTM blocks
            self._populate_block_data(row_data, block_mapping['ttm_revenue'],
                                    stock_data.get('ttm', {}), 'ttm_revenue')
            self._populate_block_data(row_data, block_mapping['ttm_revenue_ff'],
                                    stock_data.get('ttm', {}), 'ttm_revenue_free_float')
            self._populate_block_data(row_data, block_mapping['ttm_pat'],
                                    stock_data.get('ttm', {}), 'ttm_pat')
            self._populate_block_data(row_data, block_mapping['ttm_pat_ff'],
                                    stock_data.get('ttm', {}), 'ttm_pat_free_float')

            # Quarterly blocks
            self._populate_block_data(row_data, block_mapping['quarterly_revenue'],
                                    stock_data.get('quarterly', {}), 'quarterly_revenue')
            self._populate_block_data(row_data, block_mapping['quarterly_revenue_ff'],
                                    stock_data.get('quarterly', {}), 'quarterly_revenue_free_float')
            self._populate_block_data(row_data, block_mapping['quarterly_pat'],
                                    stock_data.get('quarterly', {}), 'quarterly_pat')
            self._populate_block_data(row_data, block_mapping['quarterly_pat_ff'],
                                    stock_data.get('quarterly', {}), 'quarterly_pat_free_float')

            # Annual blocks
            self._populate_block_data(row_data, block_mapping['roce'],
                                    stock_data.get('annual', {}), 'roce_percentage')
            self._populate_block_data(row_data, block_mapping['roe'],
                                    stock_data.get('annual', {}), 'roe_percentage')
            self._populate_block_data(row_data, block_mapping['retention'],
                                    stock_data.get('annual', {}), 'retention_percentage')

            # Price blocks
            self._populate_block_data(row_data, block_mapping['share_price'],
                                    stock_data.get('price', {}), 'share_price')
            self._populate_block_data(row_data, block_mapping['pr_ratio'],
                                    stock_data.get('price', {}), 'pr_ratio')
            self._populate_block_data(row_data, block_mapping['pe_ratio'],
                                    stock_data.get('price', {}), 'pe_ratio')

            # Identifiers block
            identifiers_block = block_mapping['identifiers']
            row_data[identifiers_block['start_col']] = stock.bse_code
            row_data[identifiers_block['start_col'] + 1] = stock.nse_symbol
            row_data[identifiers_block['start_col'] + 2] = stock.isin

            return row_data

        except Exception as e:
            logger.error(f"Error populating stock row by blocks: {str(e)}")
            return [''] * self.total_columns

    def _populate_block_data(self, row_data, block_info, data_dict, field_name):
        """
        Helper method to populate data for a specific block.

        Args:
            row_data: Row data array to populate
            block_info: Block information with positions and periods
            data_dict: Dictionary of data keyed by period
            field_name: Name of the field to extract from data
        """
        try:
            for i, period in enumerate(block_info['periods']):
                col_pos = block_info['start_col'] + i
                if col_pos < len(row_data):
                    period_data = data_dict.get(period, {})
                    value = period_data.get(field_name)
                    row_data[col_pos] = value if value is not None else ''
        except Exception as e:
            logger.error(f"Error populating block data for {field_name}: {str(e)}")

    def get_complete_block_based_export_structure(self):
        """
        Get complete export structure using block-based approach.
        This is the main method that orchestrates the entire export generation.

        Returns:
            dict: Complete export structure with headers, mapping, and data population methods
        """
        try:
            # Step 1: Collect all periods from database
            periods = self.collect_all_periods_from_database()

            # Step 2: Calculate block sizes and positions
            block_mapping = self.calculate_block_sizes_and_positions(periods)

            # Step 3: Generate headers using block-based positioning
            header_structure = self.generate_block_based_headers(block_mapping, self.total_columns)

            return {
                'periods': periods,
                'block_mapping': block_mapping,
                'total_columns': self.total_columns,
                'header_structure': header_structure,
                'is_block_based': True,
                'generator_instance': self
            }

        except Exception as e:
            logger.error(f"Error generating complete block-based export structure: {str(e)}")
            raise

    # Deprecated methods for backward compatibility
    def get_complete_ultra_dynamic_export_structure(self):
        """DEPRECATED: Use get_complete_block_based_export_structure() instead."""
        logger.warning("Using deprecated ultra-dynamic method. Switching to block-based approach.")
        return self.get_complete_block_based_export_structure()

    def get_complete_dynamic_export_structure(self):
        """DEPRECATED: Use get_complete_block_based_export_structure() instead."""
        logger.warning("Using deprecated dynamic method. Switching to block-based approach.")
        return self.get_complete_block_based_export_structure()


# For backward compatibility
DynamicAdminExportGenerator = BlockBasedExportGenerator