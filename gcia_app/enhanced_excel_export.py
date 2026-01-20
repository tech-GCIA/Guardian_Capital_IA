"""
Enhanced Excel Export with Portfolio Analysis Metrics
=====================================================

This module provides the enhanced Excel export functionality that includes
calculated portfolio analysis metrics in the Portfolio Analysis format.
Now aligned with Stock Base Sheet export structure for consistency.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from django.utils import timezone
from io import BytesIO
from .metrics_calculator import DynamicHeaderGenerator
from .models import FundMetricsLog, AMCFundScheme, FundHolding, PortfolioMetricsLog
from .dynamic_admin_export import BlockBasedExportGenerator
from datetime import datetime
import logging

# Import calculation functions for portfolio metrics
from .excel_calc_functions import (
    calculate_patm_from_totals, calculate_qoq_from_totals,
    calculate_yoy_from_totals, calculate_6yr_cagr_from_totals,
    calculate_pe_pr_from_totals, calculate_pe_pr_averages_from_totals,
    calculate_reval_deval_from_totals, calculate_pr_10q_extremes_from_totals,
    calculate_pe_yield_from_totals, calculate_growth_from_totals, get_bond_rate,
    build_section_column_mapping, create_metric_row
)

logger = logging.getLogger(__name__)


class FundPortfolioExportGenerator(BlockBasedExportGenerator):
    """
    Extends BlockBasedExportGenerator to add fund-specific columns and portfolio metrics.
    Maintains same block-based structure as Stock Base Sheet export for consistency.
    """

    def __init__(self, scheme):
        """Initialize with specific fund scheme"""
        super().__init__()
        self.scheme = scheme
        self.section_start_columns = {}  # Track where each data section starts for metric rows

    def _define_block_structure(self, periods):
        """
        Override to add fund-specific data block after basic info.
        Maintains same structure as parent but inserts Fund Data block.
        """
        blocks = [
            {
                'name': 'basic_info',
                'type': 'fixed',
                'size': 18,  # Expanded: removed S.No, added Market Cap, moved fund data here, added Share Price & Free Float
                'columns': [
                    'Company Name', 'Accord Code', 'Sector', 'Cap',
                    'Market Cap', 'Weights', 'Factor', 'Value', 'No.of shares',
                    'Share Price', 'Free Float',
                    '6 Year CAGR', 'TTM', '6 Year CAGR', 'TTM',
                    'Current', '2 Yr Avg', 'Reval/deval'
                ],
                'row6_label': None,  # First 11 columns have no category
                'row7_label': None
            },
            {'name': 'sep_0', 'type': 'separator', 'size': 1},
            {
                'name': 'market_cap',
                'type': 'dynamic',
                'data_type': 'market_cap',
                'periods': periods['market_cap_dates'],
                'row6_label': 'Market Cap (in crores)',
                'row7_label': 'Market Cap (in crores)'
            },
            {'name': 'sep_2', 'type': 'separator', 'size': 1},
            {
                'name': 'market_cap_ff',
                'type': 'dynamic',
                'data_type': 'market_cap_free_float',
                'periods': periods['market_cap_dates'],
                'row6_label': 'Market Cap- Free Float (in crores)',
                'row7_label': 'Market Cap- Free Float (in crores)'
            },
            {'name': 'sep_3', 'type': 'separator', 'size': 1},
            {
                'name': 'ttm_revenue',
                'type': 'dynamic',
                'data_type': 'ttm_revenue',
                'periods': periods['ttm_periods'],
                'row6_label': 'TTM Revenue',
                'row7_label': 'TTM Revenue'
            },
            {'name': 'sep_4', 'type': 'separator', 'size': 1},
            {
                'name': 'ttm_revenue_ff',
                'type': 'dynamic',
                'data_type': 'ttm_revenue_free_float',
                'periods': periods['ttm_periods'],
                'row6_label': 'TTM Revenue- Free Float',
                'row7_label': 'TTM Revenue- Free Float'
            },
            {'name': 'sep_5', 'type': 'separator', 'size': 1},
            {
                'name': 'ttm_pat',
                'type': 'dynamic',
                'data_type': 'ttm_pat',
                'periods': periods['ttm_periods'],
                'row6_label': 'TTM PAT',
                'row7_label': 'TTM PAT'
            },
            {'name': 'sep_6', 'type': 'separator', 'size': 1},
            {
                'name': 'ttm_pat_ff',
                'type': 'dynamic',
                'data_type': 'ttm_pat_free_float',
                'periods': periods['ttm_periods'],
                'row6_label': 'TTM PAT- Free Float',
                'row7_label': 'TTM PAT- Free Float'
            },
            {'name': 'sep_7', 'type': 'separator', 'size': 1},
            {
                'name': 'quarterly_revenue',
                'type': 'dynamic',
                'data_type': 'quarterly_revenue',
                'periods': periods['quarterly_periods'],
                'row6_label': 'Quarterly- Revenue',
                'row7_label': 'Quarterly- Revenue'
            },
            {'name': 'sep_8', 'type': 'separator', 'size': 1},
            {
                'name': 'quarterly_revenue_ff',
                'type': 'dynamic',
                'data_type': 'quarterly_revenue_free_float',
                'periods': periods['quarterly_periods'],
                'row6_label': 'Quarterly- Revenue- Free Float',
                'row7_label': 'Quarterly- Revenue- Free Float'
            },
            {'name': 'sep_9', 'type': 'separator', 'size': 1},
            {
                'name': 'quarterly_pat',
                'type': 'dynamic',
                'data_type': 'quarterly_pat',
                'periods': periods['quarterly_periods'],
                'row6_label': 'Quarterly- PAT',
                'row7_label': 'Quarterly- PAT'
            },
            {'name': 'sep_10', 'type': 'separator', 'size': 1},
            {
                'name': 'quarterly_pat_ff',
                'type': 'dynamic',
                'data_type': 'quarterly_pat_free_float',
                'periods': periods['quarterly_periods'],
                'row6_label': 'Quarterly- PAT- Free Float',
                'row7_label': 'Quarterly- PAT- Free Float'
            },
            {'name': 'sep_11', 'type': 'separator', 'size': 1},
            {
                'name': 'roce',
                'type': 'dynamic',
                'data_type': 'roce',
                'periods': periods['annual_years'],
                'row6_label': 'ROCE (%)',
                'row7_label': 'ROCE (%)'
            },
            {'name': 'sep_12', 'type': 'separator', 'size': 1},
            {
                'name': 'roe',
                'type': 'dynamic',
                'data_type': 'roe',
                'periods': periods['annual_years'],
                'row6_label': 'ROE (%)',
                'row7_label': 'ROE (%)'
            },
            {'name': 'sep_13', 'type': 'separator', 'size': 1},
            {
                'name': 'retention',
                'type': 'dynamic',
                'data_type': 'retention',
                'periods': periods['annual_years'],
                'row6_label': 'Retention (%)',
                'row7_label': 'Retention (%)'
            },
            {'name': 'sep_14', 'type': 'separator', 'size': 1},
            {
                'name': 'share_price',
                'type': 'dynamic',
                'data_type': 'share_price',
                'periods': periods['share_price_dates'],
                'row6_label': 'Share Price',
                'row7_label': 'Share Price'
            },
            {'name': 'sep_15', 'type': 'separator', 'size': 1},
            {
                'name': 'pr_ratio',
                'type': 'dynamic',
                'data_type': 'pr_ratio',
                'periods': periods['pr_dates'],
                'row6_label': 'PR',
                'row7_label': 'PR'
            },
            {'name': 'sep_16', 'type': 'separator', 'size': 1},
            {
                'name': 'pe_ratio',
                'type': 'dynamic',
                'data_type': 'pe_ratio',
                'periods': periods['pe_dates'],
                'row6_label': 'PE',
                'row7_label': 'PE'
            },
            {'name': 'sep_17', 'type': 'separator', 'size': 1},
            {
                'name': 'identifiers',
                'type': 'fixed',
                'size': 3,
                'columns': ['BSE Code', 'NSE Code', 'ISIN']
            }
        ]

        return blocks

    def collect_all_periods_for_fund(self):
        """
        Collect all available periods from holdings' stocks in this fund.
        Similar to parent's collect_all_periods_from_database but filtered to fund stocks.
        """
        from .models import StockMarketCap, StockTTMData, StockQuarterlyData, StockAnnualRatios, StockPrice

        # Get all stocks in this fund
        holdings = FundHolding.objects.filter(scheme=self.scheme).select_related('stock')
        stock_ids = [h.stock.stock_id for h in holdings]

        # Collect periods from these stocks only
        market_cap_dates = set(StockMarketCap.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('date', flat=True).distinct())

        ttm_periods = set(StockTTMData.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('period', flat=True).distinct())

        quarterly_periods = set(StockQuarterlyData.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('period', flat=True).distinct())

        annual_years = set(StockAnnualRatios.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('financial_year', flat=True).distinct())

        share_price_dates = set(StockPrice.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('price_date', flat=True).distinct())

        # PR and PE dates from StockPrice
        pr_dates = set(StockPrice.objects.filter(
            stock__stock_id__in=stock_ids,
            pr_ratio__isnull=False
        ).values_list('price_date', flat=True).distinct())

        pe_dates = set(StockPrice.objects.filter(
            stock__stock_id__in=stock_ids,
            pe_ratio__isnull=False
        ).values_list('price_date', flat=True).distinct())

        # Sort all periods
        return {
            'market_cap_dates': sorted(market_cap_dates, reverse=True),
            'ttm_periods': sorted(ttm_periods, reverse=True),
            'quarterly_periods': sorted(quarterly_periods, reverse=True),
            'annual_years': sorted(annual_years, reverse=True),
            'share_price_dates': sorted(share_price_dates, reverse=True),
            'pr_dates': sorted(pr_dates, reverse=True),
            'pe_dates': sorted(pe_dates, reverse=True)
        }

    def collect_periods_for_filtered_holdings(self, filtered_holdings):
        """
        Collect all available periods from filtered holdings' stocks.
        Similar to collect_all_periods_for_fund but uses pre-filtered holdings QuerySet.

        Args:
            filtered_holdings: QuerySet of FundHolding objects (already filtered)

        Returns:
            Dict with sorted period lists for each data type
        """
        from .models import StockMarketCap, StockTTMData, StockQuarterlyData, StockAnnualRatios, StockPrice

        # Get stock IDs from filtered holdings
        stock_ids = [h.stock.stock_id for h in filtered_holdings]

        if not stock_ids:
            # Return empty periods if no stocks
            logger.warning("No stocks in filtered holdings - returning empty periods")
            return {
                'market_cap_dates': [],
                'ttm_periods': [],
                'quarterly_periods': [],
                'annual_years': [],
                'share_price_dates': [],
                'pr_dates': [],
                'pe_dates': []
            }

        # Collect periods from filtered stocks only
        market_cap_dates = set(StockMarketCap.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('date', flat=True).distinct())

        ttm_periods = set(StockTTMData.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('period', flat=True).distinct())

        quarterly_periods = set(StockQuarterlyData.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('period', flat=True).distinct())

        annual_years = set(StockAnnualRatios.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('financial_year', flat=True).distinct())

        share_price_dates = set(StockPrice.objects.filter(
            stock__stock_id__in=stock_ids
        ).values_list('price_date', flat=True).distinct())

        # PR and PE dates from StockPrice
        pr_dates = set(StockPrice.objects.filter(
            stock__stock_id__in=stock_ids,
            pr_ratio__isnull=False
        ).values_list('price_date', flat=True).distinct())

        pe_dates = set(StockPrice.objects.filter(
            stock__stock_id__in=stock_ids,
            pe_ratio__isnull=False
        ).values_list('price_date', flat=True).distinct())

        logger.info(f"Collected periods from {len(stock_ids)} filtered stocks")

        # Sort all periods
        return {
            'market_cap_dates': sorted(market_cap_dates, reverse=True),
            'ttm_periods': sorted(ttm_periods, reverse=True),
            'quarterly_periods': sorted(quarterly_periods, reverse=True),
            'annual_years': sorted(annual_years, reverse=True),
            'share_price_dates': sorted(share_price_dates, reverse=True),
            'pr_dates': sorted(pr_dates, reverse=True),
            'pe_dates': sorted(pe_dates, reverse=True)
        }

    def _generate_import_style_headers(self, blocks, total_columns):
        """
        Override parent to correctly position fundamental headers for Fund Analysis.
        Fund has 18 columns in basic_info, so fundamentals are at columns 11-17 (not 6-12).
        """
        # Call parent to get base headers
        headers = super()._generate_import_style_headers(blocks, total_columns)

        # Clear parent's incorrect fundamental headers at columns 6-10
        # Parent class hardcodes "Stock Fundamentals" at 6-12 for Stock Base Sheet
        # But Fund Analysis has different structure: fundamentals are at 11-17
        for col_idx in range(6, 11):
            if col_idx < total_columns:
                headers['row_6'][col_idx] = ''  # Clear incorrect section header
                headers['row_7'][col_idx] = ''  # Clear incorrect subcategory

        # Override the hardcoded fundamental headers from parent
        # Fund Analysis: Columns 0-10 are Company, Accord, Sector, Cap, Market Cap,
        #                Weights, Factor, Value, Shares, Price, Free Float (NO fundamentals header)
        # Fund Analysis: Columns 11-17 are Fundamentals (Revenue/PAT CAGR/TTM, PR ratios)
        fundamental_start_col = 11

        # Row 6: Section header spanning fundamental columns (11-17)
        fundamentals_section = "Stock wise Fundamentals and Valuations"
        for col_idx in range(fundamental_start_col, fundamental_start_col + 7):
            if col_idx < total_columns:
                headers['row_6'][col_idx] = fundamentals_section

        # Row 7: Subcategory labels
        if total_columns > 11:
            headers['row_7'][11] = "Revenue"   # 6 Year CAGR
        if total_columns > 12:
            headers['row_7'][12] = "Revenue"   # TTM
        if total_columns > 13:
            headers['row_7'][13] = "PAT"       # 6 Year CAGR
        if total_columns > 14:
            headers['row_7'][14] = "PAT"       # TTM
        if total_columns > 15:
            headers['row_7'][15] = "PR"        # Current
        if total_columns > 16:
            headers['row_7'][16] = "PR"        # 2 Yr Avg
        if total_columns > 17:
            headers['row_7'][17] = "PR"        # Reval/Deval

        return headers

    def populate_fund_stock_row(self, holding, blocks, total_columns):
        """
        Populate a single stock row with fund-specific data.
        Extends parent's populate_stock_row_header_driven with fund data.
        """
        stock = holding.stock
        row_data = [''] * total_columns
        current_col = 0

        # Calculate total portfolio value for factor calculation
        holdings = FundHolding.objects.filter(scheme=self.scheme)
        total_portfolio_value = sum(h.market_value or 0 for h in holdings if h.market_value)

        for block in blocks:
            if block['type'] == 'separator':
                # Empty column
                row_data[current_col] = ''
                current_col += 1

            elif block['type'] == 'fixed':
                if block['name'] == 'basic_info':
                    # 18 columns: Company info + Fund Data + Share Price + Free Float + Fundamentals
                    from .models import StockMarketCap, StockPrice, FundMetricsLog

                    # Col 1: Company Name (NO S.No!)
                    row_data[current_col] = stock.company_name

                    # Col 2: Accord Code
                    row_data[current_col + 1] = stock.accord_code or ''

                    # Col 3: Sector
                    row_data[current_col + 2] = stock.sector or 'Unknown'

                    # Col 4: Cap
                    row_data[current_col + 3] = stock.cap or 'Unknown'

                    # Col 5: Market Cap (latest)
                    market_cap_data = StockMarketCap.objects.filter(stock=stock).order_by('-date').first()
                    row_data[current_col + 4] = market_cap_data.market_cap if market_cap_data else None

                    # Col 6-9: Fund Data
                    weights = (holding.holding_percentage / 100) if holding.holding_percentage else 0
                    factor = (holding.market_value / total_portfolio_value) if holding.market_value and total_portfolio_value else 0
                    row_data[current_col + 5] = weights
                    row_data[current_col + 6] = factor
                    row_data[current_col + 7] = holding.market_value or 0
                    row_data[current_col + 8] = holding.number_of_shares or 0

                    # Col 10: Share Price
                    price_data = StockPrice.objects.filter(stock=stock).order_by('-price_date').first()
                    row_data[current_col + 9] = price_data.share_price if price_data else None

                    # Col 11: Free Float
                    row_data[current_col + 10] = stock.free_float if stock.free_float else 0

                    # Col 12-18: Fundamentals (get from latest metrics)
                    latest_metrics = FundMetricsLog.objects.filter(
                        scheme=self.scheme,
                        stock=stock
                    ).order_by('-period_date').first()

                    if latest_metrics:
                        row_data[current_col + 11] = latest_metrics.revenue_6yr_cagr
                        row_data[current_col + 12] = None  # TTM Revenue
                        row_data[current_col + 13] = latest_metrics.pat_6yr_cagr
                        row_data[current_col + 14] = None  # TTM PAT
                        row_data[current_col + 15] = latest_metrics.current_pr
                        row_data[current_col + 16] = latest_metrics.pr_2yr_avg
                        row_data[current_col + 17] = latest_metrics.pr_2yr_reval_deval

                    current_col += 18

                elif block['name'] == 'identifiers':
                    # Stock identifiers (3 columns)
                    row_data[current_col] = stock.bse_code or ''
                    row_data[current_col + 1] = stock.nse_symbol or ''
                    row_data[current_col + 2] = stock.isin or ''
                    current_col += 3

            elif block['type'] == 'dynamic':
                # Track section start column for metric rows
                data_type = block['data_type']
                if data_type not in self.section_start_columns:
                    self.section_start_columns[data_type] = current_col

                # Populate dynamic data columns using parent's logic
                periods = block['periods']

                for period in periods:
                    value = self._get_stock_data_value(stock, data_type, period)
                    row_data[current_col] = value if value is not None else ''
                    current_col += 1

        return row_data

    def _get_stock_data_value(self, stock, data_type, period):
        """Get data value for stock at specific period"""
        from .models import StockMarketCap, StockTTMData, StockQuarterlyData, StockAnnualRatios, StockPrice

        try:
            if data_type == 'market_cap':
                obj = StockMarketCap.objects.filter(stock=stock, date=period).first()
                return obj.market_cap if obj else None
            elif data_type == 'market_cap_free_float':
                obj = StockMarketCap.objects.filter(stock=stock, date=period).first()
                return obj.market_cap_free_float if obj else None
            elif data_type == 'ttm_revenue':
                obj = StockTTMData.objects.filter(stock=stock, period=period).first()
                return obj.ttm_revenue if obj else None
            elif data_type == 'ttm_revenue_free_float':
                obj = StockTTMData.objects.filter(stock=stock, period=period).first()
                return obj.ttm_revenue_free_float if obj else None
            elif data_type == 'ttm_pat':
                obj = StockTTMData.objects.filter(stock=stock, period=period).first()
                return obj.ttm_pat if obj else None
            elif data_type == 'ttm_pat_free_float':
                obj = StockTTMData.objects.filter(stock=stock, period=period).first()
                return obj.ttm_pat_free_float if obj else None
            elif data_type == 'quarterly_revenue':
                obj = StockQuarterlyData.objects.filter(stock=stock, period=period).first()
                return obj.quarterly_revenue if obj else None
            elif data_type == 'quarterly_revenue_free_float':
                obj = StockQuarterlyData.objects.filter(stock=stock, period=period).first()
                return obj.quarterly_revenue_free_float if obj else None
            elif data_type == 'quarterly_pat':
                obj = StockQuarterlyData.objects.filter(stock=stock, period=period).first()
                return obj.quarterly_pat if obj else None
            elif data_type == 'quarterly_pat_free_float':
                obj = StockQuarterlyData.objects.filter(stock=stock, period=period).first()
                return obj.quarterly_pat_free_float if obj else None
            elif data_type == 'roce':
                obj = StockAnnualRatios.objects.filter(stock=stock, financial_year=period).first()
                return obj.roce_percentage if obj else None
            elif data_type == 'roe':
                obj = StockAnnualRatios.objects.filter(stock=stock, financial_year=period).first()
                return obj.roe_percentage if obj else None
            elif data_type == 'retention':
                obj = StockAnnualRatios.objects.filter(stock=stock, financial_year=period).first()
                return obj.retention_percentage if obj else None
            elif data_type == 'share_price':
                obj = StockPrice.objects.filter(stock=stock, price_date=period).first()
                return obj.share_price if obj else None
            elif data_type == 'pr_ratio':
                obj = StockPrice.objects.filter(stock=stock, price_date=period).first()
                return obj.pr_ratio if obj else None
            elif data_type == 'pe_ratio':
                obj = StockPrice.objects.filter(stock=stock, price_date=period).first()
                return obj.pe_ratio if obj else None
        except Exception as e:
            logger.error(f"Error getting {data_type} for {stock.company_name} at {period}: {e}")
            return None

        return None


def generate_enhanced_portfolio_analysis_excel(scheme):
    """
    Generate enhanced Portfolio Analysis Excel file with calculated metrics

    Args:
        scheme: AMCFundScheme instance

    Returns:
        BytesIO: Excel file content
    """

    logger.info(f"Generating Fund Portfolio Analysis Excel with block-based structure for {scheme.name}")

    # Get fund holdings
    holdings = FundHolding.objects.filter(scheme=scheme).select_related('stock').order_by('-holding_percentage')

    if not holdings.exists():
        raise ValueError(f"No holdings data found for {scheme.name}")

    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio Analysis"

    # Initialize FundPortfolioExportGenerator
    generator = FundPortfolioExportGenerator(scheme)

    # Collect periods for this fund's stocks
    periods = generator.collect_all_periods_for_fund()

    logger.info(f"Periods collected:")
    logger.info(f"  - Market cap dates: {len(periods['market_cap_dates'])}")
    logger.info(f"  - TTM periods: {len(periods['ttm_periods'])}")
    logger.info(f"  - Quarterly periods: {len(periods['quarterly_periods'])}")
    logger.info(f"  - Annual years: {len(periods['annual_years'])}")
    logger.info(f"  - Share price dates: {len(periods['share_price_dates'])}")
    logger.info(f"  - PR dates: {len(periods['pr_dates'])}")
    logger.info(f"  - PE dates: {len(periods['pe_dates'])}")

    # Define block structure
    blocks = generator._define_block_structure(periods)

    # Calculate total columns
    total_columns = generator._calculate_total_columns(blocks)

    # Generate 8-row headers
    headers = generator._generate_import_style_headers(blocks, total_columns)

    logger.info(f"Generated {total_columns} columns with block-based structure")

    # Add Row 1: Fund name
    row_1 = [''] * total_columns
    row_1[0] = scheme.name
    ws.append(row_1)

    # Add Row 2: Column indicators
    ws.append(headers['row_2'])

    # Add Row 3: Portfolio as on: [date]
    latest_date = max(periods['market_cap_dates']) if periods['market_cap_dates'] else None
    row_3_text = f"Portfolio as on: {latest_date.strftime('%d %B %Y')}" if latest_date else "Portfolio as on: N/A"
    row_3 = [row_3_text] + [''] * (total_columns - 1)
    ws.append(row_3)

    # Add Rows 4-8: Remaining headers (blank rows + category headers + detail headers)
    for row_num in range(3, 9):
        header_row = headers[f'row_{row_num}']
        ws.append(header_row)

    # Add stock data rows (starting from row 9)
    logger.info(f"Populating {len(holdings)} stock rows using block-based structure")

    # Collect stock row data for TOTALS calculation
    stock_rows_data = []
    for idx, holding in enumerate(holdings, start=1):
        row_data = generator.populate_fund_stock_row(holding, blocks, total_columns)
        # NOTE: No S.No column - removed as per user requirement
        stock_rows_data.append(row_data)
        ws.append(row_data)

    # Add 5 blank rows gap before TOTALS (as per sample file)
    for _ in range(5):
        blank_row = [''] * total_columns
        ws.append(blank_row)

    # Calculate TOTALS row by summing each column across all stock rows
    totals_row = ['TOTALS'] + ['' for _ in range(total_columns - 1)]

    logger.info("Calculating TOTALS row by summing all numeric columns")

    # Sum columns 4 onwards (skip Company Name, Accord Code, Sector, Cap which are non-numeric)
    for col_idx in range(4, total_columns):
        column_sum = 0
        has_numeric_data = False

        for stock_row in stock_rows_data:
            value = stock_row[col_idx]
            # Try to add numeric values
            if value is not None and value != '':
                try:
                    column_sum += float(value)
                    has_numeric_data = True
                except (ValueError, TypeError):
                    pass  # Skip non-numeric values (e.g., text)

        # Set total if column had any numeric data
        if has_numeric_data:
            totals_row[col_idx] = column_sum

    ws.append(totals_row)
    totals_row_index = ws.max_row

    # Add 27 portfolio metric rows at bottom (22 with data + 5 blank)
    logger.info("Adding portfolio metric rows")
    add_portfolio_metric_rows(ws, scheme, generator.section_start_columns, periods, total_columns, totals_row_index)

    logger.info("Block-based export completed with portfolio metrics")

    # Apply formatting
    apply_portfolio_analysis_formatting(ws, total_columns)

    logger.info("Enhanced Excel generation completed successfully")

    # Save to BytesIO
    excel_content = BytesIO()
    wb.save(excel_content)
    excel_content.seek(0)

    return excel_content


def generate_recalculated_analysis_excel(scheme, filtered_holdings):
    """
    Generate recalculated Portfolio Analysis Excel file using filtered holdings.
    This is used for the "Recalculated Analysis" sheet in dual-sheet export.

    Args:
        scheme: AMCFundScheme instance
        filtered_holdings: QuerySet of FundHolding objects (pre-filtered by exclusion criteria)

    Returns:
        BytesIO: Excel file content with recalculated metrics
    """

    logger.info(f"Generating Recalculated Portfolio Analysis Excel for {scheme.name} with {filtered_holdings.count()} filtered holdings")

    if not filtered_holdings.exists():
        raise ValueError(f"No holdings data found after filtering for {scheme.name}")

    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Recalculated Analysis"

    # Initialize FundPortfolioExportGenerator
    generator = FundPortfolioExportGenerator(scheme)

    # Collect periods ONLY from filtered stocks
    periods = generator.collect_periods_for_filtered_holdings(filtered_holdings)

    logger.info(f"Periods collected from filtered stocks:")
    logger.info(f"  - Market cap dates: {len(periods['market_cap_dates'])}")
    logger.info(f"  - TTM periods: {len(periods['ttm_periods'])}")
    logger.info(f"  - Quarterly periods: {len(periods['quarterly_periods'])}")
    logger.info(f"  - Annual years: {len(periods['annual_years'])}")
    logger.info(f"  - Share price dates: {len(periods['share_price_dates'])}")
    logger.info(f"  - PR dates: {len(periods['pr_dates'])}")
    logger.info(f"  - PE dates: {len(periods['pe_dates'])}")

    # Define block structure
    blocks = generator._define_block_structure(periods)

    # Calculate total columns
    total_columns = generator._calculate_total_columns(blocks)

    # Generate 8-row headers
    headers = generator._generate_import_style_headers(blocks, total_columns)

    logger.info(f"Generated {total_columns} columns with block-based structure")

    # Add Row 1: Fund name
    row_1 = [''] * total_columns
    row_1[0] = scheme.name
    ws.append(row_1)

    # Add Row 2: Column indicators
    ws.append(headers['row_2'])

    # Add Row 3: Portfolio as on: [date]
    latest_date = max(periods['market_cap_dates']) if periods['market_cap_dates'] else None
    row_3_text = f"Portfolio as on: {latest_date.strftime('%d %B %Y')}" if latest_date else "Portfolio as on: N/A"
    row_3 = [row_3_text] + [''] * (total_columns - 1)
    ws.append(row_3)

    # Add Rows 4-8: Remaining headers (blank rows + category headers + detail headers)
    for row_num in range(3, 9):
        header_row = headers[f'row_{row_num}']
        ws.append(header_row)

    # Add stock data rows (starting from row 9) - using FILTERED holdings
    logger.info(f"Populating {filtered_holdings.count()} filtered stock rows using block-based structure")

    # Collect stock row data for TOTALS calculation
    stock_rows_data = []
    for idx, holding in enumerate(filtered_holdings, start=1):
        row_data = generator.populate_fund_stock_row(holding, blocks, total_columns)
        stock_rows_data.append(row_data)
        ws.append(row_data)

    # Add 5 blank rows gap before TOTALS (as per sample file)
    for _ in range(5):
        blank_row = [''] * total_columns
        ws.append(blank_row)

    # Calculate TOTALS row by summing each column across all stock rows
    totals_row = ['TOTALS'] + ['' for _ in range(total_columns - 1)]

    logger.info("Calculating TOTALS row by summing all numeric columns (from filtered stocks)")

    # Sum columns 4 onwards (skip Company Name, Accord Code, Sector, Cap which are non-numeric)
    for col_idx in range(4, total_columns):
        column_sum = 0
        has_numeric_data = False

        for stock_row in stock_rows_data:
            value = stock_row[col_idx]
            # Try to add numeric values
            if value is not None and value != '':
                try:
                    column_sum += float(value)
                    has_numeric_data = True
                except (ValueError, TypeError):
                    pass  # Skip non-numeric values (e.g., text)

        # Set total if column had any numeric data
        if has_numeric_data:
            totals_row[col_idx] = column_sum

    ws.append(totals_row)
    totals_row_index = ws.max_row

    # Add 27 portfolio metric rows at bottom (22 with data + 5 blank)
    # NOTE: Portfolio metrics are recalculated from TOTALS row, not from database
    logger.info("Adding portfolio metric rows (recalculated from filtered TOTALS)")
    add_portfolio_metric_rows(ws, scheme, generator.section_start_columns, periods, total_columns, totals_row_index)

    logger.info("Recalculated Excel generation completed")

    # Apply formatting
    apply_portfolio_analysis_formatting(ws, total_columns)

    logger.info("Recalculated Excel formatting applied successfully")

    # Save to BytesIO
    excel_content = BytesIO()
    wb.save(excel_content)
    excel_content.seek(0)

    return excel_content


def apply_portfolio_analysis_formatting(ws, total_columns):
    """
    Apply professional formatting to the Portfolio Analysis worksheet
    """

    # Header formatting
    header_font = Font(name='Calibri', size=11, bold=True)
    header_alignment = Alignment(horizontal='center', vertical='center')
    header_fill = PatternFill(start_color='E6E6FA', end_color='E6E6FA', fill_type='solid')

    # Apply formatting to header rows (1-8)
    for row_num in range(1, 9):
        for col_num in range(1, min(total_columns + 1, 100)):  # Limit formatting for performance
            cell = ws.cell(row=row_num, column=col_num)
            if row_num <= 3:
                cell.font = header_font
                cell.alignment = header_alignment
                cell.fill = header_fill

    # Data formatting
    data_font = Font(name='Calibri', size=10)
    data_alignment = Alignment(horizontal='right', vertical='center')

    # Metric row formatting (bold font, light yellow background)
    metric_font = Font(name='Calibri', size=10, bold=True)
    metric_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')

    # Calculate metric row range
    # Structure: 8 header rows + N stock rows + 5 blank rows + TOTALS row + 27 metric rows
    # Metric rows: 22 with data + 5 blank = 27 total
    total_rows = ws.max_row
    metric_start_row = total_rows - 26  # Last 27 rows are metrics (22 data + 5 blank)
    totals_row = metric_start_row - 1   # TOTALS row is before metrics

    logger.info(f"Applying formatting - Total rows: {total_rows}, Metric rows: {metric_start_row}-{total_rows}, Totals row: {totals_row}")

    # Apply data formatting to stock rows and identify special rows
    for row_num in range(9, total_rows + 1):
        for col_num in range(1, min(total_columns + 1, 100)):  # Limit columns for performance
            cell = ws.cell(row=row_num, column=col_num)

            # Check if this is a totals row
            if row_num == totals_row:
                cell.font = Font(name='Calibri', size=10, bold=True)
                cell.fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
                if col_num > 4:
                    cell.alignment = data_alignment

            # Check if this is a metric row
            elif row_num >= metric_start_row:
                cell.font = metric_font
                if col_num == 1:  # Metric label column
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                else:  # Metric value columns
                    # No color fill - user requested no colors
                    cell.alignment = data_alignment

            # Regular stock data rows
            else:
                cell.font = data_font
                if col_num > 4:  # Numeric columns
                    cell.alignment = data_alignment

    # Set column widths
    ws.column_dimensions['A'].width = 25  # Company Name / Metric Label
    ws.column_dimensions['B'].width = 15  # Accord Code
    ws.column_dimensions['C'].width = 20  # Sector
    ws.column_dimensions['D'].width = 10  # Cap

    logger.info("Excel formatting applied successfully with metric row highlighting")


def add_portfolio_metric_rows(ws, scheme, section_start_columns, periods, total_columns, totals_row_index):
    """
    Calculate and add 27 portfolio metric rows from TOTALS row data

    Args:
        ws: Worksheet object
        scheme: AMCFundScheme instance
        section_start_columns: Dict mapping data_type to column index
        periods: Dict with period lists (ttm_periods, quarterly_periods, etc.)
        total_columns: Total number of columns in sheet
        totals_row_index: Row index of TOTALS row in worksheet
    """
    logger.info(f"Calculating portfolio metrics from TOTALS row (index: {totals_row_index}) for {scheme.name}")

    # Step 1: Read TOTALS row data from worksheet
    totals_row_data = []
    for col_idx in range(total_columns):
        cell_value = ws.cell(row=totals_row_index, column=col_idx + 1).value
        totals_row_data.append(cell_value if cell_value is not None else 0)

    logger.info(f"Read TOTALS row with {len(totals_row_data)} columns")

    # Step 2: Build section column mapping
    section_cols = build_section_column_mapping(section_start_columns, periods)
    logger.info(f"Built section column mapping for {len(section_cols)} sections")

    # Step 3: Calculate all metrics from TOTALS row
    patm_metrics = calculate_patm_from_totals(totals_row_data, section_cols)
    qoq_metrics = calculate_qoq_from_totals(totals_row_data, section_cols)
    yoy_metrics = calculate_yoy_from_totals(totals_row_data, section_cols)
    cagr_metrics = calculate_6yr_cagr_from_totals(totals_row_data, section_cols)
    pe_pr_metrics = calculate_pe_pr_from_totals(totals_row_data, section_cols)
    pe_pr_avgs = calculate_pe_pr_averages_from_totals(totals_row_data, section_cols)
    reval_deval = calculate_reval_deval_from_totals({**pe_pr_metrics, **pe_pr_avgs})
    pr_extremes = calculate_pr_10q_extremes_from_totals(totals_row_data, section_cols)
    pe_yield = calculate_pe_yield_from_totals(pe_pr_metrics)
    growth = calculate_growth_from_totals(cagr_metrics)
    bond_rate = get_bond_rate()

    logger.info("Calculated all portfolio metrics from TOTALS")

    # Step 4: Define 27 metric rows
    metric_rows = [
        {
            'label': 'PATM',
            'data': patm_metrics,
            'sections': ['ttm_pat', 'ttm_pat_free_float', 'quarterly_pat', 'quarterly_pat_free_float']
        },
        {
            'label': 'QoQ',
            'data': qoq_metrics,
            'sections': ['quarterly_revenue', 'quarterly_revenue_free_float',
                        'quarterly_pat', 'quarterly_pat_free_float']
        },
        {
            'label': 'YoY',
            'data': yoy_metrics,
            'sections': ['ttm_revenue', 'ttm_revenue_free_float', 'quarterly_revenue',
                        'quarterly_revenue_free_float', 'ttm_pat', 'ttm_pat_free_float',
                        'quarterly_pat', 'quarterly_pat_free_float']
        },
        {
            'label': '6 year CAGR',
            'data': cagr_metrics,
            'sections': ['ttm_revenue', 'ttm_revenue_free_float', 'ttm_pat', 'ttm_pat_free_float']
        },
        None,  # Blank
        {
            'label': 'Current PE',
            'data': {'value': pe_pr_metrics.get('current_pe')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '2 year average',
            'data': {'value': pe_pr_avgs.get('pe_2yr_avg')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '5 year average',
            'data': {'value': pe_pr_avgs.get('pe_5yr_avg')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '2 years - Reval / Deval',
            'data': {'value': reval_deval.get('pe_2yr_reval_deval')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '5 years - Reval / Deval',
            'data': {'value': reval_deval.get('pe_5yr_reval_deval')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        None,  # Blank
        {
            'label': 'Current PR',
            'data': {'value': pe_pr_metrics.get('current_pr')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '2 year average',
            'data': {'value': pe_pr_avgs.get('pr_2yr_avg')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '5 year average',
            'data': {'value': pe_pr_avgs.get('pr_5yr_avg')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '2 years - Reval / Deval',
            'data': {'value': reval_deval.get('pr_2yr_reval_deval')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '5 years - Reval / Deval',
            'data': {'value': reval_deval.get('pr_5yr_reval_deval')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        None,  # Blank
        {
            'label': '10 quarter- PR- low',
            'data': {'value': pr_extremes.get('pr_10q_low')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': '10 quarter- PR- high',
            'data': {'value': pr_extremes.get('pr_10q_high')},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        None,  # Blank
        {
            'label': 'Alpha over the bond- CAGR',
            'data': {'value': 0.0},  # TODO: Implement if needed
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': 'Alpha- Absolute',
            'data': {'value': 0.0},  # TODO: Implement if needed
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': 'PE Yield',
            'data': {'value': pe_yield},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': 'Growth',
            'data': {'value': growth},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        {
            'label': 'Bond Rate',
            'data': {'value': bond_rate},
            'sections': ['market_cap_free_float'],
            'single_value': True
        },
        None,  # Blank
        None,  # Blank
    ]

    # Step 5: Write rows to worksheet
    for row_def in metric_rows:
        metric_row = create_metric_row(row_def, section_cols, total_columns)
        ws.append(metric_row)

    logger.info(f"Added 27 portfolio metric rows calculated from TOTALS")