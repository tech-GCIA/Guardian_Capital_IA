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

    def _define_block_structure(self, periods):
        """
        Override to add fund-specific data block after basic info.
        Maintains same structure as parent but inserts Fund Data block.
        """
        blocks = [
            {
                'name': 'basic_info',
                'type': 'fixed',
                'size': 13,
                'columns': [
                    'S. No.', 'Company Name', 'Accord Code', 'Sector', 'Cap',
                    'Free Float', '6 Year CAGR', 'TTM', '6 Year CAGR', 'TTM',
                    'Current', '2 Yr Avg', 'Reval/deval'
                ]
            },
            {'name': 'sep_0', 'type': 'separator', 'size': 1},
            {
                'name': 'fund_data',
                'type': 'fixed',
                'size': 4,
                'columns': ['Weights', 'Factor', 'Value', 'No.of shares']
            },
            {'name': 'sep_1', 'type': 'separator', 'size': 1},
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
                    # Basic stock info (13 columns)
                    row_data[current_col] = ''  # S.No (will be filled with row number)
                    row_data[current_col + 1] = stock.company_name
                    row_data[current_col + 2] = stock.accord_code or ''
                    row_data[current_col + 3] = stock.sector or ''
                    row_data[current_col + 4] = stock.cap or ''
                    row_data[current_col + 5] = stock.free_float or ''
                    row_data[current_col + 6] = stock.revenue_6yr_cagr or ''
                    row_data[current_col + 7] = stock.revenue_ttm or ''
                    row_data[current_col + 8] = stock.pat_6yr_cagr or ''
                    row_data[current_col + 9] = stock.pat_ttm or ''
                    row_data[current_col + 10] = stock.current_value or ''
                    row_data[current_col + 11] = stock.two_yr_avg or ''
                    row_data[current_col + 12] = stock.reval_deval or ''
                    current_col += 13

                elif block['name'] == 'fund_data':
                    # Fund-specific data (4 columns)
                    weights = (holding.holding_percentage / 100) if holding.holding_percentage else 0
                    factor = (holding.market_value / total_portfolio_value) if holding.market_value and total_portfolio_value else 0

                    row_data[current_col] = weights
                    row_data[current_col + 1] = factor
                    row_data[current_col + 2] = holding.market_value or 0
                    row_data[current_col + 3] = holding.number_of_shares or 0
                    current_col += 4

                elif block['name'] == 'identifiers':
                    # Stock identifiers (3 columns)
                    row_data[current_col] = stock.bse_code or ''
                    row_data[current_col + 1] = stock.nse_symbol or ''
                    row_data[current_col + 2] = stock.isin or ''
                    current_col += 3

            elif block['type'] == 'dynamic':
                # Populate dynamic data columns using parent's logic
                data_type = block['data_type']
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

    # Add Row 2-8: Block-based headers
    for row_num in range(2, 9):
        header_row = headers[f'row_{row_num}']
        ws.append(header_row)

    # Add stock data rows (starting from row 9)
    logger.info(f"Populating {len(holdings)} stock rows using block-based structure")

    for idx, holding in enumerate(holdings, start=1):
        row_data = generator.populate_fund_stock_row(holding, blocks, total_columns)
        row_data[0] = idx  # Set S.No
        ws.append(row_data)

    # Add TOTALS row
    totals_row = ['TOTALS'] + [''] * (total_columns - 1)
    total_market_cap = sum(h.market_value or 0 for h in holdings if h.market_value)
    total_holdings_pct = sum(h.holding_percentage or 0 for h in holdings if h.holding_percentage)

    # Find column indices for totals in block structure
    # Basic info: columns 0-12, sep: 13, fund_data: columns 14-17
    totals_row[4] = total_market_cap  # Cap column in basic_info
    totals_row[14] = total_holdings_pct / 100 if total_holdings_pct else 0  # Weights in fund_data
    totals_row[16] = total_market_cap  # Value in fund_data

    ws.append(totals_row)

    # NOTE: In block-based structure, portfolio metrics are not added as separate rows at bottom
    # All stock data including metrics are in the stock rows themselves
    # Portfolio-level weighted metrics are stored in PortfolioMetricsLog table for future use

    logger.info("Block-based export completed - all data is in stock rows with block structure")

    # Apply formatting
    apply_portfolio_analysis_formatting(ws, total_columns)

    logger.info("Enhanced Excel generation completed successfully")

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
    # Metric rows start after: 8 header rows + N stock rows + 1 totals row
    # We know metric rows are the last 22 rows (plus totals row before them)
    total_rows = ws.max_row
    metric_start_row = total_rows - 21  # Last 22 rows are metrics
    totals_row = metric_start_row - 1

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
                    cell.fill = metric_fill
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