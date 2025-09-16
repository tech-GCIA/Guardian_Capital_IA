"""
Enhanced Excel Export with Portfolio Analysis Metrics
=====================================================

This module provides the enhanced Excel export functionality that includes
calculated portfolio analysis metrics in the Portfolio Analysis format.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from django.utils import timezone
from io import BytesIO
from .metrics_calculator import DynamicHeaderGenerator
from .models import FundMetricsLog, AMCFundScheme, FundHolding
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


def generate_enhanced_portfolio_analysis_excel(scheme):
    """
    Generate enhanced Portfolio Analysis Excel file with calculated metrics

    Args:
        scheme: AMCFundScheme instance

    Returns:
        BytesIO: Excel file content
    """

    logger.info(f"Generating enhanced Portfolio Analysis Excel for {scheme.name}")

    # Get fund holdings with related stock data
    holdings = FundHolding.objects.filter(scheme=scheme).select_related('stock').prefetch_related(
        'stock__market_cap_data', 'stock__ttm_data', 'stock__quarterly_data',
        'stock__annual_ratios', 'stock__price_data'
    ).order_by('-holding_percentage')

    if not holdings.exists():
        raise ValueError(f"No holdings data found for {scheme.name}")

    # Create workbook with Portfolio Analysis format
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio Analysis"

    # Generate dynamic headers and available periods
    header_generator = DynamicHeaderGenerator()
    available_periods = header_generator.get_available_periods_for_fund(scheme)

    logger.info(f"Available periods: {len(available_periods['all'])} total periods")

    # Create dynamic headers structure
    dynamic_headers = header_generator.generate_dynamic_headers(available_periods)

    # Calculate total columns dynamically
    base_columns = 9  # Company Name, Accord Code, Sector, Cap, Market Cap, Weights, Factor, Value, No.of shares

    # Add columns for each period type
    market_cap_columns = len(available_periods['market_cap'])
    ttm_columns = len(available_periods['ttm']) * 2  # Revenue + PAT for each period
    quarterly_columns = len(available_periods['quarterly']) * 2  # Revenue + PAT for each period
    metrics_columns = len(available_periods['all']) * 22  # 22 metrics for each period
    identifier_columns = 3  # BSE, NSE, ISIN

    total_columns = (base_columns + market_cap_columns + ttm_columns +
                    quarterly_columns + metrics_columns + identifier_columns)

    logger.info(f"Generated {total_columns} columns dynamically")

    # Add professional 8-row header structure
    portfolio_date = holdings.first().holding_date if holdings.first().holding_date else datetime.now().date()

    # ROW 1: Fund name only
    row_1 = [''] * total_columns
    row_1[0] = scheme.name
    ws.append(row_1)

    # ROW 2: Column position indicators
    row_2 = [''] * total_columns
    for i in range(min(total_columns, 100)):  # Show first 100 column numbers
        row_2[i] = f"Col_{i+1}"
    ws.append(row_2)

    # ROW 3: Portfolio date + main headers
    row_3 = [''] * total_columns
    row_3[0] = f"Portfolio as on: {portfolio_date.strftime('%d %B %Y')}"

    # Add section headers
    col_index = 0

    # Basic info headers
    basic_headers = ['Company Name', 'Accord Code', 'Sector', 'Cap', 'Market Cap', 'Weights', 'Factor', 'Value', 'No.of shares']
    for i, header in enumerate(basic_headers):
        row_3[col_index] = header
        col_index += 1

    # Market cap section
    if available_periods['market_cap']:
        row_3[col_index] = 'Market Cap Data'
        col_index += len(available_periods['market_cap'])

    # TTM section
    if available_periods['ttm']:
        row_3[col_index] = 'TTM Data'
        col_index += len(available_periods['ttm']) * 2

    # Quarterly section
    if available_periods['quarterly']:
        row_3[col_index] = 'Quarterly Data'
        col_index += len(available_periods['quarterly']) * 2

    # Metrics section
    if available_periods['all']:
        row_3[col_index] = 'Portfolio Analysis Metrics'
        col_index += len(available_periods['all']) * 22

    # Identifiers section
    row_3[col_index] = 'Stock Identifiers'

    ws.append(row_3)

    # ROW 4-7: Additional header details
    for row_num in range(4, 8):
        ws.append([''] * total_columns)

    # ROW 8: Detailed column headers
    row_8 = [''] * total_columns
    col_index = 0

    # Basic columns
    for header in basic_headers:
        row_8[col_index] = header
        col_index += 1

    # Market cap period headers
    for period in available_periods['market_cap']:
        row_8[col_index] = f"MC_{period}"
        col_index += 1

    # TTM period headers
    for period in available_periods['ttm']:
        row_8[col_index] = f"TTM_Rev_{period}"
        row_8[col_index + 1] = f"TTM_PAT_{period}"
        col_index += 2

    # Quarterly period headers
    for period in available_periods['quarterly']:
        row_8[col_index] = f"Q_Rev_{period}"
        row_8[col_index + 1] = f"Q_PAT_{period}"
        col_index += 2

    # Metrics headers for each period
    metric_names = [
        'PATM', 'QoQ', 'YoY', 'Rev_6Y_CAGR', 'PAT_6Y_CAGR', 'Curr_PE', 'PE_2Y_Avg', 'PE_5Y_Avg',
        'PE_2Y_RevDev', 'PE_5Y_RevDev', 'Curr_PR', 'PR_2Y_Avg', 'PR_5Y_Avg', 'PR_2Y_RevDev',
        'PR_5Y_RevDev', 'PR_10Q_Low', 'PR_10Q_High', 'Alpha_Bond', 'Alpha_Abs', 'PE_Yield',
        'Growth', 'Bond_Rate'
    ]

    for period in available_periods['all']:
        for metric in metric_names:
            row_8[col_index] = f"{metric}_{period}"
            col_index += 1

    # Identifier headers
    row_8[col_index] = 'BSE Code'
    row_8[col_index + 1] = 'NSE Symbol'
    row_8[col_index + 2] = 'ISIN'

    ws.append(row_8)

    # Add stock data rows with integrated fund holdings and calculated metrics
    logger.info("Populating stock data rows with calculated metrics")

    for holding in holdings:
        stock = holding.stock
        row_data = [''] * total_columns
        col_index = 0

        # Basic stock info (columns 0-8)
        row_data[col_index] = stock.company_name; col_index += 1
        row_data[col_index] = stock.accord_code; col_index += 1
        row_data[col_index] = stock.sector; col_index += 1
        row_data[col_index] = stock.cap; col_index += 1

        # Fund-specific data (columns 4-8)
        row_data[col_index] = holding.market_value or 0; col_index += 1  # Market Cap

        weights = (holding.holding_percentage / 100) if holding.holding_percentage else 0
        row_data[col_index] = weights; col_index += 1  # Weights

        # Calculate factor (Position Value / Total Market Cap)
        total_portfolio_value = sum(h.market_value for h in holdings if h.market_value)
        factor = (holding.market_value / total_portfolio_value) if holding.market_value and total_portfolio_value else 0
        row_data[col_index] = factor; col_index += 1  # Factor

        row_data[col_index] = holding.market_value or 0; col_index += 1  # Value
        row_data[col_index] = holding.number_of_shares or 0; col_index += 1  # No.of shares

        # Market cap data for all available periods
        for period in available_periods['market_cap']:
            market_cap_entry = stock.market_cap_data.filter(date=period).first()
            row_data[col_index] = market_cap_entry.market_cap if market_cap_entry else 0
            col_index += 1

        # TTM data for all available periods
        for period in available_periods['ttm']:
            try:
                ttm_entry = stock.ttm_data.filter(period=period).first()
                if ttm_entry:
                    row_data[col_index] = ttm_entry.ttm_revenue or 0
                    row_data[col_index + 1] = ttm_entry.ttm_pat or 0
                else:
                    row_data[col_index] = 0
                    row_data[col_index + 1] = 0
            except Exception as e:
                logger.error(f"Error getting TTM data for stock {stock.company_name}, period {period}: {e}")
                row_data[col_index] = 0
                row_data[col_index + 1] = 0
            col_index += 2

        # Quarterly data for all available periods
        for period in available_periods['quarterly']:
            try:
                quarterly_entry = stock.quarterly_data.filter(period=period).first()
                if quarterly_entry:
                    row_data[col_index] = quarterly_entry.quarterly_revenue or 0
                    row_data[col_index + 1] = quarterly_entry.quarterly_pat or 0
                else:
                    row_data[col_index] = 0
                    row_data[col_index + 1] = 0
            except Exception as e:
                logger.error(f"Error getting Quarterly data for stock {stock.company_name}, period {period}: {e}")
                row_data[col_index] = 0
                row_data[col_index + 1] = 0
            col_index += 2

        # Add calculated metrics for all available periods
        logger.debug(f"Adding calculated metrics for stock {stock.company_name}")

        # Get all calculated metrics for this stock
        stock_metrics = FundMetricsLog.objects.filter(
            scheme=scheme,
            stock=stock
        ).order_by('-period_date')  # Explicit ordering since we removed default model ordering

        # Create a metrics lookup by period
        metrics_by_period = {}
        for metric in stock_metrics:
            metrics_by_period[metric.period_date] = metric

        # Add metrics for each period (22 metrics per period)
        for period in available_periods['all']:
            metrics = metrics_by_period.get(period)
            if metrics:
                # Add all 22 calculated metrics
                metric_values = [
                    metrics.patm or 0,
                    metrics.qoq_growth or 0,
                    metrics.yoy_growth or 0,
                    metrics.revenue_6yr_cagr or 0,
                    metrics.pat_6yr_cagr or 0,
                    metrics.current_pe or 0,
                    metrics.pe_2yr_avg or 0,
                    metrics.pe_5yr_avg or 0,
                    metrics.pe_2yr_reval_deval or 0,
                    metrics.pe_5yr_reval_deval or 0,
                    metrics.current_pr or 0,
                    metrics.pr_2yr_avg or 0,
                    metrics.pr_5yr_avg or 0,
                    metrics.pr_2yr_reval_deval or 0,
                    metrics.pr_5yr_reval_deval or 0,
                    metrics.pr_10q_low or 0,
                    metrics.pr_10q_high or 0,
                    metrics.alpha_bond_cagr or 0,
                    metrics.alpha_absolute or 0,
                    metrics.pe_yield or 0,
                    metrics.growth_rate or 0,
                    metrics.bond_rate or 0
                ]
            else:
                # No metrics calculated for this period - fill with zeros
                metric_values = [0] * 22

            # Add metric values to row
            for value in metric_values:
                if col_index < len(row_data):
                    row_data[col_index] = value
                    col_index += 1

        # Add stock identifiers at the end
        if col_index < len(row_data) - 2:
            row_data[col_index] = stock.bse_code or ''
            row_data[col_index + 1] = stock.nse_symbol or ''
            row_data[col_index + 2] = stock.isin or ''

        ws.append(row_data)

    # Add portfolio totals row (row 36 equivalent)
    totals_row = ['TOTALS'] + [''] * (total_columns - 1)

    # Calculate portfolio-level totals for key metrics
    total_market_cap = sum(h.market_value for h in holdings if h.market_value)
    total_holdings_pct = sum(h.holding_percentage for h in holdings if h.holding_percentage)

    totals_row[4] = total_market_cap  # Total Market Cap
    totals_row[5] = total_holdings_pct / 100 if total_holdings_pct else 0  # Total Weights
    totals_row[7] = total_market_cap  # Total Value

    ws.append(totals_row)

    # Add metric summary rows (equivalent to Excel rows 37-61)
    logger.info("Adding portfolio-level metric summary rows")

    metric_labels = [
        'PATM', 'QoQ', 'YoY', '6 year CAGR', 'Current PE',
        '2 year average', '5 year average', '2 years - Reval / Deval',
        '5 years - Reval / Deval', 'Current PR', '2 year average',
        '5 year average', '2 years - Reval / Deval',
        '5 years - Reval / Deval', '10 quarter- PR- low',
        '10 quarter- PR- high', 'Alpha over the bond- CAGR',
        'Alpha- Absolute', 'PE Yield', 'Growth', 'Bond Rate'
    ]

    # Add metric summary rows with portfolio-level values
    for i, label in enumerate(metric_labels):
        metric_row = [label] + [''] * (total_columns - 1)

        # Map metric labels to actual field names in AMCFundScheme
        field_mapping = {
            'PATM': 'latest_patm',
            'QoQ': 'latest_qoq_growth',
            'YoY': 'latest_yoy_growth',
            '6 year CAGR': 'latest_revenue_6yr_cagr',
            'Current PE': 'latest_current_pe',
            '2 year average': 'latest_pe_2yr_avg',
            '5 year average': 'latest_pe_5yr_avg',
            '2 years - Reval / Deval': 'latest_pe_2yr_reval_deval',
            '5 years - Reval / Deval': 'latest_pe_5yr_reval_deval',
            'Current PR': 'latest_current_pr',
            '10 quarter- PR- low': 'latest_pr_10q_low',
            '10 quarter- PR- high': 'latest_pr_10q_high',
            'Alpha over the bond- CAGR': 'latest_alpha_bond_cagr',
            'Alpha- Absolute': 'latest_alpha_absolute',
            'PE Yield': 'latest_pe_yield',
            'Growth': 'latest_growth_rate',
            'Bond Rate': 'latest_bond_rate'
        }

        # Get the appropriate field value from scheme
        field_name = field_mapping.get(label)
        if field_name and hasattr(scheme, field_name):
            metric_value = getattr(scheme, field_name) or 0
            metric_row[9] = metric_value  # Add in first data column

        ws.append(metric_row)

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

    # Apply data formatting to a reasonable range
    for row_num in range(9, min(ws.max_row + 1, 100)):  # Limit for performance
        for col_num in range(1, min(total_columns + 1, 50)):  # Limit columns for performance
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = data_font
            if col_num > 4:  # Numeric columns
                cell.alignment = data_alignment

    # Set column widths
    ws.column_dimensions['A'].width = 25  # Company Name
    ws.column_dimensions['B'].width = 15  # Accord Code
    ws.column_dimensions['C'].width = 20  # Sector
    ws.column_dimensions['D'].width = 10  # Cap

    logger.info("Excel formatting applied successfully")