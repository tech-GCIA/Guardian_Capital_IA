"""
Portfolio Analysis Metrics Calculator
====================================

This module implements the calculation engine for all 22 portfolio analysis metrics
based on the formulas discovered in the Excel analysis.
"""

import logging
import uuid
from datetime import datetime, timedelta
from django.db import transaction
from django.utils import timezone
from django.db.models import Q, Min, Max, Avg
from .models import (
    AMCFundScheme, FundHolding, Stock, StockMarketCap, StockTTMData,
    StockQuarterlyData, StockAnnualRatios, StockPrice, FundMetricsLog,
    MetricsCalculationSession, Customer
)

# Set up logging
logger = logging.getLogger(__name__)

class PortfolioMetricsCalculator:
    """
    Main calculator class for portfolio analysis metrics
    """

    def __init__(self, session_id=None, user=None):
        self.session_id = session_id or str(uuid.uuid4())
        self.user = user
        self.progress_session = None

    def calculate_metrics_for_all_funds(self, limit_periods=None):
        """
        Calculate all 22 metrics for all funds with real-time progress tracking

        Args:
            limit_periods (dict): Optional limits like {'quarters': 10, 'years': 6}
        """

        # Get all funds with holdings
        funds_with_holdings = AMCFundScheme.objects.filter(
            holdings__isnull=False, is_active=True
        ).distinct()

        total_funds = funds_with_holdings.count()
        logger.info(f"Starting metrics calculation for {total_funds} funds")

        # Create progress session only if not already created
        if not self.progress_session:
            self.progress_session = MetricsCalculationSession.objects.create(
                session_id=self.session_id,
                user=self.user,
                total_funds=total_funds,
                status='started'
            )

        try:
            for index, fund in enumerate(funds_with_holdings):
                logger.info(f"Processing fund {index+1}/{total_funds}: {fund.name}")

                # Update progress
                self.update_progress(
                    processed_funds=index,
                    current_fund_name=fund.name,
                    status='processing'
                )

                # Calculate metrics for this fund
                self.calculate_all_metrics_for_fund(fund.amcfundscheme_id, limit_periods)

                # Update progress
                progress_percentage = ((index + 1) / total_funds) * 100
                self.update_progress(
                    processed_funds=index + 1,
                    progress_percentage=progress_percentage
                )

            # Mark as completed
            self.update_progress(status='completed', completed_at=timezone.now())
            logger.info("Metrics calculation completed successfully")

        except Exception as e:
            logger.error(f"Error in metrics calculation: {str(e)}")
            self.update_progress(
                status='failed',
                error_message=str(e),
                completed_at=timezone.now()
            )
            raise

    def calculate_all_metrics_for_fund(self, scheme_id, limit_periods=None):
        """
        Calculate all 22 metrics for all periods for a specific fund
        """

        scheme = AMCFundScheme.objects.get(amcfundscheme_id=scheme_id)
        holdings = FundHolding.objects.filter(scheme=scheme).select_related('stock')

        if not holdings.exists():
            logger.warning(f"No holdings found for fund {scheme.name}")
            return

        logger.info(f"Processing {holdings.count()} holdings for fund {scheme.name}")

        # Track latest metrics for portfolio-level aggregation
        latest_metrics = {}
        latest_period_date = None

        for holding in holdings:
            stock = holding.stock

            # Update progress with current stock
            if self.progress_session:
                self.update_progress(current_stock_name=stock.company_name)

            logger.debug(f"Processing stock: {stock.company_name}")

            # Get all available periods for this stock
            periods = self.get_available_periods(stock, limit_periods)

            if not periods:
                logger.warning(f"No periods found for stock {stock.company_name}")
                continue

            for period in periods:
                metrics = self.calculate_metrics_for_period(stock, period)

                # Convert period date to proper format for DateField
                period_date = period['date']

                # Handle different date formats
                if isinstance(period_date, str):
                    if len(period_date) == 6:  # YYYYMM format (TTM/Quarterly)
                        from datetime import datetime
                        try:
                            year = int(period_date[:4])
                            month = int(period_date[4:6])
                            period_date = datetime(year, month, 1).date()
                        except (ValueError, IndexError) as e:
                            logger.error(f"Error converting period date {period_date}: {e}")
                            continue  # Skip this period if date conversion fails
                    else:
                        # Try to parse other string date formats
                        from datetime import datetime
                        try:
                            period_date = datetime.strptime(str(period_date), '%Y-%m-%d').date()
                        except ValueError:
                            logger.error(f"Unable to parse date format: {period_date}")
                            continue  # Skip this period if date conversion fails
                elif not hasattr(period_date, 'year'):  # Not a date object
                    logger.error(f"Invalid date type: {type(period_date)} for {period_date}")
                    continue  # Skip this period

                # Store in log table with update_or_create
                try:
                    log_entry, created = FundMetricsLog.objects.update_or_create(
                        scheme=scheme,
                        stock=stock,
                        period_date=period_date,
                        period_type=period['type'],
                        defaults=metrics
                    )
                except Exception as db_error:
                    logger.error(f"Database error creating FundMetricsLog: {db_error}")
                    logger.error(f"Period date: {period_date}, Period type: {period['type']}")
                    logger.error(f"Stock: {stock}, Scheme: {scheme}")
                    continue  # Skip this period and continue with next

                # Keep track of latest period metrics for AMCFundScheme update
                if not latest_period_date or period_date > latest_period_date:
                    latest_period_date = period_date
                    # Aggregate metrics for this period
                    for metric, value in metrics.items():
                        if metric not in latest_metrics:
                            latest_metrics[metric] = []
                        # Only include non-None values
                        if value is not None:
                            latest_metrics[metric].append(value)

        # Update AMCFundScheme with portfolio-level aggregated metrics
        self.update_fund_latest_metrics(scheme, latest_metrics)

    def get_available_periods(self, stock, limit_periods=None):
        """
        Get all available periods for a stock based on actual data availability
        """

        periods = []

        logger.debug(f"Getting available periods for stock: {stock}")

        # Get available TTM periods - NO ORDER_BY to avoid field conflicts
        try:
            ttm_queryset = StockTTMData.objects.filter(stock=stock)
            ttm_list = list(ttm_queryset.values_list('period', flat=True))

            # Sort in Python instead of database to avoid ordering conflicts
            ttm_list.sort(reverse=True)

            if limit_periods and 'quarters' in limit_periods:
                ttm_list = ttm_list[:limit_periods['quarters']]

            for period in ttm_list:
                periods.append({
                    'date': period,
                    'type': 'ttm'
                })

            logger.debug(f"Found {len(ttm_list)} TTM periods for stock {stock}")
        except Exception as e:
            logger.error(f"Error getting TTM periods for stock {stock}: {e}")
            logger.error(f"Full error: {repr(e)}")

        # Get available quarterly periods - NO ORDER_BY to avoid field conflicts
        try:
            quarterly_queryset = StockQuarterlyData.objects.filter(stock=stock)
            quarterly_list = list(quarterly_queryset.values_list('period', flat=True))

            # Sort in Python instead of database to avoid ordering conflicts
            quarterly_list.sort(reverse=True)

            if limit_periods and 'quarters' in limit_periods:
                quarterly_list = quarterly_list[:limit_periods['quarters']]

            for period in quarterly_list:
                periods.append({
                    'date': period,
                    'type': 'quarterly'
                })

            logger.debug(f"Found {len(quarterly_list)} Quarterly periods for stock {stock}")
        except Exception as e:
            logger.error(f"Error getting Quarterly periods for stock {stock}: {e}")
            logger.error(f"Full error: {repr(e)}")

        # Remove duplicates and sort by date (latest first)
        unique_periods = []
        seen_dates = set()
        for period in periods:
            if period['date'] not in seen_dates:
                unique_periods.append(period)
                seen_dates.add(period['date'])

        # Sort in Python to avoid any Django ordering issues
        unique_periods.sort(key=lambda x: x['date'], reverse=True)

        logger.debug(f"Returning {len(unique_periods)} unique periods for stock {stock}")
        return unique_periods

    def calculate_metrics_for_period(self, stock, period):
        """
        Calculate all 22 metrics for a specific stock and period
        """

        logger.debug(f"Calculating metrics for {stock.company_name} - {period['date']}")

        metrics = {}

        try:
            # Convert period date to proper format for database queries
            period_date_original = period['date']

            # Convert YYYYMM string format to proper date for market cap queries
            if isinstance(period_date_original, str) and len(period_date_original) == 6:
                from datetime import datetime
                year = int(period_date_original[:4])
                month = int(period_date_original[4:6])
                period_date_for_market_cap = datetime(year, month, 1).date()
            else:
                period_date_for_market_cap = period_date_original

            # Get data for calculations (all data up to this period)
            market_cap_data = StockMarketCap.objects.filter(
                stock=stock,
                date__lte=period_date_for_market_cap
            )

            # For TTM and Quarterly data, use the original string format since they have CharField periods
            try:
                ttm_data = StockTTMData.objects.filter(
                    stock=stock,
                    period__lte=period_date_original
                )
            except Exception as e:
                logger.error(f"Error querying TTM data for stock {stock}: {e}")
                ttm_data = StockTTMData.objects.none()

            try:
                quarterly_data = StockQuarterlyData.objects.filter(
                    stock=stock,
                    period__lte=period_date_original
                )
            except Exception as e:
                logger.error(f"Error querying Quarterly data for stock {stock}: {e}")
                quarterly_data = StockQuarterlyData.objects.none()

            price_data = StockPrice.objects.filter(
                stock=stock,
                price_date__lte=period_date_for_market_cap
            )

            # 1. PATM (Profit After Tax Margin)
            metrics['patm'] = self.calculate_patm(ttm_data)

            # 2. QoQ (Quarter over Quarter)
            metrics['qoq_growth'] = self.calculate_qoq_growth(quarterly_data)

            # 3. YoY (Year over Year)
            metrics['yoy_growth'] = self.calculate_yoy_growth(quarterly_data)

            # 4-5. 6 Year CAGR (Revenue & PAT)
            metrics['revenue_6yr_cagr'] = self.calculate_revenue_6yr_cagr(ttm_data)
            metrics['pat_6yr_cagr'] = self.calculate_pat_6yr_cagr(ttm_data)

            # 6-11. PE Ratios and related metrics
            pe_metrics = self.calculate_pe_metrics(market_cap_data, ttm_data)
            metrics.update(pe_metrics)

            # 12-17. PR Ratios and related metrics
            pr_metrics = self.calculate_pr_metrics(market_cap_data, ttm_data)
            metrics.update(pr_metrics)

            # 18-22. Alpha, PE Yield, Growth, Bond Rate
            alpha_metrics = self.calculate_alpha_and_other_metrics(stock, period, ttm_data, market_cap_data)
            metrics.update(alpha_metrics)

        except Exception as e:
            logger.error(f"Error calculating metrics for {stock.company_name} - {period['date']}: {e}")
            # Set all metrics to 0 for missing data as per requirement
            metrics = {
                'patm': 0.0, 'qoq_growth': 0.0, 'yoy_growth': 0.0,
                'revenue_6yr_cagr': 0.0, 'pat_6yr_cagr': 0.0,
                'current_pe': 0.0, 'pe_2yr_avg': 0.0, 'pe_5yr_avg': 0.0,
                'pe_2yr_reval_deval': 0.0, 'pe_5yr_reval_deval': 0.0,
                'current_pr': 0.0, 'pr_2yr_avg': 0.0, 'pr_5yr_avg': 0.0,
                'pr_2yr_reval_deval': 0.0, 'pr_5yr_reval_deval': 0.0,
                'pr_10q_low': 0.0, 'pr_10q_high': 0.0, 'alpha_bond_cagr': 0.0,
                'alpha_absolute': 0.0, 'pe_yield': 0.0, 'growth_rate': 0.0, 'bond_rate': 0.0
            }

        return metrics

    # INDIVIDUAL METRIC CALCULATION METHODS
    # =====================================

    def calculate_cagr(self, start_value, end_value, years):
        """
        CAGR Formula: ((end_value/start_value)**(1/years)) - 1
        """
        if not start_value or not end_value or years <= 0 or start_value <= 0:
            return 0.0
        try:
            return ((end_value / start_value) ** (1 / years)) - 1
        except (ZeroDivisionError, ValueError, OverflowError):
            return 0.0

    def calculate_patm(self, ttm_data):
        """
        PATM Formula: PAT ÷ Revenue × 100
        """
        if ttm_data.exists():
            latest = ttm_data.first()
            if latest.ttm_pat and latest.ttm_revenue and latest.ttm_revenue != 0:
                return (latest.ttm_pat / latest.ttm_revenue) * 100
        return 0.0

    def calculate_qoq_growth(self, quarterly_data):
        """
        QoQ Formula: (Current Quarter - Previous Quarter) ÷ Previous Quarter
        """
        if quarterly_data.count() >= 2:
            current = quarterly_data[0]
            previous = quarterly_data[1]

            if (current.quarterly_revenue and previous.quarterly_revenue and
                previous.quarterly_revenue != 0):
                return ((current.quarterly_revenue - previous.quarterly_revenue) /
                       previous.quarterly_revenue)
        return 0.0

    def calculate_yoy_growth(self, quarterly_data):
        """
        YoY Formula: (Current Period - Same Period Last Year) ÷ Same Period Last Year
        """
        if quarterly_data.count() >= 4:  # Need at least 4 quarters for YoY
            current = quarterly_data[0]
            year_ago = quarterly_data[3]  # 4 quarters ago

            if (current.quarterly_revenue and year_ago.quarterly_revenue and
                year_ago.quarterly_revenue != 0):
                return ((current.quarterly_revenue - year_ago.quarterly_revenue) /
                       year_ago.quarterly_revenue)
        return 0.0

    def calculate_revenue_6yr_cagr(self, ttm_data):
        """
        Revenue 6Y CAGR Formula: POWER(Current_TTM_Revenue/6YearAgo_Revenue, 1/6) - 1
        """
        if ttm_data.count() >= 24:  # Approximate 6 years of quarterly data
            current = ttm_data.first()
            six_years_ago = ttm_data[23]  # Roughly 6 years

            if current.ttm_revenue and six_years_ago.ttm_revenue:
                return self.calculate_cagr(six_years_ago.ttm_revenue, current.ttm_revenue, 6)
        return 0.0

    def calculate_pat_6yr_cagr(self, ttm_data):
        """
        PAT 6Y CAGR Formula: POWER(Current_TTM_PAT/6YearAgo_PAT, 1/6) - 1
        """
        if ttm_data.count() >= 24:  # Approximate 6 years of quarterly data
            current = ttm_data.first()
            six_years_ago = ttm_data[23]  # Roughly 6 years

            if current.ttm_pat and six_years_ago.ttm_pat:
                return self.calculate_cagr(six_years_ago.ttm_pat, current.ttm_pat, 6)
        return 0.0

    def calculate_pe_metrics(self, market_cap_data, ttm_data):
        """
        Calculate all PE-related metrics
        """
        metrics = {
            'current_pe': 0.0,
            'pe_2yr_avg': 0.0,
            'pe_5yr_avg': 0.0,
            'pe_2yr_reval_deval': 0.0,
            'pe_5yr_reval_deval': 0.0
        }

        if not market_cap_data.exists() or not ttm_data.exists():
            return metrics

        # Calculate PE ratios for available periods
        pe_ratios = []
        for market_cap in market_cap_data[:20]:  # Last 5 years approx
            # Convert market cap date to YYYYMM format to match TTM period format
            market_cap_period = f"{market_cap.date.year}{market_cap.date.month:02d}"

            # Find closest TTM data for this period
            ttm = ttm_data.filter(period__lte=market_cap_period).first()
            if ttm and ttm.ttm_pat and ttm.ttm_pat != 0:
                pe_ratio = market_cap.market_cap / ttm.ttm_pat
                pe_ratios.append(pe_ratio)

        if pe_ratios:
            # Current PE
            metrics['current_pe'] = pe_ratios[0]

            # 2 Year Average PE (8 quarters)
            if len(pe_ratios) >= 8:
                metrics['pe_2yr_avg'] = sum(pe_ratios[:8]) / 8

            # 5 Year Average PE (20 quarters)
            if len(pe_ratios) >= 20:
                metrics['pe_5yr_avg'] = sum(pe_ratios[:20]) / 20

            # Revaluation/Devaluation calculations
            if metrics['pe_2yr_avg'] and metrics['current_pe']:
                metrics['pe_2yr_reval_deval'] = ((metrics['pe_2yr_avg'] - metrics['current_pe']) /
                                               metrics['current_pe'])

            if metrics['pe_5yr_avg'] and metrics['current_pe']:
                metrics['pe_5yr_reval_deval'] = ((metrics['pe_5yr_avg'] - metrics['current_pe']) /
                                               metrics['current_pe'])

        return metrics

    def calculate_pr_metrics(self, market_cap_data, ttm_data):
        """
        Calculate all PR (Price to Revenue) related metrics
        """
        metrics = {
            'current_pr': 0.0,
            'pr_2yr_avg': 0.0,
            'pr_5yr_avg': 0.0,
            'pr_2yr_reval_deval': 0.0,
            'pr_5yr_reval_deval': 0.0,
            'pr_10q_low': 0.0,
            'pr_10q_high': 0.0
        }

        if not market_cap_data.exists() or not ttm_data.exists():
            return metrics

        # Calculate PR ratios for available periods
        pr_ratios = []
        for market_cap in market_cap_data[:20]:  # Last 5 years approx
            # Convert market cap date to YYYYMM format to match TTM period format
            market_cap_period = f"{market_cap.date.year}{market_cap.date.month:02d}"

            # Find closest TTM data for this period
            ttm = ttm_data.filter(period__lte=market_cap_period).first()
            if ttm and ttm.ttm_revenue and ttm.ttm_revenue != 0:
                pr_ratio = market_cap.market_cap / ttm.ttm_revenue
                pr_ratios.append(pr_ratio)

        if pr_ratios:
            # Current PR
            metrics['current_pr'] = pr_ratios[0]

            # 2 Year Average PR (8 quarters)
            if len(pr_ratios) >= 8:
                metrics['pr_2yr_avg'] = sum(pr_ratios[:8]) / 8

            # 5 Year Average PR (20 quarters)
            if len(pr_ratios) >= 20:
                metrics['pr_5yr_avg'] = sum(pr_ratios[:20]) / 20

            # 10 Quarter High/Low
            if len(pr_ratios) >= 10:
                ten_q_ratios = pr_ratios[:10]
                metrics['pr_10q_low'] = min(ten_q_ratios)
                metrics['pr_10q_high'] = max(ten_q_ratios)

            # Revaluation/Devaluation calculations
            if metrics['pr_2yr_avg'] and metrics['current_pr']:
                metrics['pr_2yr_reval_deval'] = ((metrics['pr_2yr_avg'] - metrics['current_pr']) /
                                               metrics['current_pr'])

            if metrics['pr_5yr_avg'] and metrics['current_pr']:
                metrics['pr_5yr_reval_deval'] = ((metrics['pr_5yr_avg'] - metrics['current_pr']) /
                                               metrics['current_pr'])

        return metrics

    def calculate_alpha_and_other_metrics(self, stock, period, ttm_data, market_cap_data):
        """
        Calculate Alpha, PE Yield, Growth Rate, and Bond Rate
        """
        metrics = {
            'alpha_bond_cagr': 0.0,
            'alpha_absolute': 0.0,
            'pe_yield': 0.0,
            'growth_rate': 0.0,
            'bond_rate': 0.0  # This would typically come from external data
        }

        # PE Yield calculation: 1/PE * 100
        if market_cap_data.exists() and ttm_data.exists():
            latest_market_cap = market_cap_data.first()
            latest_ttm = ttm_data.first()

            if (latest_market_cap.market_cap and latest_ttm.ttm_pat and
                latest_ttm.ttm_pat != 0):
                pe_ratio = latest_market_cap.market_cap / latest_ttm.ttm_pat
                if pe_ratio != 0:
                    metrics['pe_yield'] = (1 / pe_ratio) * 100

        # Growth Rate: Combined revenue and PAT growth
        if ttm_data.count() >= 2:
            current = ttm_data[0]
            previous = ttm_data[1]

            revenue_growth = 0.0
            pat_growth = 0.0

            if (current.ttm_revenue and previous.ttm_revenue and
                previous.ttm_revenue != 0):
                revenue_growth = ((current.ttm_revenue - previous.ttm_revenue) /
                                previous.ttm_revenue)

            if (current.ttm_pat and previous.ttm_pat and
                previous.ttm_pat != 0):
                pat_growth = ((current.ttm_pat - previous.ttm_pat) /
                            previous.ttm_pat)

            # Average of revenue and PAT growth
            metrics['growth_rate'] = (revenue_growth + pat_growth) / 2

        # Bond rate would typically be fetched from external source
        # For now, using a default rate (this should be configurable)
        metrics['bond_rate'] = 0.06  # 6% default

        # Alpha calculations would require portfolio-level returns
        # These are placeholders and would need proper implementation
        metrics['alpha_bond_cagr'] = 0.0
        metrics['alpha_absolute'] = 0.0

        return metrics

    def update_fund_latest_metrics(self, scheme, latest_metrics):
        """
        Update AMCFundScheme with aggregated portfolio-level metrics
        """

        update_fields = {}

        # Calculate weighted averages for portfolio-level metrics
        for metric_name, values in latest_metrics.items():
            if values:  # Only if we have values
                avg_value = sum(values) / len(values)
                field_name = f"latest_{metric_name}"
                update_fields[field_name] = avg_value

        # Add timestamp
        update_fields['metrics_last_updated'] = timezone.now()

        # Update the scheme
        for field, value in update_fields.items():
            setattr(scheme, field, value)

        scheme.save()
        logger.info(f"Updated latest metrics for fund: {scheme.name}")

    def update_progress(self, **kwargs):
        """Update progress session with new values"""
        if self.progress_session:
            for key, value in kwargs.items():
                if hasattr(self.progress_session, key):
                    setattr(self.progress_session, key, value)
            self.progress_session.save()


# Utility functions for dynamic header generation
class DynamicHeaderGenerator:
    """
    Generate dynamic headers based on available periods instead of fixed column structure
    """

    @staticmethod
    def get_available_periods_for_fund(scheme):
        """
        Get all available periods across all stocks in a fund
        """

        holdings = FundHolding.objects.filter(scheme=scheme).select_related('stock')
        all_periods = {
            'market_cap': set(),
            'ttm': set(),
            'quarterly': set(),
            'all': set()
        }

        for holding in holdings:
            stock = holding.stock

            # Market cap periods
            market_cap_dates = StockMarketCap.objects.filter(stock=stock).values_list('date', flat=True)
            all_periods['market_cap'].update(market_cap_dates)
            all_periods['all'].update(market_cap_dates)

            # TTM periods
            try:
                ttm_dates = StockTTMData.objects.filter(stock=stock).values_list('period', flat=True)
                all_periods['ttm'].update(ttm_dates)
                all_periods['all'].update(ttm_dates)
            except Exception as e:
                logger.error(f"Error getting TTM periods for stock {stock}: {e}")

            # Quarterly periods
            try:
                quarterly_dates = StockQuarterlyData.objects.filter(stock=stock).values_list('period', flat=True)
                all_periods['quarterly'].update(quarterly_dates)
                all_periods['all'].update(quarterly_dates)
            except Exception as e:
                logger.error(f"Error getting Quarterly periods for stock {stock}: {e}")

        # Convert to sorted lists (latest first)
        for key in all_periods:
            # Handle different data types by sorting separately
            periods_list = list(all_periods[key])
            if periods_list:
                # Check if we have mixed types
                if key == 'all':
                    # For 'all', we need to handle mixed types differently
                    # Separate dates and strings
                    dates = [p for p in periods_list if hasattr(p, 'year')]
                    strings = [p for p in periods_list if isinstance(p, str)]

                    # Sort each type separately
                    dates_sorted = sorted(dates, reverse=True)
                    strings_sorted = sorted(strings, reverse=True)

                    # Combine - dates first, then strings
                    all_periods[key] = dates_sorted + strings_sorted
                else:
                    # For specific types, all items should be the same type
                    all_periods[key] = sorted(periods_list, reverse=True)
            else:
                all_periods[key] = []

        return all_periods

    @staticmethod
    def generate_dynamic_headers(available_periods):
        """
        Generate headers based on actual available periods (not fixed 431)
        """

        # Base headers (stock info + fund columns)
        headers = {
            'row_1': ['', '', '', '', '', '', '', '', ''] + ['Market Cap'] * len(available_periods['market_cap']),
            'row_2': ['', '', '', '', '', '', '', '', ''] + list(available_periods['market_cap']),
            'row_3': ['Company Name', 'Accord Code', 'Sector', 'Cap', 'Market Cap', 'Weights', 'Factor', 'Value', 'No.of shares'],
            'row_4': ['', '', '', '', '', '', '', '', ''],
            'row_5': ['', '', '', '', '', '', '', '', ''],
            'row_6': ['', '', '', '', '', '', '', '', ''],
            'row_7': ['', '', '', '', '', '', '', '', '']
        }

        # Add dynamic columns for each period
        start_col = 9  # After base columns

        # Market cap columns
        for period in available_periods['market_cap']:
            headers['row_3'].append(f"Market Cap {period}")
            for row in ['row_1', 'row_2', 'row_4', 'row_5', 'row_6', 'row_7']:
                if row not in ['row_1', 'row_2']:  # These are handled above
                    headers[row].append('')

        # TTM columns
        for period in available_periods['ttm']:
            headers['row_3'].extend([f"TTM Revenue {period}", f"TTM PAT {period}"])
            for row in ['row_1', 'row_2', 'row_4', 'row_5', 'row_6', 'row_7']:
                headers[row].extend(['', ''])

        # Quarterly columns
        for period in available_periods['quarterly']:
            headers['row_3'].extend([f"Q Revenue {period}", f"Q PAT {period}"])
            for row in ['row_1', 'row_2', 'row_4', 'row_5', 'row_6', 'row_7']:
                headers[row].extend(['', ''])

        return headers