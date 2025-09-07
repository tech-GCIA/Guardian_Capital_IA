# mf_metrics_calculator.py
"""
Comprehensive Mutual Fund Metrics Calculator
Based on Portfolio Analysis functionality from Old Bridge Focused Equity Fund.xlsm

This module calculates portfolio-level metrics by analyzing:
1. AMCFundScheme (mutual fund)
2. SchemeUnderlyingHoldings (portfolio composition) 
3. Stock (individual stock details)
4. StockQuarterlyData (quarterly financial data)

Key Calculations:
- TOTALS: Portfolio market cap, free float mcap
- PATM: Profit After Tax metrics
- QoQ/YoY Growth rates
- 6-year CAGR calculations
- PE Ratios (current, 2yr avg, 5yr avg)
- P/R Ratios and Reval/Deval
- Alpha, Beta, ROE, ROCE metrics
"""

import logging
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, timedelta
from django.db.models import Q, Sum, Avg, Count, Max
from django.utils import timezone
from typing import Dict, List, Tuple, Optional

from gcia_app.models import (
    AMCFundScheme, SchemeUnderlyingHoldings, Stock, StockQuarterlyData,
    MutualFundMetrics, MetricsCalculationLog
)

logger = logging.getLogger(__name__)
data_quality_logger = logging.getLogger('gcia_app.data_quality')


class MFMetricsCalculator:
    """
    Main calculator class for mutual fund portfolio metrics
    """
    
    def __init__(self):
        self.calculation_log = None
        self.current_user = None
    
    def calculate_all_fund_metrics(self, user, force_recalculate=False):
        """
        Calculate metrics for all active mutual funds
        
        Args:
            user: User initiating the calculation
            force_recalculate: Whether to recalculate existing metrics
        
        Returns:
            MetricsCalculationLog: Log object with results
        """
        # Create calculation log
        self.current_user = user
        self.calculation_log = MetricsCalculationLog.objects.create(
            initiated_by=user,
            calculation_type='bulk_all',
            status='running'
        )
        
        logger.info(f"Starting bulk metrics calculation initiated by {user.username}")
        
        try:
            # Get all active funds with holdings
            active_funds = AMCFundScheme.objects.filter(
                is_active=True
            ).prefetch_related(
                'schemeunderlyingholdings_set__holding__quarterly_data'
            ).order_by('name')
            
            self.calculation_log.total_funds_targeted = active_funds.count()
            self.calculation_log.save()
            
            logger.info(f"Found {active_funds.count()} active funds to process")
            
            # Process each fund
            successful_count = 0
            partial_count = 0
            failed_count = 0
            
            for fund in active_funds:
                try:
                    result = self.calculate_single_fund_metrics(fund, force_recalculate)
                    
                    if result['status'] == 'success':
                        successful_count += 1
                    elif result['status'] == 'partial':
                        partial_count += 1
                    else:
                        failed_count += 1
                        
                    logger.info(f"Processed {fund.name}: {result['status']}")
                    
                except Exception as e:
                    failed_count += 1
                    logger.error(f"Error processing fund {fund.name}: {str(e)}")
            
            # Update log with final results
            self.calculation_log.funds_processed_successfully = successful_count
            self.calculation_log.funds_with_partial_data = partial_count
            self.calculation_log.funds_failed = failed_count
            self.calculation_log.completed_at = timezone.now()
            self.calculation_log.status = 'completed'
            self.calculation_log.save()
            
            logger.info(f"Bulk calculation completed: {successful_count} successful, {partial_count} partial, {failed_count} failed")
            
            return self.calculation_log
            
        except Exception as e:
            # Handle overall failure
            self.calculation_log.status = 'failed'
            self.calculation_log.error_summary = str(e)
            self.calculation_log.completed_at = timezone.now()
            self.calculation_log.save()
            logger.error(f"Bulk calculation failed: {str(e)}")
            raise
    
    def calculate_single_fund_metrics(self, fund: AMCFundScheme, force_recalculate=False) -> Dict:
        """
        Calculate metrics for a single mutual fund
        
        Args:
            fund: AMCFundScheme instance
            force_recalculate: Whether to recalculate existing metrics
            
        Returns:
            Dict with status and calculated metrics
        """
        logger.info(f"Calculating metrics for fund: {fund.name}")
        
        # Check if metrics already exist and not forcing recalculation
        if not force_recalculate:
            existing_metrics = MutualFundMetrics.objects.filter(
                amcfundscheme=fund
            ).order_by('-calculation_date').first()
            
            if existing_metrics and existing_metrics.calculation_date.date() == timezone.now().date():
                logger.info(f"Metrics already calculated today for {fund.name}, skipping")
                return {'status': 'skipped', 'message': 'Already calculated today'}
        
        try:
            # Get portfolio holdings with stock data
            holdings_data = self.get_portfolio_holdings_data(fund)
            
            if not holdings_data:
                logger.warning(f"No holdings data found for fund {fund.name}")
                data_quality_logger.error(f"MF_METRICS_FAILED: {fund.name} (ID: {fund.amcfundscheme_id}) - No holdings data found")
                data_quality_logger.info(f"Fund Status Check - ID: {fund.amcfundscheme_id}, Name: {fund.name}, Active: {fund.is_active}")
                
                # Additional debugging for data quality
                holdings_count = SchemeUnderlyingHoldings.objects.filter(amcfundscheme=fund).count()
                active_holdings_count = SchemeUnderlyingHoldings.objects.filter(amcfundscheme=fund, is_active=True).count()
                data_quality_logger.info(f"Holdings Debug - Total holdings: {holdings_count}, Active holdings: {active_holdings_count}")
                
                return {'status': 'failed', 'message': 'No holdings data available'}
            
            # Assess data quality first to determine what can be calculated
            data_quality = self.assess_data_quality(holdings_data)
            
            # Calculate portfolio metrics based on available data
            calculated_metrics = self.calculate_portfolio_metrics(holdings_data, data_quality)
            
            # Determine calculation status based on data quality and successful calculations
            metrics_calculated = sum(1 for v in calculated_metrics.values() if v is not None and v != 0)
            
            if data_quality['completeness_pct'] > 70 and metrics_calculated >= 5:
                calc_status = 'success'
            elif data_quality['completeness_pct'] > 30 and metrics_calculated >= 3:
                calc_status = 'partial'
            elif calculated_metrics.get('portfolio_market_cap', 0) > 0:
                # If we have at least market cap data, it's partial
                calc_status = 'partial'
            else:
                calc_status = 'failed'
            
            # Create or update metrics record
            metrics_obj, created = MutualFundMetrics.objects.update_or_create(
                amcfundscheme=fund,
                defaults={
                    'total_holdings': len(holdings_data),
                    'total_weightage': sum(h['weightage'] for h in holdings_data),
                    'data_as_of_date': data_quality['latest_quarter_date'],
                    'calculation_status': calc_status,
                    'calculation_notes': data_quality.get('notes', ''),
                    **calculated_metrics
                }
            )
            
            logger.info(f"{'Created' if created else 'Updated'} metrics for {fund.name} with {calc_status} status")
            
            return {
                'status': calc_status,
                'metrics_obj': metrics_obj,
                'data_quality': data_quality,
                'calculated_metrics': calculated_metrics
            }
            
        except Exception as e:
            logger.error(f"Error calculating metrics for fund {fund.name}: {str(e)}")
            return {'status': 'failed', 'message': str(e)}
    
    def get_portfolio_holdings_data(self, fund: AMCFundScheme) -> List[Dict]:
        """
        Get comprehensive holdings data for a fund including quarterly data
        
        Args:
            fund: AMCFundScheme instance
            
        Returns:
            List of holdings with stock and quarterly data
        """
        holdings = SchemeUnderlyingHoldings.objects.filter(
            amcfundscheme=fund,
            is_active=True
        ).select_related(
            'holding'
        ).prefetch_related(
            'holding__quarterly_data'
        ).order_by('-weightage')
        
        holdings_data = []
        
        for holding in holdings:
            if not holding.holding:
                continue
                
            stock = holding.holding
            
            # Get latest quarterly data (within last 2 years)
            two_years_ago = timezone.now().date() - timedelta(days=730)
            quarterly_data = stock.quarterly_data.filter(
                quarter_date__gte=two_years_ago
            ).order_by('-quarter_date')
            
            holding_info = {
                'holding_id': holding.schemeunderlyingholding_id,
                'stock_id': stock.stock_id,
                'stock_name': stock.name,
                'stock_symbol': stock.symbol,
                'sector': stock.sector,
                'weightage': float(holding.weightage or 0),
                'shares': float(holding.no_of_shares or 0),
                'quarterly_data': []
            }
            
            # Add quarterly data with corrected field mapping
            for qdata in quarterly_data[:20]:  # Last 20 quarters (5 years)
                quarterly_info = {
                    'quarter_date': qdata.quarter_date,
                    'quarter_year': qdata.quarter_year,
                    'quarter_number': qdata.quarter_number,
                    'mcap': float(qdata.mcap) if qdata.mcap else None,
                    'free_float_mcap': float(qdata.free_float_mcap) if qdata.free_float_mcap else None,
                    # Fix PAT field mapping - use multiple sources
                    'pat': float(qdata.quarterly_pat) if qdata.quarterly_pat else (float(qdata.pat) if qdata.pat else None),
                    'ttm_pat': float(qdata.ttm) if qdata.ttm else None,
                    'revenue': float(qdata.quarterly_revenue) if qdata.quarterly_revenue else (float(qdata.ttm_revenue) if qdata.ttm_revenue else None),
                    'pe_ratio': float(qdata.pe_quarterly) if qdata.pe_quarterly else None,
                    'pr_ratio': float(qdata.pr_quarterly) if qdata.pr_quarterly else None,
                    'roe': float(qdata.roe) if qdata.roe else None,
                    'roce': float(qdata.roce) if qdata.roce else None,
                }
                holding_info['quarterly_data'].append(quarterly_info)
            
            holdings_data.append(holding_info)
        
        return holdings_data
    
    def calculate_portfolio_metrics(self, holdings_data: List[Dict], data_quality: Dict = None) -> Dict:
        """
        Calculate comprehensive portfolio metrics based on holdings and available data
        
        Args:
            holdings_data: List of holdings with quarterly data
            data_quality: Optional data quality assessment to determine which metrics to calculate
            
        Returns:
            Dict with calculated portfolio metrics
        """
        if not holdings_data:
            return {}
        
        total_weightage = sum(h['weightage'] for h in holdings_data)
        if total_weightage == 0:
            logger.warning("Total portfolio weightage is 0, cannot calculate weighted metrics")
            return {}
        
        # Initialize metrics
        metrics = {}
        
        # Always calculate basic metrics (market cap)
        latest_metrics = self.calculate_weighted_latest_quarter_metrics(holdings_data, total_weightage)
        metrics.update(latest_metrics)
        
        # Only calculate advanced metrics if we have sufficient data quality
        if data_quality and data_quality.get('metrics_available'):
            available = data_quality['metrics_available']
            
            # Calculate growth metrics if PAT data is available
            if available.get('growth_metrics', False):
                growth_metrics = self.calculate_growth_metrics(holdings_data, total_weightage)
                metrics.update(growth_metrics)
            else:
                logger.info("Skipping growth metrics calculation - insufficient PAT data")
            
            # Calculate CAGR metrics if revenue/PAT data is available
            if available.get('pat', False) or available.get('revenue', False):
                cagr_metrics = self.calculate_cagr_metrics(holdings_data, total_weightage)
                metrics.update(cagr_metrics)
            else:
                logger.info("Skipping CAGR metrics calculation - insufficient revenue/PAT data")
            
            # Calculate historical averages if ratio data is available
            if available.get('pe_ratio', False) or available.get('pr_ratio', False):
                avg_metrics = self.calculate_historical_averages(holdings_data, total_weightage)
                metrics.update(avg_metrics)
            else:
                logger.info("Skipping historical averages calculation - insufficient ratio data")
        else:
            # Fallback - try to calculate all metrics (legacy behavior)
            logger.info("Data quality not assessed - attempting all metric calculations")
            
            growth_metrics = self.calculate_growth_metrics(holdings_data, total_weightage)
            metrics.update(growth_metrics)
            
            cagr_metrics = self.calculate_cagr_metrics(holdings_data, total_weightage)
            metrics.update(cagr_metrics)
            
            avg_metrics = self.calculate_historical_averages(holdings_data, total_weightage)
            metrics.update(avg_metrics)
        
        # Performance metrics are placeholder for now (require benchmark data)
        performance_metrics = self.calculate_performance_metrics(holdings_data, total_weightage)
        metrics.update(performance_metrics)
        
        return metrics
    
    def calculate_weighted_latest_quarter_metrics(self, holdings_data: List[Dict], total_weightage: float) -> Dict:
        """Calculate weighted metrics for latest available quarter"""
        metrics = {}
        
        weighted_mcap = 0
        weighted_free_float_mcap = 0
        weighted_pat = 0
        weighted_ttm_pat = 0
        weighted_pe = 0
        weighted_pr = 0
        weighted_roe = 0
        weighted_roce = 0
        
        valid_pe_weightage = 0
        valid_pr_weightage = 0
        valid_roe_weightage = 0
        valid_roce_weightage = 0
        
        for holding in holdings_data:
            weight = holding['weightage'] / total_weightage if total_weightage > 0 else 0
            
            if holding['quarterly_data']:
                latest_quarter = holding['quarterly_data'][0]
                
                # Market cap metrics
                if latest_quarter['mcap']:
                    weighted_mcap += latest_quarter['mcap'] * weight
                if latest_quarter['free_float_mcap']:
                    weighted_free_float_mcap += latest_quarter['free_float_mcap'] * weight
                
                # PAT metrics
                if latest_quarter['pat']:
                    weighted_pat += latest_quarter['pat'] * weight
                if latest_quarter['ttm_pat']:
                    weighted_ttm_pat += latest_quarter['ttm_pat'] * weight
                
                # Ratio metrics (only include if valid) - improved handling
                if latest_quarter.get('pe_ratio') and latest_quarter['pe_ratio'] > 0:
                    weighted_pe += latest_quarter['pe_ratio'] * weight
                    valid_pe_weightage += weight
                
                if latest_quarter.get('pr_ratio') and latest_quarter['pr_ratio'] > 0:
                    weighted_pr += latest_quarter['pr_ratio'] * weight
                    valid_pr_weightage += weight
                
                if latest_quarter.get('roe') and latest_quarter['roe'] != 0:
                    weighted_roe += latest_quarter['roe'] * weight
                    valid_roe_weightage += weight
                    
                if latest_quarter.get('roce') and latest_quarter['roce'] != 0:
                    weighted_roce += latest_quarter['roce'] * weight
                    valid_roce_weightage += weight
        
        # Store calculated metrics
        metrics['portfolio_market_cap'] = Decimal(str(weighted_mcap)).quantize(Decimal('0.01'))
        metrics['portfolio_free_float_mcap'] = Decimal(str(weighted_free_float_mcap)).quantize(Decimal('0.01'))
        metrics['portfolio_pat'] = Decimal(str(weighted_pat)).quantize(Decimal('0.01'))
        metrics['portfolio_ttm_pat'] = Decimal(str(weighted_ttm_pat)).quantize(Decimal('0.01'))
        
        # Normalize ratios by valid weightage
        if valid_pe_weightage > 0:
            metrics['portfolio_current_pe'] = Decimal(str(weighted_pe / valid_pe_weightage)).quantize(Decimal('0.0001'))
        
        if valid_pr_weightage > 0:
            metrics['portfolio_current_pr'] = Decimal(str(weighted_pr / valid_pr_weightage)).quantize(Decimal('0.0001'))
        
        if valid_roe_weightage > 0:
            metrics['portfolio_roe'] = Decimal(str(weighted_roe / valid_roe_weightage)).quantize(Decimal('0.0001'))
            
        if valid_roce_weightage > 0:
            metrics['portfolio_roce'] = Decimal(str(weighted_roce / valid_roce_weightage)).quantize(Decimal('0.0001'))
        
        return metrics
    
    def calculate_growth_metrics(self, holdings_data: List[Dict], total_weightage: float) -> Dict:
        """Calculate QoQ and YoY growth metrics with improved error handling"""
        metrics = {}
        
        weighted_qoq_growth = 0
        weighted_yoy_growth = 0
        valid_qoq_weightage = 0
        valid_yoy_weightage = 0
        
        qoq_calculations = 0
        yoy_calculations = 0
        
        for holding in holdings_data:
            try:
                weight = holding['weightage'] / total_weightage if total_weightage > 0 else 0
                quarterly_data = holding.get('quarterly_data', [])
                
                if len(quarterly_data) >= 2:
                    # QoQ Growth (current vs previous quarter)
                    current_pat = quarterly_data[0].get('pat')
                    prev_quarter_pat = quarterly_data[1].get('pat')
                    
                    if (current_pat is not None and prev_quarter_pat is not None and 
                        current_pat > 0 and prev_quarter_pat > 0 and 
                        abs(prev_quarter_pat) > 0.01):  # Avoid division by very small numbers
                        
                        qoq_growth = ((current_pat - prev_quarter_pat) / prev_quarter_pat) * 100
                        
                        # Sanity check - exclude extreme growth rates (likely data errors)
                        if -1000 <= qoq_growth <= 1000:
                            weighted_qoq_growth += qoq_growth * weight
                            valid_qoq_weightage += weight
                            qoq_calculations += 1
                
                if len(quarterly_data) >= 4:
                    # YoY Growth (current vs same quarter last year)
                    current_pat = quarterly_data[0].get('pat')
                    yoy_pat = quarterly_data[3].get('pat')  # 4 quarters = 1 year
                    
                    if (current_pat is not None and yoy_pat is not None and 
                        current_pat > 0 and yoy_pat > 0 and 
                        abs(yoy_pat) > 0.01):  # Avoid division by very small numbers
                        
                        yoy_growth = ((current_pat - yoy_pat) / yoy_pat) * 100
                        
                        # Sanity check - exclude extreme growth rates
                        if -1000 <= yoy_growth <= 1000:
                            weighted_yoy_growth += yoy_growth * weight
                            valid_yoy_weightage += weight
                            yoy_calculations += 1
                            
            except Exception as e:
                logger.warning(f"Error calculating growth metrics for holding {holding.get('stock_name', 'unknown')}: {e}")
                continue
        
        # Only set metrics if we have meaningful calculations
        if valid_qoq_weightage > 0.01 and qoq_calculations >= 3:  # At least 3 holdings with valid data
            metrics['portfolio_qoq_growth'] = Decimal(str(weighted_qoq_growth / valid_qoq_weightage)).quantize(Decimal('0.0001'))
            logger.info(f"QoQ growth calculated from {qoq_calculations} holdings with {valid_qoq_weightage:.2%} weightage")
        else:
            logger.info(f"Insufficient data for QoQ growth: {qoq_calculations} calculations, {valid_qoq_weightage:.2%} weightage")
        
        if valid_yoy_weightage > 0.01 and yoy_calculations >= 3:  # At least 3 holdings with valid data
            metrics['portfolio_yoy_growth'] = Decimal(str(weighted_yoy_growth / valid_yoy_weightage)).quantize(Decimal('0.0001'))
            logger.info(f"YoY growth calculated from {yoy_calculations} holdings with {valid_yoy_weightage:.2%} weightage")
        else:
            logger.info(f"Insufficient data for YoY growth: {yoy_calculations} calculations, {valid_yoy_weightage:.2%} weightage")
        
        return metrics
    
    def calculate_cagr_metrics(self, holdings_data: List[Dict], total_weightage: float) -> Dict:
        """Calculate CAGR for revenue and PAT with improved error handling"""
        metrics = {}
        
        weighted_revenue_cagr = 0
        weighted_pat_cagr = 0
        valid_revenue_weightage = 0
        valid_pat_weightage = 0
        
        revenue_calculations = 0
        pat_calculations = 0
        
        for holding in holdings_data:
            try:
                weight = holding['weightage'] / total_weightage if total_weightage > 0 else 0
                quarterly_data = holding.get('quarterly_data', [])
                
                # Use available data to estimate CAGR (minimum 2 years = 8 quarters)
                if len(quarterly_data) >= 8:
                    recent_data = quarterly_data[:4]  # Last 4 quarters
                    older_data = quarterly_data[-4:]  # Oldest 4 quarters available
                    
                    # Calculate average revenue for recent vs older periods with better error handling
                    recent_revenues = [q.get('revenue') for q in recent_data if q.get('revenue') and q.get('revenue') > 0]
                    older_revenues = [q.get('revenue') for q in older_data if q.get('revenue') and q.get('revenue') > 0]
                    
                    if len(recent_revenues) >= 2 and len(older_revenues) >= 2:  # Need at least 2 quarters each
                        recent_revenue = sum(recent_revenues) / len(recent_revenues)
                        older_revenue = sum(older_revenues) / len(older_revenues)
                        
                        if recent_revenue > 0.01 and older_revenue > 0.01:  # Avoid tiny numbers
                            years = len(quarterly_data) / 4  # Approximate years
                            if years > 0.5:  # At least 6 months of data
                                try:
                                    revenue_cagr = (pow(recent_revenue / older_revenue, 1/years) - 1) * 100
                                    
                                    # Sanity check - exclude extreme CAGR rates
                                    if -100 <= revenue_cagr <= 200:  # -100% to +200% seems reasonable
                                        weighted_revenue_cagr += revenue_cagr * weight
                                        valid_revenue_weightage += weight
                                        revenue_calculations += 1
                                except (ValueError, OverflowError, ZeroDivisionError) as e:
                                    logger.debug(f"CAGR calculation error for revenue: {e}")
                    
                    # Similar calculation for PAT with better error handling
                    recent_pats = [q.get('pat') for q in recent_data if q.get('pat') and q.get('pat') > 0]
                    older_pats = [q.get('pat') for q in older_data if q.get('pat') and q.get('pat') > 0]
                    
                    if len(recent_pats) >= 2 and len(older_pats) >= 2:  # Need at least 2 quarters each
                        recent_pat = sum(recent_pats) / len(recent_pats)
                        older_pat = sum(older_pats) / len(older_pats)
                        
                        if recent_pat > 0.01 and older_pat > 0.01:  # Avoid tiny numbers
                            years = len(quarterly_data) / 4
                            if years > 0.5:  # At least 6 months of data
                                try:
                                    pat_cagr = (pow(recent_pat / older_pat, 1/years) - 1) * 100
                                    
                                    # Sanity check - exclude extreme CAGR rates
                                    if -100 <= pat_cagr <= 200:  # -100% to +200% seems reasonable
                                        weighted_pat_cagr += pat_cagr * weight
                                        valid_pat_weightage += weight
                                        pat_calculations += 1
                                except (ValueError, OverflowError, ZeroDivisionError) as e:
                                    logger.debug(f"CAGR calculation error for PAT: {e}")
                                    
            except Exception as e:
                logger.warning(f"Error calculating CAGR metrics for holding {holding.get('stock_name', 'unknown')}: {e}")
                continue
        
        # Only set metrics if we have meaningful calculations
        if valid_revenue_weightage > 0.01 and revenue_calculations >= 3:
            metrics['portfolio_revenue_cagr'] = Decimal(str(weighted_revenue_cagr / valid_revenue_weightage)).quantize(Decimal('0.0001'))
            logger.info(f"Revenue CAGR calculated from {revenue_calculations} holdings with {valid_revenue_weightage:.2%} weightage")
        else:
            logger.info(f"Insufficient data for revenue CAGR: {revenue_calculations} calculations, {valid_revenue_weightage:.2%} weightage")
        
        if valid_pat_weightage > 0.01 and pat_calculations >= 3:
            metrics['portfolio_pat_cagr'] = Decimal(str(weighted_pat_cagr / valid_pat_weightage)).quantize(Decimal('0.0001'))
            logger.info(f"PAT CAGR calculated from {pat_calculations} holdings with {valid_pat_weightage:.2%} weightage")
        else:
            logger.info(f"Insufficient data for PAT CAGR: {pat_calculations} calculations, {valid_pat_weightage:.2%} weightage")
        
        return metrics
    
    def calculate_historical_averages(self, holdings_data: List[Dict], total_weightage: float) -> Dict:
        """Calculate multi-period historical averages with enhanced data support"""
        metrics = {}
        
        # Initialize counters for different time periods
        weighted_2yr_pe = weighted_2yr_pr = weighted_5yr_pe = 0
        valid_2yr_pe_weightage = valid_2yr_pr_weightage = valid_5yr_pe_weightage = 0
        
        pe_2yr_calculations = pr_2yr_calculations = pe_5yr_calculations = 0
        
        for holding in holdings_data:
            try:
                weight = holding['weightage'] / total_weightage if total_weightage > 0 else 0
                quarterly_data = holding.get('quarterly_data', [])
                
                # Enhanced 2-year averages (8 quarters) - more flexible data requirements
                if len(quarterly_data) >= 4:  # Lowered requirement from 8 to 4 quarters minimum
                    # Look at up to 8 quarters for 2-year average
                    data_2yr = quarterly_data[:min(8, len(quarterly_data))]
                    
                    # PE values with better filtering
                    pe_values_2yr = []
                    for q in data_2yr:
                        pe = q.get('pe_ratio')
                        if pe is not None and 0.1 <= pe <= 500:  # Reasonable PE range
                            pe_values_2yr.append(pe)
                    
                    if len(pe_values_2yr) >= 2:  # Need at least 2 valid quarters
                        avg_pe_2yr = sum(pe_values_2yr) / len(pe_values_2yr)
                        weighted_2yr_pe += avg_pe_2yr * weight
                        valid_2yr_pe_weightage += weight
                        pe_2yr_calculations += 1
                    
                    # PR values with better filtering
                    pr_values_2yr = []
                    for q in data_2yr:
                        pr = q.get('pr_ratio')
                        if pr is not None and 0.01 <= pr <= 100:  # Reasonable P/R range
                            pr_values_2yr.append(pr)
                    
                    if len(pr_values_2yr) >= 2:  # Need at least 2 valid quarters
                        avg_pr_2yr = sum(pr_values_2yr) / len(pr_values_2yr)
                        weighted_2yr_pr += avg_pr_2yr * weight
                        valid_2yr_pr_weightage += weight
                        pr_2yr_calculations += 1
                
                # Enhanced 5-year averages - flexible data requirements
                if len(quarterly_data) >= 8:  # Lowered requirement from 20 to 8 quarters
                    # Look at all available data up to 20 quarters (5 years)
                    data_5yr = quarterly_data[:min(20, len(quarterly_data))]
                    
                    # PE values with better filtering
                    pe_values_5yr = []
                    for q in data_5yr:
                        pe = q.get('pe_ratio')
                        if pe is not None and 0.1 <= pe <= 500:  # Reasonable PE range
                            pe_values_5yr.append(pe)
                    
                    if len(pe_values_5yr) >= 4:  # Need at least 4 valid quarters
                        avg_pe_5yr = sum(pe_values_5yr) / len(pe_values_5yr)
                        weighted_5yr_pe += avg_pe_5yr * weight
                        valid_5yr_pe_weightage += weight
                        pe_5yr_calculations += 1
                        
            except Exception as e:
                logger.warning(f"Error calculating historical averages for holding {holding.get('stock_name', 'unknown')}: {e}")
                continue
        
        # Store calculated averages with meaningful thresholds
        if valid_2yr_pe_weightage > 0.01 and pe_2yr_calculations >= 3:
            metrics['portfolio_2yr_avg_pe'] = Decimal(str(weighted_2yr_pe / valid_2yr_pe_weightage)).quantize(Decimal('0.0001'))
            logger.info(f"2-year PE average calculated from {pe_2yr_calculations} holdings with {valid_2yr_pe_weightage:.2%} weightage")
        else:
            logger.info(f"Insufficient data for 2-year PE average: {pe_2yr_calculations} calculations, {valid_2yr_pe_weightage:.2%} weightage")
        
        if valid_2yr_pr_weightage > 0.01 and pr_2yr_calculations >= 3:
            metrics['portfolio_2yr_avg_pr'] = Decimal(str(weighted_2yr_pr / valid_2yr_pr_weightage)).quantize(Decimal('0.0001'))
            logger.info(f"2-year P/R average calculated from {pr_2yr_calculations} holdings with {valid_2yr_pr_weightage:.2%} weightage")
        else:
            logger.info(f"Insufficient data for 2-year P/R average: {pr_2yr_calculations} calculations, {valid_2yr_pr_weightage:.2%} weightage")
        
        if valid_5yr_pe_weightage > 0.01 and pe_5yr_calculations >= 3:
            metrics['portfolio_5yr_avg_pe'] = Decimal(str(weighted_5yr_pe / valid_5yr_pe_weightage)).quantize(Decimal('0.0001'))
            logger.info(f"5-year PE average calculated from {pe_5yr_calculations} holdings with {valid_5yr_pe_weightage:.2%} weightage")
        else:
            logger.info(f"Insufficient data for 5-year PE average: {pe_5yr_calculations} calculations, {valid_5yr_pe_weightage:.2%} weightage")
        
        # Enhanced Reval/Deval calculation with error handling
        try:
            current_pe = metrics.get('portfolio_current_pe')
            avg_pe_2yr = metrics.get('portfolio_2yr_avg_pe')
            
            if current_pe and avg_pe_2yr and float(avg_pe_2yr) > 0:
                current_pe_val = float(current_pe)
                avg_pe_2yr_val = float(avg_pe_2yr)
                
                reval_deval = ((current_pe_val - avg_pe_2yr_val) / avg_pe_2yr_val) * 100
                
                # Sanity check for reasonable revaluation range
                if -90 <= reval_deval <= 500:  # Allow -90% to +500% revaluation
                    metrics['portfolio_reval_deval'] = Decimal(str(reval_deval)).quantize(Decimal('0.0001'))
                    logger.info(f"Reval/Deval calculated: {reval_deval:.2f}% (Current PE: {current_pe_val:.2f}, 2yr Avg PE: {avg_pe_2yr_val:.2f})")
                else:
                    logger.warning(f"Extreme revaluation value {reval_deval:.2f}% excluded")
            else:
                logger.info("Cannot calculate Reval/Deval - missing current PE or 2-year average PE")
        except Exception as e:
            logger.warning(f"Error calculating Reval/Deval: {e}")
        
        return metrics
    
    def calculate_performance_metrics(self, holdings_data: List[Dict], total_weightage: float) -> Dict:
        """Calculate Alpha, Beta and other performance metrics"""
        metrics = {}
        
        # For Alpha and Beta calculations, we'd need benchmark data
        # This is a placeholder implementation
        # In reality, you'd compare portfolio returns against benchmark returns
        
        # Placeholder Alpha and Beta (would require benchmark data)
        metrics['portfolio_alpha'] = None  # Requires benchmark comparison
        metrics['portfolio_beta'] = None   # Requires benchmark comparison
        
        return metrics
    
    def assess_data_quality(self, holdings_data: List[Dict]) -> Dict:
        """Assess the quality and completeness of data for the portfolio"""
        total_holdings = len(holdings_data)
        if total_holdings == 0:
            return {
                'completeness_pct': 0, 
                'notes': 'No holdings data available',
                'metrics_available': {
                    'market_cap': False,
                    'pat': False,
                    'pe_ratio': False,
                    'pr_ratio': False,
                    'roe': False,
                    'roce': False,
                    'growth_metrics': False
                }
            }
        
        holdings_with_quarterly_data = 0
        latest_quarter_dates = []
        
        # Count holdings with specific types of data
        metrics_count = {
            'market_cap': 0,
            'pat': 0,
            'pe_ratio': 0,
            'pr_ratio': 0,
            'roe': 0,
            'roce': 0,
            'revenue': 0
        }
        
        total_weightage_with_data = 0
        
        for holding in holdings_data:
            holding_weight = holding['weightage'] or 0
            
            if holding['quarterly_data']:
                holdings_with_quarterly_data += 1
                latest_quarter_dates.append(holding['quarterly_data'][0]['quarter_date'])
                
                # Check latest quarter for specific metrics
                latest_quarter = holding['quarterly_data'][0]
                
                if latest_quarter.get('mcap'):
                    metrics_count['market_cap'] += 1
                    total_weightage_with_data += holding_weight
                    
                if latest_quarter.get('pat'):
                    metrics_count['pat'] += 1
                    
                if latest_quarter.get('pe_ratio') and latest_quarter['pe_ratio'] > 0:
                    metrics_count['pe_ratio'] += 1
                    
                if latest_quarter.get('pr_ratio') and latest_quarter['pr_ratio'] > 0:
                    metrics_count['pr_ratio'] += 1
                    
                if latest_quarter.get('roe') and latest_quarter['roe'] != 0:
                    metrics_count['roe'] += 1
                    
                if latest_quarter.get('roce') and latest_quarter['roce'] != 0:
                    metrics_count['roce'] += 1
                    
                if latest_quarter.get('revenue'):
                    metrics_count['revenue'] += 1
        
        completeness_pct = (holdings_with_quarterly_data / total_holdings) * 100
        
        # Calculate availability percentage for each metric type
        metrics_availability = {}
        for metric, count in metrics_count.items():
            metrics_availability[metric] = (count / total_holdings) * 100
        
        # Determine which metrics are available (>30% of holdings have the data)
        metrics_available = {
            'market_cap': metrics_availability['market_cap'] > 30,
            'pat': metrics_availability['pat'] > 20,  # Lower threshold for PAT
            'pe_ratio': metrics_availability['pe_ratio'] > 20,
            'pr_ratio': metrics_availability['pr_ratio'] > 20,
            'roe': metrics_availability['roe'] > 20,
            'roce': metrics_availability['roce'] > 20,
            'growth_metrics': metrics_availability['pat'] > 20 and holdings_with_quarterly_data >= 2  # Need 2+ quarters for growth
        }
        
        # Find most recent quarter date across all holdings
        latest_quarter_date = max(latest_quarter_dates) if latest_quarter_dates else None
        
        # Generate detailed notes
        notes = []
        if completeness_pct < 50:
            notes.append(f"Low data coverage: {holdings_with_quarterly_data}/{total_holdings} holdings have quarterly data")
        elif completeness_pct < 80:
            notes.append(f"Moderate data coverage: {holdings_with_quarterly_data}/{total_holdings} holdings have quarterly data")
        else:
            notes.append(f"Good data coverage: {holdings_with_quarterly_data}/{total_holdings} holdings have quarterly data")
        
        # Add specific metric availability notes
        available_metrics = [k for k, v in metrics_available.items() if v]
        if available_metrics:
            notes.append(f"Available metrics: {', '.join(available_metrics)}")
        
        missing_metrics = [k for k, v in metrics_available.items() if not v]
        if missing_metrics:
            notes.append(f"Limited/missing data for: {', '.join(missing_metrics)}")
        
        return {
            'completeness_pct': completeness_pct,
            'latest_quarter_date': latest_quarter_date,
            'holdings_with_data': holdings_with_quarterly_data,
            'total_holdings': total_holdings,
            'metrics_availability': metrics_availability,
            'metrics_available': metrics_available,
            'total_weightage_with_data': total_weightage_with_data,
            'notes': '; '.join(notes) if notes else 'Excellent data quality'
        }