#!/usr/bin/env python
"""
Test stock financial data quality to identify current issues
"""
import os
import sys
import django

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

from gcia_app.models import (
    AMCFundScheme, FundHolding, Stock, StockMarketCap,
    StockTTMData, StockQuarterlyData, StockAnnualRatios, StockPrice
)
from datetime import datetime

def test_stock_financial_data():
    print("=== TESTING STOCK FINANCIAL DATA QUALITY ===")
    print()

    # Get a sample fund for testing
    sample_fund = AMCFundScheme.objects.filter(is_active=True).first()
    print(f"Testing with fund: {sample_fund.name}")

    holdings = FundHolding.objects.filter(scheme=sample_fund)[:5]
    print(f"Testing with {holdings.count()} holdings")
    print()

    for holding in holdings:
        stock = holding.stock
        print(f"=== STOCK: {stock.company_name} ===")

        # Check each type of stock data
        market_cap_count = StockMarketCap.objects.filter(stock=stock).count()
        ttm_count = StockTTMData.objects.filter(stock=stock).count()
        quarterly_count = StockQuarterlyData.objects.filter(stock=stock).count()
        annual_count = StockAnnualRatios.objects.filter(stock=stock).count()
        price_count = StockPrice.objects.filter(stock=stock).count()

        print(f"  Market Cap records: {market_cap_count}")
        print(f"  TTM Data records: {ttm_count}")
        print(f"  Quarterly Data records: {quarterly_count}")
        print(f"  Annual Ratios records: {annual_count}")
        print(f"  Price Data records: {price_count}")

        # Test period data formats
        if ttm_count > 0:
            ttm_sample = StockTTMData.objects.filter(stock=stock).first()
            print(f"  Sample TTM period: {ttm_sample.period} (type: {type(ttm_sample.period)})")

        if quarterly_count > 0:
            quarterly_sample = StockQuarterlyData.objects.filter(stock=stock).first()
            print(f"  Sample Quarterly period: {quarterly_sample.period} (type: {type(quarterly_sample.period)})")

        # Test date conversion issues
        print(f"  Testing date conversions...")
        try:
            if ttm_count > 0:
                ttm_data = StockTTMData.objects.filter(stock=stock)
                for ttm in ttm_data[:2]:
                    period_date = ttm.period
                    if isinstance(period_date, str) and len(period_date) == 6:
                        year = int(period_date[:4])
                        month = int(period_date[4:6])
                        converted_date = datetime(year, month, 1).date()
                        print(f"    TTM {period_date} -> {converted_date} ✓")
                    else:
                        print(f"    TTM {period_date} -> Invalid format ✗")
        except Exception as e:
            print(f"    Date conversion error: {e}")

        print()

    # Test the exact metrics calculation for one stock
    print("=== TESTING METRICS CALCULATION ===")
    test_stock = holdings.first().stock
    print(f"Testing metrics calculation for: {test_stock.company_name}")

    try:
        # Import the calculator
        from gcia_app.metrics_calculator import PortfolioMetricsCalculator

        calculator = PortfolioMetricsCalculator()

        # Test the get_latest_period_optimized method
        stock_data_cache = {
            'market_cap_data': list(test_stock.market_cap_data.all()),
            'ttm_data': list(test_stock.ttm_data.all()),
            'quarterly_data': list(test_stock.quarterly_data.all()),
            'annual_ratios': list(test_stock.annual_ratios.all()),
            'price_data': list(test_stock.price_data.all())
        }

        periods = calculator.get_latest_period_optimized(test_stock, stock_data_cache, None)
        print(f"Available periods for calculation: {len(periods) if periods else 0}")

        if periods:
            latest_period = periods[0]
            print(f"Latest period: {latest_period}")

            # Test metrics calculation
            try:
                metrics = calculator.calculate_metrics_for_period_cached(
                    test_stock, latest_period, stock_data_cache
                )
                print(f"Metrics calculation: SUCCESS")
                print(f"Sample metrics: PATM={metrics.get('patm', 'N/A')}, Current PE={metrics.get('current_pe', 'N/A')}")
            except Exception as calc_error:
                print(f"Metrics calculation: FAILED - {calc_error}")
        else:
            print("No periods available for calculation")

    except Exception as e:
        print(f"Error importing/testing calculator: {e}")

if __name__ == "__main__":
    test_stock_financial_data()