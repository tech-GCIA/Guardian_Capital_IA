#!/usr/bin/env python
"""
Database analysis script to understand the fund holdings data issue
"""
import os
import sys
import django

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

from gcia_app.models import AMCFundScheme, FundHolding, Stock, StockMarketCap, StockTTMData, StockQuarterlyData

def analyze_database():
    print("=== GUARDIAN CAPITAL DATABASE ANALYSIS ===")
    print()

    # Basic counts
    print("1. BASIC DATABASE COUNTS:")
    print(f"   Total AMCFundScheme records: {AMCFundScheme.objects.count()}")
    print(f"   Active AMCFundScheme records: {AMCFundScheme.objects.filter(is_active=True).count()}")
    print(f"   Total FundHolding records: {FundHolding.objects.count()}")
    print(f"   Total Stock records: {Stock.objects.count()}")
    print()

    # Holdings analysis
    print("2. HOLDINGS ANALYSIS:")
    schemes_with_holdings = AMCFundScheme.objects.filter(holdings__isnull=False, is_active=True).distinct().count()
    schemes_without_holdings = AMCFundScheme.objects.filter(holdings__isnull=True, is_active=True).count()

    print(f"   Schemes WITH holdings: {schemes_with_holdings}")
    print(f"   Schemes WITHOUT holdings: {schemes_without_holdings}")
    print()

    # Sample fund analysis
    print("3. SAMPLE FUND ANALYSIS:")
    sample_funds = AMCFundScheme.objects.filter(is_active=True)[:5]

    for fund in sample_funds:
        holdings_count = FundHolding.objects.filter(scheme=fund).count()
        print(f"   Fund: {fund.name[:50]}... (ID: {fund.amcfundscheme_id})")
        print(f"   Holdings count: {holdings_count}")

        if holdings_count > 0:
            sample_holding = FundHolding.objects.filter(scheme=fund).first()
            print(f"   Sample holding: {sample_holding.stock.company_name} - {sample_holding.holding_percentage}%")
        print()

    # Stock financial data analysis
    print("4. STOCK FINANCIAL DATA ANALYSIS:")
    print(f"   StockMarketCap records: {StockMarketCap.objects.count()}")
    print(f"   StockTTMData records: {StockTTMData.objects.count()}")
    print(f"   StockQuarterlyData records: {StockQuarterlyData.objects.count()}")
    print()

    # Sample stock analysis
    print("5. SAMPLE STOCK FINANCIAL DATA:")
    if Stock.objects.exists():
        sample_stock = Stock.objects.first()
        print(f"   Sample stock: {sample_stock.company_name}")
        print(f"   Market cap records: {StockMarketCap.objects.filter(stock=sample_stock).count()}")
        print(f"   TTM data records: {StockTTMData.objects.filter(stock=sample_stock).count()}")
        print(f"   Quarterly data records: {StockQuarterlyData.objects.filter(stock=sample_stock).count()}")

        # Show sample TTM data
        ttm_data = StockTTMData.objects.filter(stock=sample_stock).first()
        if ttm_data:
            print(f"   Sample TTM data: Period {ttm_data.period}, Revenue: {ttm_data.ttm_revenue}")
    print()

    # Check for funds mentioned in logs
    print("6. SPECIFIC FUND CHECK (from logs):")
    fund_names = [
        "WhiteOak Capital Digital Bharat Fund (G) Direct",
        "WhiteOak Capital ELSS Tax Saver Fund (G) Direct",
        "Aditya Birla Sun Life Banking & Financial Services Fund (G)"
    ]

    for fund_name in fund_names:
        try:
            fund = AMCFundScheme.objects.get(name=fund_name, is_active=True)
            holdings_count = FundHolding.objects.filter(scheme=fund).count()
            print(f"   {fund_name[:50]}...")
            print(f"   Found in DB: Yes (ID: {fund.amcfundscheme_id})")
            print(f"   Holdings count: {holdings_count}")
        except AMCFundScheme.DoesNotExist:
            print(f"   {fund_name[:50]}...")
            print(f"   Found in DB: No")
        print()

if __name__ == "__main__":
    analyze_database()