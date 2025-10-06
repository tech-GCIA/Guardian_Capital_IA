"""
Check metrics calculation status after update
"""
import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

from gcia_app.models import AMCFundScheme, FundHolding, FundMetricsLog, PortfolioMetricsLog, Stock

def check_status():
    """Check metrics calculation status"""

    # Get funds with holdings
    funds_with_holdings = AMCFundScheme.objects.filter(holdings__isnull=False).distinct()

    print("=== FUNDS WITH HOLDINGS ===\n")
    for i, fund in enumerate(funds_with_holdings, 1):
        print(f"{i}. {fund.name}")
        print(f"   AMFI Code: {fund.amfi_scheme_code}")
        print(f"   Holdings: {fund.holdings.count()} stocks")

        # Check which stocks have insufficient data
        holdings = fund.holdings.all()
        stocks_with_data = 0
        stocks_without_data = []

        for holding in holdings:
            stock = holding.stock
            # Quick check: does this stock have any StockTTMData?
            from gcia_app.models import StockTTMData
            has_data = StockTTMData.objects.filter(stock=stock).exists()
            if has_data:
                stocks_with_data += 1
            else:
                stocks_without_data.append(stock.company_name)

        print(f"   Stocks WITH data: {stocks_with_data}")
        print(f"   Stocks WITHOUT data: {len(stocks_without_data)}")
        if stocks_without_data:
            print(f"   Missing data for: {stocks_without_data[:5]}...")
        print()

    print("\n=== FUNDMETRICSLOG RECORDS ===\n")
    print(f"Total records: {FundMetricsLog.objects.count()}")
    for fund in funds_with_holdings:
        count = FundMetricsLog.objects.filter(scheme=fund).count()
        print(f"{fund.name}: {count} records")

        # Check periods
        periods = FundMetricsLog.objects.filter(scheme=fund).values_list('period_date', flat=True).distinct()
        print(f"  Periods: {len(periods)} unique periods")
        if count > 0:
            latest = FundMetricsLog.objects.filter(scheme=fund).order_by('-period_date').first()
            print(f"  Latest period: {latest.period_date} ({latest.period_type})")
        print()

    print("\n=== PORTFOLIOMETRICSLOG RECORDS ===\n")
    print(f"Total records: {PortfolioMetricsLog.objects.count()}")
    for fund in funds_with_holdings:
        count = PortfolioMetricsLog.objects.filter(scheme=fund).count()
        print(f"{fund.name}: {count} records")

        if count > 0:
            # Show sample metrics
            latest = PortfolioMetricsLog.objects.filter(scheme=fund).order_by('-period_date').first()
            print(f"  Latest period: {latest.period_date} ({latest.period_type})")
            print(f"  PATM: {latest.patm}")
            print(f"  QoQ Growth: {latest.qoq_growth}")
            print(f"  Current PE: {latest.current_pe}")
        print()

    print("\n=== ANALYSIS ===\n")

    # Check the problematic stocks mentioned in logs
    problematic_stocks = [
        'Clearing Corporation Of India Ltd.',
        'IndiGrid Infrastructure Trust',
        'Nexus Select Trust',
        'Other Derivaties',
        'Ather Energy Ltd.',
        'Embassy Office Parks REIT',
        'Brookfield India Real Estate Trust REIT',
        'Jubilant Bevco Ltd. (31-May-2028)',
        'Indus Infra Trust',
        'IRB InvIT Fund',
        'Akums Drugs & Pharmaceuticals Ltd.',
        'Cash & Cash Equivalent',
        'Aegis Vopak Terminals Ltd.',
        'Oswal Pumps Ltd.',
        'Anthem Biosciences Ltd.',
        'Capital Small Finance Bank Ltd.'
    ]

    print("Checking problematic stocks mentioned in logs:")
    for stock_name in problematic_stocks[:5]:  # Check first 5
        stocks = Stock.objects.filter(company_name__icontains=stock_name.split()[0])
        if stocks.exists():
            stock = stocks.first()
            from gcia_app.models import StockTTMData, StockQuarterlyData
            ttm_count = StockTTMData.objects.filter(stock=stock).count()
            quarterly_count = StockQuarterlyData.objects.filter(stock=stock).count()
            print(f"\n  {stock.company_name}:")
            print(f"    TTM records: {ttm_count}")
            print(f"    Quarterly records: {quarterly_count}")
            print(f"    Stock type: {stock.stock_type}")
        else:
            print(f"\n  {stock_name}: NOT FOUND in database")

if __name__ == "__main__":
    check_status()
