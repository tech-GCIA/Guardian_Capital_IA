import datetime
from gcia_app.models import AMCFundScheme, AMCFundSchemeNavLog
from django.db.models import Avg
from django.templatetags.static import static
import os
from django.conf import settings
import numpy as np
import numpy_financial as npf

def calculate_scheme_age(launch_date):
    """
    Calculate the age of a scheme based on the launch date from database.
    
    Args:
        launch_date: Launch date (datetime object or string in format like 'YYYY-MM-DD')
    
    Returns:
        float: Age of the scheme in years
    """
    if isinstance(launch_date, str):
        launch_date = datetime.datetime.strptime(launch_date, '%Y-%m-%d')
    current_date = datetime.datetime.now()
    years_diff = current_date.year - launch_date.year
    if current_date.month < launch_date.month or (current_date.month == launch_date.month and current_date.day < launch_date.day):
        years_diff -= 1
    return years_diff

def format_currency(amount):
    """Format amount as currency"""
    if amount >= 10000000:  # 1 crore or more
        return f"₹{amount/10000000:.1f} Cr"
    elif amount >= 100000:  # 1 lakh or more
        return f"₹{amount/100000:.1f} L"
    else:
        return f"₹{amount:.1f}"

def update_avg_category_returns(fund_class=None):
    if fund_class:
        amcfs_list = AMCFundScheme.objects.filter(fund_class=fund_class, is_scheme_benchmark=False)
    else:
        amcfs_list = AMCFundScheme.objects.filter(is_scheme_benchmark=False)

    # Group by fund_class and calculate averages for all return fields explicitly
    average_data = (
        amcfs_list.filter().values("fund_class")
        .annotate(
            avg_1_day=Avg("returns_1_day"),
            avg_7_day=Avg("returns_7_day"),
            avg_15_day=Avg("returns_15_day"),
            avg_1_mth=Avg("returns_1_mth"),
            avg_3_mth=Avg("returns_3_mth"),
            avg_6_mth=Avg("returns_6_mth"),
            avg_1_yr=Avg("returns_1_yr"),
            avg_2_yr=Avg("returns_2_yr"),
            avg_3_yr=Avg("returns_3_yr"),
            avg_5_yr=Avg("returns_5_yr"),
            avg_7_yr=Avg("returns_7_yr"),
            avg_10_yr=Avg("returns_10_yr"),
            avg_15_yr=Avg("returns_15_yr"),
            avg_20_yr=Avg("returns_20_yr"),
            avg_25_yr=Avg("returns_25_yr"),
            avg_from_launch=Avg("returns_from_launch"),
        )
    )

    # Update the table with calculated averages for each fund_class
    for data in average_data:
        fund_class = data["fund_class"]

        # Update the corresponding average fields in the table
        AMCFundScheme.objects.filter(fund_class=fund_class, is_scheme_benchmark=False).update(
            fund_class_avg_1_day_returns=round(data["avg_1_day"], 2),
            fund_class_avg_7_day_returns=round(data["avg_7_day"], 2),
            fund_class_avg_15_day_returns=round(data["avg_15_day"], 2),
            fund_class_avg_1_mth_returns=round(data["avg_1_mth"], 2),
            fund_class_avg_3_mth_returns=round(data["avg_3_mth"], 2),
            fund_class_avg_6_mth_returns=round(data["avg_6_mth"], 2),
            fund_class_avg_1_yr_returns=round(data["avg_1_yr"], 2),
            fund_class_avg_2_yr_returns=round(data["avg_2_yr"], 2),
            fund_class_avg_3_yr_returns=round(data["avg_3_yr"], 2),
            fund_class_avg_5_yr_returns=round(data["avg_5_yr"], 2),
            fund_class_avg_7_yr_returns=round(data["avg_7_yr"], 2),
            fund_class_avg_10_yr_returns=round(data["avg_10_yr"], 2),
            fund_class_avg_15_yr_returns=round(data["avg_15_yr"], 2),
            fund_class_avg_20_yr_returns=round(data["avg_20_yr"], 2),
            fund_class_avg_25_yr_returns=round(data["avg_25_yr"], 2),
            fund_class_avg_returns_from_launch=round(data["avg_from_launch"], 2),
        )

    return "Successfully Update Average Category Returns"

def xirr(cashflows, dates):
    """
    Calculate the XIRR (Extended Internal Rate of Return) for irregular cash flows
    
    Args:
        cashflows: List of cash flow amounts (negative for outflows, positive for inflows)
        dates: List of datetime objects corresponding to each cash flow
    
    Returns:
        float: XIRR value as a decimal (multiply by 100 for percentage)
    """
    # Sort cashflows and dates by date
    sorted_data = sorted(zip(dates, cashflows), key=lambda x: x[0])
    dates = [x[0] for x in sorted_data]
    cashflows = [x[1] for x in sorted_data]
    
    # Convert dates to year fractions
    years = [(date - dates[0]).days / 365.0 for date in dates]
    
    # Define function to find the zero of
    def f(rate):
        return sum(cf * (1 + rate) ** (-year) for cf, year in zip(cashflows, years))
    
    # Use secant method to find the root
    try:
        # Try with numpy financial's irr as starting point
        normalized_cashflows = [cf * (1 + 0.1) ** (-year) for cf, year in zip(cashflows, years)]
        if npf:
            guess = npf.irr(normalized_cashflows)
            if np.isnan(guess):
                guess = 0.1  # Default guess if irr fails
        else:
            guess = 0.1  # Default guess if numpy_financial not available
    except:
        guess = 0.1  # Default guess
    
    # Use scipy's root finder for more robust calculation
    try:
        from scipy import optimize
        rate = optimize.newton(f, guess)
        return rate
    except:
        # Fallback to a simple secant method if scipy is not available
        x0, x1 = -0.999999, 0.999999
        f0, f1 = f(x0), f(x1)
        
        for _ in range(100):
            if abs(f1) < 1e-6:
                return x1
            x0, x1 = x1, x1 - f1 * (x1 - x0) / (f1 - f0)
            f0, f1 = f1, f(x1)
        
        return x1

def calculate_portfolio_and_benchmark_xirr(transaction_data, benchmark_name="BSE 500 - TRI"):
    """
    Calculate both actual XIRR and benchmark XIRR for an entire portfolio
    
    Args:
        transaction_data: List of all transaction dictionaries
        benchmark_name: Name of the benchmark index to compare against
    
    Returns:
        dict: Dictionary containing actual_xirr and benchmark_xirr as percentages
    """
    
    # Get the benchmark AMCFundScheme
    try:
        benchmark_scheme = AMCFundScheme.objects.get(name=benchmark_name)
    except AMCFundScheme.DoesNotExist:
        raise ValueError(f"Benchmark scheme '{benchmark_name}' not found")
    
    # Extract dates and cash flows for actual XIRR
    actual_dates = []
    actual_cash_flows = []
    
    # Variables for benchmark calculation
    benchmark_dates = []
    benchmark_cash_flows = []
    benchmark_units = 0
    
    # Process each transaction
    for transaction in transaction_data:
        if transaction['PURCHASE DATE']:
            # Parse the purchase date
            purchase_date = datetime.datetime.strptime(transaction['PURCHASE DATE'], '%d/%m/%Y')
            purchase_value = -float(transaction['PURCHASE VALUE'])  # Negative for outflow
            
            # Add to actual portfolio calculations
            actual_dates.append(purchase_date)
            actual_cash_flows.append(purchase_value)
            
            # Add to benchmark calculations
            benchmark_dates.append(purchase_date)
            benchmark_cash_flows.append(purchase_value)
            
            # Get benchmark NAV on purchase date
            purchase_date_obj = purchase_date.date()
            try:
                benchmark_nav = AMCFundSchemeNavLog.objects.filter(
                    amcfundscheme=benchmark_scheme,
                    as_on_date__lte=purchase_date_obj
                ).order_by('-as_on_date').first().nav
                
                # Calculate equivalent units of benchmark purchased
                benchmark_units += abs(purchase_value) / benchmark_nav
            except (AttributeError, TypeError):
                # If no NAV found, try to get the earliest available NAV
                try:
                    benchmark_nav = AMCFundSchemeNavLog.objects.filter(
                        amcfundscheme=benchmark_scheme
                    ).order_by('as_on_date').first().nav
                    
                    # Calculate equivalent units of benchmark purchased
                    benchmark_units += abs(purchase_value) / benchmark_nav
                except (AttributeError, TypeError):
                    print(f"Warning: No benchmark NAV found for {purchase_date_obj}")
    
    # Add current total value as the final positive cash flow for actual XIRR
    current_total_value = sum(transaction['CURRENT VALUE'] for transaction in transaction_data)
    current_date = datetime.datetime.now()
    actual_dates.append(current_date)
    actual_cash_flows.append(current_total_value)
    
    # Get current benchmark NAV
    current_date_obj = current_date.date()
    try:
        current_benchmark_nav = AMCFundSchemeNavLog.objects.filter(
            amcfundscheme=benchmark_scheme,
            as_on_date__lte=current_date_obj
        ).order_by('-as_on_date').first().nav
    except (AttributeError, TypeError):
        # If no NAV found, get the latest available NAV
        try:
            current_benchmark_nav = AMCFundSchemeNavLog.objects.filter(
                amcfundscheme=benchmark_scheme
            ).order_by('-as_on_date').first().nav
        except (AttributeError, TypeError):
            raise ValueError("No benchmark NAV data available")
    
    # Calculate current value of benchmark investment
    benchmark_current_value = benchmark_units * current_benchmark_nav
    
    # Add benchmark current value as inflow
    benchmark_dates.append(current_date)
    benchmark_cash_flows.append(benchmark_current_value)
    
    # Calculate both XIRRs
    try:
        actual_xirr_result = round(xirr(actual_cash_flows, actual_dates) * 100, 2)
    except Exception as e:
        print(f"Error calculating actual XIRR: {e}")
        actual_xirr_result = 0
    
    try:
        benchmark_xirr_result = round(xirr(benchmark_cash_flows, benchmark_dates) * 100, 2)
    except Exception as e:
        print(f"Error calculating benchmark XIRR: {e}")
        benchmark_xirr_result = 0
    
    return actual_xirr_result, benchmark_xirr_result
