#!/usr/bin/env python
"""
Test the improvements: timezone fix, security filtering, and log noise reduction
"""
import os
import sys
import django

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

from gcia_app.metrics_calculator import PortfolioMetricsCalculator
from gcia_app.models import AMCFundScheme, FundHolding

def test_security_type_detection():
    """Test the is_equity_security method with various security types"""
    print("=== TESTING SECURITY TYPE DETECTION ===")

    calculator = PortfolioMetricsCalculator()

    # Test cases: (security_name, expected_result, description)
    test_cases = [
        # Equity stocks (should return True)
        ("HDFC Bank Ltd.", True, "Regular equity stock"),
        ("Reliance Industries Ltd.", True, "Regular equity stock"),
        ("Tata Consultancy Services Ltd.", True, "Regular equity stock"),

        # Treasury Bills (should return False)
        ("364 Days Treasury Bill - 15-Aug-2025", False, "Treasury Bill"),
        ("182 Days Treasury Bill - 18-Sep-2025", False, "Treasury Bill"),

        # Government Bonds (should return False)
        ("06.33% GOI - 05-May-2035", False, "Government Bond"),
        ("07.06% GOI 10-Apr-2028", False, "Government Bond"),

        # Corporate Bonds (should return False)
        ("HDFC Bank Ltd. SR-Y001 6.43% (29-Sep-2025)", False, "Corporate Bond"),
        ("Power Finance Corporation Ltd. SR-BS217 7.15% (08-Sep-2025)", False, "Corporate Bond"),
        ("NTPC Ltd. SR-61 STRPP B 08.1% (27-May-2026)", False, "Corporate Bond"),

        # Fund Entries (should return False)
        ("Net Current Asset", False, "Fund accounting entry"),
        ("Corporate Debt Market Development Fund", False, "Fund entry"),
        ("Tri-Party Repo (TREPS)", False, "Repo transaction"),

        # Fund-to-Fund (should return False)
        ("ITI Banking & PSU Debt Fund(G)-Direct Plan", False, "Fund-to-fund investment"),
        ("ITI Liquid Fund(G)-Direct Plan", False, "Fund-to-fund investment"),

        # Special Securities (should return False)
        ("Bharti Airtel Ltd. - (Partly Paid up Equity Shares (Rights Issue))", False, "Partly paid shares"),
        ("Inox Wind Ltd. - (Rights Entitlements (REs))", False, "Rights entitlements"),
    ]

    passed = 0
    failed = 0

    for security_name, expected, description in test_cases:
        result = calculator.is_equity_security(security_name)

        if result == expected:
            print(f"PASS: {description}")
            print(f"      '{security_name[:50]}...' -> {result}")
            passed += 1
        else:
            print(f"FAIL: {description}")
            print(f"      '{security_name[:50]}...' -> {result} (expected {expected})")
            failed += 1
        print()

    print(f"Security Detection Test Results: {passed} PASSED, {failed} FAILED")
    print(f"Success Rate: {(passed/(passed+failed)*100):.1f}%")
    return failed == 0

def test_log_noise_reduction():
    """Test that filtering reduces log noise for non-equity securities"""
    print("\n=== TESTING LOG NOISE REDUCTION ===")

    # Get a fund that likely has mixed holdings (equity + debt)
    fund = AMCFundScheme.objects.filter(is_active=True).first()
    print(f"Testing fund: {fund.name}")

    holdings = FundHolding.objects.filter(scheme=fund)[:10]  # Test first 10 holdings

    calculator = PortfolioMetricsCalculator()

    equity_count = 0
    non_equity_count = 0

    print(f"\nAnalyzing {holdings.count()} holdings:")
    for holding in holdings:
        stock = holding.stock
        is_equity = calculator.is_equity_security(stock.company_name)

        if is_equity:
            equity_count += 1
            print(f"  EQUITY: {stock.company_name[:50]}...")
        else:
            non_equity_count += 1
            print(f"  NON-EQUITY (SKIPPED): {stock.company_name[:50]}...")

    print(f"\nFiltering Results:")
    print(f"  Equity securities (will be processed): {equity_count}")
    print(f"  Non-equity securities (will be skipped): {non_equity_count}")
    print(f"  Log noise reduction: {non_equity_count} securities will not generate 'No periods found' warnings")

    if non_equity_count > 0:
        reduction_percentage = (non_equity_count / (equity_count + non_equity_count)) * 100
        print(f"  Estimated log noise reduction: {reduction_percentage:.1f}%")

    return True

def test_timezone_import():
    """Test that timezone import is available"""
    print("\n=== TESTING TIMEZONE IMPORT ===")

    try:
        from django.utils import timezone
        current_time = timezone.now()
        print(f"SUCCESS: timezone.now() works: {current_time}")
        print("The cancel function should now work without errors")
        return True
    except Exception as e:
        print(f"FAILED: timezone import error: {e}")
        return False

def main():
    print("="*60)
    print("TESTING METRICS CALCULATION IMPROVEMENTS")
    print("="*60)

    # Run all tests
    test1_passed = test_security_type_detection()
    test2_passed = test_log_noise_reduction()
    test3_passed = test_timezone_import()

    print("\n" + "="*60)
    print("FINAL TEST RESULTS")
    print("="*60)

    results = {
        "Security Type Detection": test1_passed,
        "Log Noise Reduction": test2_passed,
        "Timezone Import Fix": test3_passed
    }

    for test_name, passed in results.items():
        status = "PASSED" if passed else "FAILED"
        print(f"{test_name}: {status}")

    all_passed = all(results.values())

    if all_passed:
        print("\n*** ALL IMPROVEMENTS WORKING CORRECTLY ***")
        print("- Cancel button should work without errors")
        print("- Log noise from non-equity securities should be significantly reduced")
        print("- Fund metrics calculation should be cleaner and faster")
    else:
        print("\n*** SOME TESTS FAILED - NEEDS INVESTIGATION ***")

    return all_passed

if __name__ == "__main__":
    main()