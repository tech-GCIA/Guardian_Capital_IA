#!/usr/bin/env python
"""
Test the full metrics calculation pipeline on a small subset of funds
"""
import os
import sys
import django

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

from gcia_app.models import AMCFundScheme
from gcia_app.metrics_calculator import PortfolioMetricsCalculator

def test_fund_subset():
    print("=== TESTING FUND SUBSET CALCULATION ===")
    print()

    # Get first 3 active funds for testing
    test_funds = AMCFundScheme.objects.filter(is_active=True)[:3]

    print(f"Testing {test_funds.count()} funds:")
    for fund in test_funds:
        holdings_count = fund.holdings.count()
        print(f"  - {fund.name[:50]}... ({holdings_count} holdings)")
    print()

    calculator = PortfolioMetricsCalculator(session_id="test_session")

    results = {
        'success': 0,
        'failed': 0,
        'details': []
    }

    for i, fund in enumerate(test_funds, 1):
        print(f"[{i}/{test_funds.count()}] Processing: {fund.name[:50]}...")

        try:
            # Test the individual fund calculation method
            start_time = time.time()
            calculator.calculate_all_metrics_for_fund(fund.amcfundscheme_id, None)
            duration = time.time() - start_time

            print(f"  SUCCESS: Completed in {duration:.2f} seconds")
            results['success'] += 1
            results['details'].append({
                'fund': fund.name[:50],
                'status': 'SUCCESS',
                'duration': duration,
                'holdings': fund.holdings.count()
            })

        except Exception as e:
            print(f"  FAILED: {str(e)}")
            results['failed'] += 1
            results['details'].append({
                'fund': fund.name[:50],
                'status': 'FAILED',
                'error': str(e),
                'holdings': fund.holdings.count()
            })

    print("\n" + "="*60)
    print("SUBSET TEST RESULTS")
    print("="*60)
    print(f"Funds processed: {test_funds.count()}")
    print(f"Successful: {results['success']}")
    print(f"Failed: {results['failed']}")
    print(f"Success rate: {(results['success']/test_funds.count()*100):.1f}%")

    if results['success'] > 0:
        avg_duration = sum(d['duration'] for d in results['details'] if d['status'] == 'SUCCESS') / results['success']
        print(f"Average processing time: {avg_duration:.2f} seconds per fund")

        # Calculate estimated time for all funds
        total_funds = AMCFundScheme.objects.filter(is_active=True).count()
        estimated_total_time = avg_duration * total_funds
        print(f"Estimated total time for {total_funds} funds: {estimated_total_time/60:.1f} minutes")

    print("\nDetailed results:")
    for detail in results['details']:
        if detail['status'] == 'SUCCESS':
            print(f"  {detail['fund']}: {detail['status']} ({detail['duration']:.2f}s, {detail['holdings']} holdings)")
        else:
            print(f"  {detail['fund']}: {detail['status']} - {detail.get('error', 'Unknown error')}")

    return results['success'] == test_funds.count()

if __name__ == "__main__":
    import time
    success = test_fund_subset()
    if success:
        print("\n*** ALL TESTS PASSED - READY FOR FULL DEPLOYMENT ***")
    else:
        print("\n*** SOME TESTS FAILED - NEEDS INVESTIGATION ***")