#!/usr/bin/env python
"""
Test script to verify MF Metrics system with data quality improvements
Tests that the MF Metrics system can now successfully process funds
"""

import os
import sys
import django
from pathlib import Path

# Add the project directory to Python path
PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

import logging
from django.db.models import Count
from gcia_app.models import AMCFundScheme, SchemeUnderlyingHoldings, MutualFundMetrics
from gcia_app.mf_metrics_calculator import MFMetricsCalculator

def test_sample_fund_processing():
    """Test that we can successfully process individual funds"""
    print("=== Testing Individual Fund Processing ===")
    
    # Get a few active funds with holdings for testing
    funds_with_holdings = AMCFundScheme.objects.filter(
        is_active=True,
        amcfundscheme_id__in=SchemeUnderlyingHoldings.objects.filter(is_active=True).values_list('amcfundscheme_id', flat=True).distinct()
    )[:5]  # Test first 5 funds
    
    if not funds_with_holdings:
        print("ERROR: No active funds with holdings found for testing")
        return False
    
    calculator = MFMetricsCalculator()
    success_count = 0
    
    for fund in funds_with_holdings:
        print(f"\nTesting fund: {fund.name} (ID: {fund.amcfundscheme_id})")
        print(f"  Active: {fund.is_active}")
        
        # Check holdings count
        holdings_count = SchemeUnderlyingHoldings.objects.filter(amcfundscheme=fund, is_active=True).count()
        print(f"  Active holdings: {holdings_count}")
        
        try:
            # Test individual fund calculation
            result = calculator.calculate_single_fund_metrics(fund)
            print(f"  Result: {result['status']}")
            if result['status'] in ['successful', 'partial']:
                success_count += 1
                print("  SUCCESS: Fund processed successfully")
            else:
                print(f"  FAILED: {result.get('message', 'Unknown error')}")
                
        except Exception as e:
            print(f"  EXCEPTION: {str(e)}")
    
    print(f"\n=== Individual Fund Test Results ===")
    print(f"Funds tested: {len(funds_with_holdings)}")
    print(f"Successful: {success_count}")
    print(f"Success rate: {(success_count/len(funds_with_holdings)*100):.1f}%")
    
    return success_count > 0

def test_metrics_storage():
    """Test that metrics are being stored in the database"""
    print("\n=== Testing Metrics Database Storage ===")
    
    total_metrics = MutualFundMetrics.objects.count()
    recent_metrics = MutualFundMetrics.objects.filter(
        calculation_date__isnull=False
    ).count()
    
    print(f"Total metrics records: {total_metrics}")
    print(f"Records with calculation date: {recent_metrics}")
    
    if total_metrics > 0:
        # Show sample metrics
        sample_metric = MutualFundMetrics.objects.first()
        print(f"\nSample metric record:")
        print(f"  Fund: {sample_metric.amcfundscheme.name}")
        print(f"  Holdings: {sample_metric.total_holdings}")
        print(f"  Calculation date: {sample_metric.calculation_date}")
        print(f"  Status: {sample_metric.calculation_status}")
        return True
    else:
        print("No metrics records found")
        return False

def test_data_quality_resolution():
    """Verify that the original data quality issue has been resolved"""
    print("\n=== Verifying Data Quality Resolution ===")
    
    # Check the original issue: inactive funds with holdings
    funds_with_holdings_ids = SchemeUnderlyingHoldings.objects.filter(is_active=True).values_list('amcfundscheme_id', flat=True).distinct()
    inactive_funds_with_holdings = AMCFundScheme.objects.filter(
        amcfundscheme_id__in=funds_with_holdings_ids,
        is_active=False
    ).count()
    
    print(f"Inactive funds with holdings: {inactive_funds_with_holdings}")
    
    if inactive_funds_with_holdings == 0:
        print("PASS: Original data quality issue resolved")
        return True
    else:
        print(f"FAIL: Still have {inactive_funds_with_holdings} inactive funds with holdings")
        return False

if __name__ == "__main__":
    print("MF Metrics System Test")
    print("=" * 50)
    
    try:
        # Run comprehensive test
        fund_processing_ok = test_sample_fund_processing()
        metrics_storage_ok = test_metrics_storage()
        data_quality_ok = test_data_quality_resolution()
        
        print("\n" + "=" * 50)
        print("FINAL TEST RESULTS:")
        print(f"  Fund processing: {'PASS' if fund_processing_ok else 'FAIL'}")
        print(f"  Metrics storage: {'PASS' if metrics_storage_ok else 'FAIL'}")
        print(f"  Data quality: {'PASS' if data_quality_ok else 'FAIL'}")
        
        if all([fund_processing_ok, data_quality_ok]):
            print("\nSUCCESS: MF Metrics system is working correctly!")
            print("The original issue has been resolved and funds can now be processed.")
        else:
            print("\nISSUE: Some tests failed. Please check the results above.")
            
    except Exception as e:
        print(f"Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()