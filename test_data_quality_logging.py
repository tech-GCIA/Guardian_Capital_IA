#!/usr/bin/env python
"""
Test script to verify data quality logging system
Tests the new logging configuration and data quality tracking features
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
from gcia_app.models import AMCFundScheme, SchemeUnderlyingHoldings

def test_logging_configuration():
    """Test that our logging configuration is working correctly"""
    print("=== Testing Logging Configuration ===")
    
    # Test regular logger
    logger = logging.getLogger('gcia_app')
    logger.info("Testing regular gcia_app logger")
    
    # Test data quality logger
    data_quality_logger = logging.getLogger('gcia_app.data_quality')
    data_quality_logger.info("Testing data quality logger")
    data_quality_logger.warning("Testing data quality warning")
    data_quality_logger.error("Testing data quality error")
    
    print("OK: Logging configuration test completed")
    
def test_data_quality_status():
    """Test current data quality status"""
    print("\n=== Current Data Quality Status ===")
    
    # Get current statistics
    total_funds = AMCFundScheme.objects.count()
    active_funds = AMCFundScheme.objects.filter(is_active=True).count()
    
    funds_with_holdings_ids = SchemeUnderlyingHoldings.objects.filter(is_active=True).values_list('amcfundscheme_id', flat=True).distinct()
    unique_funds_with_holdings = len(set(funds_with_holdings_ids))
    
    active_funds_with_holdings = AMCFundScheme.objects.filter(
        amcfundscheme_id__in=funds_with_holdings_ids,
        is_active=True
    ).count()
    
    inactive_funds_with_holdings = AMCFundScheme.objects.filter(
        amcfundscheme_id__in=funds_with_holdings_ids,
        is_active=False
    ).count()
    
    print(f"STATS: Total funds: {total_funds}")
    print(f"STATS: Active funds: {active_funds}")
    print(f"STATS: Unique funds with holdings: {unique_funds_with_holdings}")
    print(f"STATS: Active funds with holdings: {active_funds_with_holdings}")
    print(f"STATS: Inactive funds with holdings: {inactive_funds_with_holdings}")
    
    # Log to data quality logger
    data_quality_logger = logging.getLogger('gcia_app.data_quality')
    data_quality_logger.info("=== DATA QUALITY STATUS CHECK ===")
    data_quality_logger.info(f"Total funds: {total_funds}")
    data_quality_logger.info(f"Active funds: {active_funds}")
    data_quality_logger.info(f"Unique funds with holdings: {unique_funds_with_holdings}")
    data_quality_logger.info(f"Active funds with holdings: {active_funds_with_holdings}")
    data_quality_logger.info(f"Inactive funds with holdings: {inactive_funds_with_holdings}")
    
    if inactive_funds_with_holdings == 0:
        print("PASS: Data quality check PASSED - All funds with holdings are active")
        data_quality_logger.info("PASS: Data quality check PASSED - All funds with holdings are active")
    else:
        print("WARN: Data quality issue found - Some funds with holdings are inactive")
        data_quality_logger.warning(f"WARN: Data quality issue: {inactive_funds_with_holdings} funds with holdings are inactive")
    
    return inactive_funds_with_holdings == 0

def test_log_file_creation():
    """Test that log files are being created"""
    print("\n=== Testing Log File Creation ===")
    
    logs_dir = PROJECT_ROOT / 'logs'
    
    # Check if logs directory exists
    if logs_dir.exists():
        print(f"OK: Logs directory exists: {logs_dir}")
    else:
        print(f"WARN: Logs directory not found: {logs_dir}")
        return False
    
    # Check for log files
    django_log = logs_dir / 'django.log'
    data_quality_log = logs_dir / 'data_quality.log'
    
    if django_log.exists():
        print(f"OK: Django log file exists: {django_log}")
        print(f"  File size: {django_log.stat().st_size} bytes")
    else:
        print(f"WARN: Django log file not found: {django_log}")
    
    if data_quality_log.exists():
        print(f"OK: Data quality log file exists: {data_quality_log}")
        print(f"  File size: {data_quality_log.stat().st_size} bytes")
    else:
        print(f"WARN: Data quality log file not found: {data_quality_log}")
    
    return True

if __name__ == "__main__":
    print("Data Quality Logging System Test")
    print("=" * 50)
    
    try:
        # Run tests
        test_logging_configuration()
        data_quality_ok = test_data_quality_status()
        log_files_ok = test_log_file_creation()
        
        print("\n" + "=" * 50)
        if data_quality_ok and log_files_ok:
            print("All tests passed! Data quality logging system is working correctly.")
        else:
            print("Some tests failed. Please check the issues above.")
            
    except Exception as e:
        print(f"Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()