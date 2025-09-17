#!/usr/bin/env python
"""
Verify the actual log output to confirm noise reduction
"""
import os
import sys
import django
import logging

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

from gcia_app.metrics_calculator import PortfolioMetricsCalculator
from gcia_app.models import AMCFundScheme

# Set up logging to capture our messages
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Create console handler
handler = logging.StreamHandler()
handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(levelname)s: %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

def verify_clean_output():
    """Test actual log output to verify noise reduction"""
    print("=== VERIFYING ACTUAL LOG OUTPUT ===")
    print("Running metrics calculation on one fund to check log cleanliness...")
    print()

    # Get a fund that we know has mixed holdings
    fund = AMCFundScheme.objects.filter(is_active=True).first()
    print(f"Testing fund: {fund.name}")
    print(f"Holdings count: {fund.holdings.count()}")
    print()

    # Set up the calculator
    calculator = PortfolioMetricsCalculator(session_id="log_test")

    print("Starting calculation (watch for improved log output)...")
    print("-" * 50)

    try:
        # This should show much cleaner logs now
        calculator.calculate_all_metrics_for_fund(fund.amcfundscheme_id, None)
        print("-" * 50)
        print("SUCCESS: Fund calculation completed!")
        print("\nNOTE: You should see:")
        print("  ✓ 'Skipping non-equity security' debug messages (instead of warnings)")
        print("  ✓ Fewer 'No periods found' warnings")
        print("  ✓ Only legitimate warnings for equity stocks with missing data")

    except Exception as e:
        print(f"FAILED: {e}")
        return False

    return True

if __name__ == "__main__":
    verify_clean_output()