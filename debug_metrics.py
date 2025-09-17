#!/usr/bin/env python
"""
Debug metrics calculation to understand the discrepancy
"""
import os
import sys
import django

# Setup Django environment
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'gcia.settings')
django.setup()

from gcia_app.models import AMCFundScheme, FundHolding

def debug_metrics_query():
    print("=== DEBUGGING METRICS CALCULATION QUERY ===")
    print()

    # This is the exact query from metrics_calculator.py line 43-45
    funds_with_holdings = AMCFundScheme.objects.filter(
        holdings__isnull=False, is_active=True
    ).distinct()

    print(f"Query result count: {funds_with_holdings.count()}")
    print()

    print("First 10 funds in query:")
    for fund in funds_with_holdings[:10]:
        holdings_count = FundHolding.objects.filter(scheme=fund).count()
        print(f"  ID: {fund.amcfundscheme_id}")
        print(f"  Name: {fund.name}")
        print(f"  Holdings count: {holdings_count}")
        print()

    # Let's also check why the logs show funds that don't exist
    print("=== CHECKING LOG MISMATCH ===")

    # Check if there are funds from a different database or different active status
    all_funds = AMCFundScheme.objects.all()
    whiteoak_funds = [f for f in all_funds if 'WhiteOak' in f.name]

    print(f"Total WhiteOak funds found: {len(whiteoak_funds)}")
    for fund in whiteoak_funds[:5]:
        print(f"  {fund.name} (Active: {fund.is_active})")

    print()
    print("=== CHECKING DATABASE CONNECTION ===")
    from django.db import connection
    print(f"Database engine: {connection.settings_dict['ENGINE']}")
    print(f"Database name: {connection.settings_dict['NAME']}")

if __name__ == "__main__":
    debug_metrics_query()