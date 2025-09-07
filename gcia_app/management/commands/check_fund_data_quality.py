"""
Django management command to check and fix fund data quality issues.

Usage:
    python manage.py check_fund_data_quality --check-only    # Just report, don't fix
    python manage.py check_fund_data_quality --fix           # Report and fix issues
    python manage.py check_fund_data_quality --verbose       # Detailed output
"""

from django.core.management.base import BaseCommand, CommandError
from django.db.models import Count
from gcia_app.models import AMCFundScheme, SchemeUnderlyingHoldings, Stock, StockQuarterlyData


class Command(BaseCommand):
    help = 'Check and fix data quality issues related to fund activation status'
    
    def add_arguments(self, parser):
        parser.add_argument(
            '--check-only',
            action='store_true',
            help='Only check for issues, do not fix them',
        )
        parser.add_argument(
            '--fix',
            action='store_true', 
            help='Fix issues found during check',
        )
        parser.add_argument(
            '--verbose',
            action='store_true',
            help='Show detailed output',
        )
    
    def handle(self, *args, **options):
        self.stdout.write("=== FUND DATA QUALITY CHECK ===\n")
        
        # Basic statistics
        total_funds = AMCFundScheme.objects.count()
        active_funds = AMCFundScheme.objects.filter(is_active=True).count()
        inactive_funds = total_funds - active_funds
        
        total_holdings = SchemeUnderlyingHoldings.objects.filter(is_active=True).count()
        unique_funds_with_holdings = SchemeUnderlyingHoldings.objects.values_list('amcfundscheme_id', flat=True).distinct().count()
        
        self.stdout.write(f"📊 Database Overview:")
        self.stdout.write(f"   • Total funds: {total_funds}")
        self.stdout.write(f"   • Active funds: {active_funds}")
        self.stdout.write(f"   • Inactive funds: {inactive_funds}")
        self.stdout.write(f"   • Total holdings records: {total_holdings}")
        self.stdout.write(f"   • Unique funds with holdings: {unique_funds_with_holdings}")
        self.stdout.write()
        
        # Check for inactive funds with holdings
        funds_with_holdings_ids = SchemeUnderlyingHoldings.objects.values_list('amcfundscheme_id', flat=True).distinct()
        
        inactive_funds_with_holdings = AMCFundScheme.objects.filter(
            amcfundscheme_id__in=funds_with_holdings_ids,
            is_active=False
        )
        
        active_funds_with_holdings = AMCFundScheme.objects.filter(
            amcfundscheme_id__in=funds_with_holdings_ids,
            is_active=True
        )
        
        active_funds_without_holdings = AMCFundScheme.objects.filter(is_active=True).exclude(
            amcfundscheme_id__in=funds_with_holdings_ids
        )
        
        self.stdout.write(f"🔍 Data Quality Analysis:")
        self.stdout.write(f"   • Active funds with holdings: {active_funds_with_holdings.count()}")
        self.stdout.write(f"   • Active funds without holdings: {active_funds_without_holdings.count()}")
        self.stdout.write(f"   • Inactive funds with holdings: {inactive_funds_with_holdings.count()}")
        
        # Check for issues
        issues_found = False
        
        if inactive_funds_with_holdings.count() > 0:
            issues_found = True
            self.stdout.write(
                self.style.WARNING(f"\n⚠️  ISSUE FOUND: {inactive_funds_with_holdings.count()} inactive funds have holdings data!")
            )
            
            if options['verbose']:
                self.stdout.write("   Affected funds:")
                for fund in inactive_funds_with_holdings[:10]:  # Show first 10
                    holdings_count = SchemeUnderlyingHoldings.objects.filter(amcfundscheme=fund, is_active=True).count()
                    self.stdout.write(f"     • {fund.name} (ID: {fund.amcfundscheme_id}, {holdings_count} holdings)")
                if inactive_funds_with_holdings.count() > 10:
                    self.stdout.write(f"     ... and {inactive_funds_with_holdings.count() - 10} more")
            
            if options['fix']:
                self.stdout.write("\n🔧 Fixing issue...")
                activated_count = inactive_funds_with_holdings.update(is_active=True)
                self.stdout.write(
                    self.style.SUCCESS(f"✅ Successfully activated {activated_count} funds with holdings data!")
                )
            elif not options['check_only']:
                self.stdout.write("\n💡 To fix this issue, run:")
                self.stdout.write("   python manage.py check_fund_data_quality --fix")
        
        # Check for funds without quarterly data
        if options['verbose']:
            funds_without_quarterly_data = 0
            self.stdout.write(f"\n📈 Checking quarterly data availability...")
            
            for fund in active_funds_with_holdings[:5]:  # Sample check
                quarterly_data_count = StockQuarterlyData.objects.filter(
                    stock__schemeunderlyingholdings__amcfundscheme=fund
                ).distinct().count()
                
                if quarterly_data_count == 0:
                    funds_without_quarterly_data += 1
                    if options['verbose']:
                        self.stdout.write(f"   ⚠️  {fund.name} has holdings but no quarterly data")
            
            if funds_without_quarterly_data > 0:
                self.stdout.write(f"   Found {funds_without_quarterly_data} funds with holdings but no quarterly data")
        
        # Summary
        self.stdout.write("\n" + "="*50)
        if not issues_found:
            self.stdout.write(self.style.SUCCESS("✅ DATA QUALITY CHECK PASSED!"))
            self.stdout.write("   All funds with holdings data are properly activated.")
            self.stdout.write("   MF Metrics system should work correctly.")
        else:
            if options['fix']:
                self.stdout.write(self.style.SUCCESS("✅ DATA QUALITY ISSUES FIXED!"))
                self.stdout.write("   All identified issues have been resolved.")
            else:
                self.stdout.write(self.style.WARNING("⚠️  DATA QUALITY ISSUES FOUND!"))
                self.stdout.write("   Run with --fix to resolve the issues.")
        
        self.stdout.write("="*50)