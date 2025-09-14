# models.py
from django.contrib.auth.models import AbstractUser
from django.db import models

class Customer(AbstractUser):
    customer_id = models.AutoField(primary_key=True)
    phone_number = models.CharField(max_length=15, blank=True, null=True)
    pan_number = models.CharField(max_length=10, blank=True, null=True)
    address = models.TextField(blank=True, null=True)
    created = models.DateTimeField(auto_now_add=True)
    modified = models.DateTimeField(auto_now=True)

    def __str__(self):
        return str(self.customer_id)

class AMCFundScheme(models.Model):
	"""
	Stores the mutual fund scheme entries (eg: growth, dividend), that belong to an :model:kbapp.AMCFund.
	"""
	amcfundscheme_id = models.AutoField(primary_key=True)
	name = models.CharField(default='', max_length=200, verbose_name='Full Scheme Name', db_index=True)
	accord_scheme_name = models.CharField(null=True, blank=True, max_length=1024,verbose_name='Accord Scheme Name', db_index=True)
	fund_name = models.CharField(default='', max_length=200, db_index=True)
	description = models.CharField(max_length=200, default=' ', blank=True)
	amfi_scheme_code = models.IntegerField(default=0, blank=True, db_index=True)
	isin_number = models.CharField(max_length=24, blank=True, null=True, db_index=True,verbose_name='ISIN Number')
	scheme_benchmark = models.CharField(max_length=255, null=True, blank=True)
	fund_class = models.CharField(max_length=255, null=True, blank=True, db_index=True)

	is_direct_fund = models.BooleanField(default=False, db_index=True)
	launch_date = models.DateField(null=True, blank=True)
	is_active = models.BooleanField(default=False)

	latest_nav = models.FloatField(default=0.0, max_length=20)
	latest_nav_as_on_date = models.DateField(null=True, blank=True)
	assets_under_management = models.FloatField(default=0.0, null=True, blank=True)
	expense_ratio = models.FloatField(null=True, blank=True)

	returns_1_day = models.FloatField(null=True, blank=True)
	returns_7_day = models.FloatField(null=True, blank=True)
	returns_15_day = models.FloatField(null=True, blank=True)
	returns_1_mth = models.FloatField(null=True, blank=True)
	returns_3_mth = models.FloatField(null=True, blank=True)
	returns_6_mth = models.FloatField(null=True, blank=True)
	returns_1_yr = models.FloatField(null=True, blank=True)
	returns_2_yr = models.FloatField(null=True, blank=True)
	returns_3_yr = models.FloatField(null=True, blank=True)
	returns_5_yr = models.FloatField(null=True, blank=True)
	returns_7_yr = models.FloatField(null=True, blank=True)
	returns_10_yr = models.FloatField(null=True, blank=True)
	returns_15_yr = models.FloatField(null=True, blank=True)
	returns_20_yr = models.FloatField(null=True, blank=True)
	returns_25_yr = models.FloatField(null=True, blank=True)
	returns_from_launch = models.FloatField(null=True, blank=True)
	is_scheme_benchmark = models.BooleanField(default=False)

	fund_class_avg_1_day_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_7_day_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_15_day_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_1_mth_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_3_mth_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_6_mth_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_1_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_2_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_3_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_5_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_7_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_10_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_15_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_20_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_25_yr_returns = models.FloatField(null=True, blank=True)
	fund_class_avg_returns_from_launch = models.FloatField(null=True, blank=True)

	fund_rating = models.IntegerField(null=True, blank=True)
	alpha = models.FloatField(null=True, blank=True)
	beta = models.FloatField(null=True, blank=True)
	mean = models.FloatField(null=True, blank=True)
	standard_dev = models.FloatField(null=True, blank=True)
	sharpe_ratio = models.FloatField(null=True, blank=True)
	sorti_i_no = models.FloatField(null=True, blank=True)
	fund_manager = models.CharField(default='', max_length=200, null=True, blank=True)
	avg_mat = models.FloatField(null=True, blank=True)
	modified_duration = models.FloatField(null=True, blank=True)
	ytm = models.FloatField(null=True, blank=True)
	purchase_minimum_amount = models.FloatField(null=True, blank=True)
	sip_minimum_amount = models.FloatField(null=True, blank=True)
	large_cap = models.FloatField(null=True, blank=True)
	mid_cap = models.FloatField(null=True, blank=True)
	small_cap = models.FloatField(null=True, blank=True)
	pb_ratio = models.FloatField(null=True, blank=True)
	pe_ratio = models.FloatField(null=True, blank=True)
	exit_load = models.CharField(max_length=2048, null=True, blank=True)
	equity_percentage = models.FloatField(null=True, blank=True)
	debt_percentage = models.FloatField(null=True, blank=True)
	gold_percentage = models.FloatField(null=True, blank=True)
	global_equity_percentage = models.FloatField(null=True, blank=True)
	other_percentage = models.FloatField(null=True, blank=True)
	rs_quard = models.FloatField(null=True, blank=True)
	SOV = models.FloatField(null=True, blank=True)
	A = models.FloatField(null=True, blank=True)
	AA = models.FloatField(null=True, blank=True)
	AAA = models.FloatField(null=True, blank=True)
	BIG = models.FloatField(null=True, blank=True)
	cash = models.FloatField(null=True, blank=True)
	downside_deviation = models.FloatField(null=True, blank=True)
	downside_probability = models.FloatField(null=True, blank=True)
	number_of_underlying_stocks = models.FloatField(null=True, blank=True)
	up_capture = models.FloatField(null=True, blank=True)
	down_capture = models.FloatField(null=True, blank=True)

	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)

	def _str_(self):
		"""
		:return: Returns the AMCFundscheme object name. 
		"""
		return str(self.name)

class AMCFundSchemeNavLog(models.Model):
	amcfundschemenavlog_id = models.AutoField(primary_key=True)
	amcfundscheme = models.ForeignKey(AMCFundScheme, on_delete=models.CASCADE,null=True, blank=True)
	as_on_date = models.DateField(null=True, blank=True)
	nav = models.FloatField(default=0.0, max_length=20)
	
	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)
 
from django.db import models
from django.core.validators import MinValueValidator
from decimal import Decimal

class Stock(models.Model):
    """
    Model to store basic stock information
    """
    stock_id = models.AutoField(primary_key=True)
    name = models.CharField(max_length=200, verbose_name='Company Name')
    symbol = models.CharField(max_length=50, blank=True, null=True, verbose_name='Stock Symbol')
    sector = models.CharField(max_length=100, blank=True, null=True)
    industry = models.CharField(max_length=100, blank=True, null=True)
    isin = models.CharField(max_length=12, blank=True, null=True, verbose_name='ISIN Code')
    
    # Market related fields
    market_cap_category = models.CharField(max_length=50, blank=True, null=True)
    listing_date = models.DateField(blank=True, null=True)
    
    # Additional fields from Excel Base Sheet
    accord_code = models.CharField(max_length=50, blank=True, null=True, verbose_name='Accord Code')
    cap = models.CharField(max_length=20, blank=True, null=True, verbose_name='Market Cap Category')
    free_float = models.DecimalField(max_digits=5, decimal_places=4, blank=True, null=True, verbose_name='Free Float %')
    revenue_6yr_cagr = models.DecimalField(max_digits=10, decimal_places=6, blank=True, null=True, verbose_name='Revenue 6Yr CAGR')
    revenue_ttm = models.DecimalField(max_digits=10, decimal_places=6, blank=True, null=True, verbose_name='Revenue TTM')
    pat_6yr_cagr = models.DecimalField(max_digits=10, decimal_places=6, blank=True, null=True, verbose_name='PAT 6Yr CAGR')
    pat_ttm = models.DecimalField(max_digits=10, decimal_places=6, blank=True, null=True, verbose_name='PAT TTM')
    current_pr = models.DecimalField(max_digits=10, decimal_places=6, blank=True, null=True, verbose_name='Current P/R')
    pr_2yr_avg = models.DecimalField(max_digits=10, decimal_places=6, blank=True, null=True, verbose_name='P/R 2Yr Avg')
    reval_deval = models.DecimalField(max_digits=10, decimal_places=6, blank=True, null=True, verbose_name='Revaluation/Devaluation')
    bse_code = models.CharField(max_length=20, blank=True, null=True, verbose_name='BSE Code')
    nse_code = models.CharField(max_length=20, blank=True, null=True, verbose_name='NSE Code')
    
    # Status fields
    is_active = models.BooleanField(default=True)
    created = models.DateTimeField(auto_now_add=True)
    modified = models.DateTimeField(auto_now=True)
    
    class Meta:
        db_table = 'stock'
        verbose_name = 'Stock'
        verbose_name_plural = 'Stocks'
        indexes = [
            models.Index(fields=['symbol']),
            models.Index(fields=['name']),
            models.Index(fields=['sector']),
            models.Index(fields=['isin']),  # For ISIN-based matching
            models.Index(fields=['bse_code']),  # For BSE code matching
            models.Index(fields=['nse_code']),  # For NSE code matching
        ]
    
    def __str__(self):
        return f"{self.name} ({self.symbol})"

class SchemeUnderlyingHoldings(models.Model):
    schemeunderlyingholding_id = models.AutoField(primary_key=True)
    amcfundscheme = models.ForeignKey(AMCFundScheme, on_delete=models.CASCADE,null=True, blank=True)
    as_on_date = models.DateField(null=True, blank=True)
    as_on_month_end = models.CharField(max_length=255, null=True, blank=True)
    holding = models.ForeignKey(Stock, on_delete=models.CASCADE,null=True, blank=True)
    weightage = models.FloatField(default=0.0, null=True, blank=True)
    no_of_shares = models.FloatField(default=0.0, null=True, blank=True)
    
    is_active = models.BooleanField(default=True)
    created = models.DateTimeField(auto_now_add=True)
    modified = models.DateTimeField(auto_now=True)
    
    def __str__(self):
        return str(self.schemeunderlyingholding_id)
 
class StockQuarterlyData(models.Model):
    """
    Model to store quarterly financial data for stocks
    """
    quarterly_data_id = models.AutoField(primary_key=True)
    stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='quarterly_data')
    
    # Quarter identification (flexible date system)
    quarter_year = models.IntegerField(verbose_name='Year', blank=True, null=True)
    quarter_number = models.IntegerField(
        choices=[(1, 'Q1'), (2, 'Q2'), (3, 'Q3'), (4, 'Q4')],
        verbose_name='Quarter', blank=True, null=True
    )
    quarter_date = models.DateField(verbose_name='Quarter Date')
    
    # Financial metrics (using Decimal for financial precision)
    mcap = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Market Cap (Cr)', validators=[MinValueValidator(Decimal('0'))]
    )
    ttm = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='TTM (Trailing Twelve Months)', validators=[MinValueValidator(Decimal('0'))]
    )
    pat = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='PAT (Profit After Tax)', validators=[MinValueValidator(Decimal('0'))]
    )
    
    # Price related metrics
    price = models.DecimalField(
        max_digits=10, decimal_places=2, null=True, blank=True,
        verbose_name='Stock Price', validators=[MinValueValidator(Decimal('0'))]
    )
    pe_ratio = models.DecimalField(
        max_digits=10, decimal_places=2, null=True, blank=True,
        verbose_name='PE Ratio', validators=[MinValueValidator(Decimal('0'))]
    )
    pb_ratio = models.DecimalField(
        max_digits=10, decimal_places=2, null=True, blank=True,
        verbose_name='PB Ratio', validators=[MinValueValidator(Decimal('0'))]
    )
    
    # Revenue and profit metrics
    revenue = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Revenue (Cr)', validators=[MinValueValidator(Decimal('0'))]
    )
    ebitda = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='EBITDA (Cr)'
    )
    net_profit = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Net Profit (Cr)'
    )
    
    # Book value and other metrics
    book_value = models.DecimalField(
        max_digits=10, decimal_places=2, null=True, blank=True,
        verbose_name='Book Value per Share', validators=[MinValueValidator(Decimal('0'))]
    )
    dividend_yield = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True,
        verbose_name='Dividend Yield (%)'
    )
    roe = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True,
        verbose_name='ROE (%)'
    )
    roa = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True,
        verbose_name='ROA (%)'
    )
    
    # Debt metrics
    debt_to_equity = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True,
        verbose_name='Debt to Equity'
    )
    
    # Additional quarterly fields from Excel Base Sheet
    free_float_mcap = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Free Float Market Cap (Cr)'
    )
    ttm_revenue = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='TTM Revenue (Cr)'
    )
    ttm_revenue_free_float = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='TTM Revenue Free Float (Cr)'
    )
    ttm_pat_free_float = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='TTM PAT Free Float (Cr)'
    )
    quarterly_revenue = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Quarterly Revenue (Cr)'
    )
    quarterly_revenue_free_float = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Quarterly Revenue Free Float (Cr)'
    )
    quarterly_pat = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Quarterly PAT (Cr)'
    )
    quarterly_pat_free_float = models.DecimalField(
        max_digits=15, decimal_places=2, null=True, blank=True,
        verbose_name='Quarterly PAT Free Float (Cr)'
    )
    roce = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True,
        verbose_name='ROCE (%)'
    )
    retention = models.DecimalField(
        max_digits=5, decimal_places=2, null=True, blank=True,
        verbose_name='Retention (%)'
    )
    share_price = models.DecimalField(
        max_digits=10, decimal_places=2, null=True, blank=True,
        verbose_name='Share Price'
    )
    pr_quarterly = models.DecimalField(
        max_digits=10, decimal_places=2, null=True, blank=True,
        verbose_name='Quarterly P/R'
    )
    pe_quarterly = models.DecimalField(
        max_digits=10, decimal_places=2, null=True, blank=True,
        verbose_name='Quarterly PE'
    )
    
    # Metadata
    created = models.DateTimeField(auto_now_add=True)
    modified = models.DateTimeField(auto_now=True)
    
    class Meta:
        db_table = 'stock_quarterly_data'
        verbose_name = 'Stock Quarterly Data'
        verbose_name_plural = 'Stock Quarterly Data'
        unique_together = ['stock', 'quarter_date']
        indexes = [
            models.Index(fields=['stock', 'quarter_year', 'quarter_number']),
            models.Index(fields=['quarter_year', 'quarter_number']),
            models.Index(fields=['quarter_date']),
        ]
        ordering = ['-quarter_year', '-quarter_number']
    
    def __str__(self):
        return f"{self.stock.symbol} - Q{self.quarter_number} {self.quarter_year}"
    
    @property
    def quarter_label(self):
        """Returns a human-readable quarter label"""
        return f"Q{self.quarter_number}-{self.quarter_year}"


class StockUploadLog(models.Model):
    """
    Model to track file uploads and processing status
    """
    upload_id = models.AutoField(primary_key=True)
    uploaded_by = models.ForeignKey('Customer', on_delete=models.CASCADE)
    filename = models.CharField(max_length=255)
    file_size = models.BigIntegerField(help_text='File size in bytes')
    
    # Processing status
    status = models.CharField(
        max_length=20,
        choices=[
            ('pending', 'Pending'),
            ('processing', 'Processing'),
            ('completed', 'Completed'),
            ('failed', 'Failed'),
        ],
        default='pending'
    )
    
    # Processing results
    stocks_added = models.IntegerField(default=0)
    stocks_updated = models.IntegerField(default=0)
    quarterly_records_added = models.IntegerField(default=0)
    quarterly_records_updated = models.IntegerField(default=0)
    
    # Error handling
    error_message = models.TextField(blank=True, null=True)
    processing_time = models.DurationField(null=True, blank=True)
    
    # Timestamps
    uploaded_at = models.DateTimeField(auto_now_add=True)
    processing_started_at = models.DateTimeField(null=True, blank=True)
    processing_completed_at = models.DateTimeField(null=True, blank=True)
    
    class Meta:
        db_table = 'stock_upload_log'
        verbose_name = 'Stock Upload Log'
        verbose_name_plural = 'Stock Upload Logs'
        ordering = ['-uploaded_at']
    
    def __str__(self):
        return f"{self.filename} - {self.status} ({self.uploaded_at})"