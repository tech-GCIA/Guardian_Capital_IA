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
	name = models.CharField(default='', max_length=200, unique=True, verbose_name='Full Scheme Name', db_index=True)
	fund_name = models.CharField(default='', max_length=200, db_index=True)
	description = models.CharField(max_length=200, default=' ', blank=True)
	amfi_scheme_code = models.IntegerField(default=0, blank=True, db_index=True)
	isin_number = models.CharField(max_length=24, blank=True, null=True, db_index=True, verbose_name='ISIN Number')
	accord_mf_name = models.CharField(max_length=1024, blank=True, null=True, verbose_name='Accord MF Name', db_column='accord_scheme_name')
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

	# Latest Portfolio Analysis Metrics (from most recent calculation)
	latest_patm = models.FloatField(null=True, blank=True, verbose_name='Latest PATM')
	latest_qoq_growth = models.FloatField(null=True, blank=True, verbose_name='Latest QoQ Growth')
	latest_yoy_growth = models.FloatField(null=True, blank=True, verbose_name='Latest YoY Growth')
	latest_revenue_6yr_cagr = models.FloatField(null=True, blank=True, verbose_name='Latest Revenue 6Y CAGR')
	latest_pat_6yr_cagr = models.FloatField(null=True, blank=True, verbose_name='Latest PAT 6Y CAGR')
	latest_current_pe = models.FloatField(null=True, blank=True, verbose_name='Latest Current PE')
	latest_pe_2yr_avg = models.FloatField(null=True, blank=True, verbose_name='Latest PE 2Y Avg')
	latest_pe_5yr_avg = models.FloatField(null=True, blank=True, verbose_name='Latest PE 5Y Avg')
	latest_pe_2yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='Latest PE 2Y Reval/Deval')
	latest_pe_5yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='Latest PE 5Y Reval/Deval')
	latest_current_pr = models.FloatField(null=True, blank=True, verbose_name='Latest Current PR')
	latest_pr_2yr_avg = models.FloatField(null=True, blank=True, verbose_name='Latest PR 2Y Avg')
	latest_pr_5yr_avg = models.FloatField(null=True, blank=True, verbose_name='Latest PR 5Y Avg')
	latest_pr_2yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='Latest PR 2Y Reval/Deval')
	latest_pr_5yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='Latest PR 5Y Reval/Deval')
	latest_pr_10q_low = models.FloatField(null=True, blank=True, verbose_name='Latest PR 10Q Low')
	latest_pr_10q_high = models.FloatField(null=True, blank=True, verbose_name='Latest PR 10Q High')
	latest_alpha_bond_cagr = models.FloatField(null=True, blank=True, verbose_name='Latest Alpha Bond CAGR')
	latest_alpha_absolute = models.FloatField(null=True, blank=True, verbose_name='Latest Alpha Absolute')
	latest_pe_yield = models.FloatField(null=True, blank=True, verbose_name='Latest PE Yield')
	latest_growth_rate = models.FloatField(null=True, blank=True, verbose_name='Latest Growth Rate')
	latest_bond_rate = models.FloatField(null=True, blank=True, verbose_name='Latest Bond Rate')
	metrics_last_updated = models.DateTimeField(null=True, blank=True, verbose_name='Metrics Last Updated')

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

class Stock(models.Model):
	"""
	Main Stock model to store basic stock information and identifiers
	"""
	stock_id = models.AutoField(primary_key=True)
	company_name = models.CharField(max_length=255, verbose_name='Company Name', default='Unknown')
	accord_code = models.CharField(max_length=50, unique=True, verbose_name='Accord Code', db_index=True)
	sector = models.CharField(max_length=100, verbose_name='Sector', default='Unknown')
	cap = models.CharField(max_length=50, verbose_name='Cap Classification', default='Unknown')  # Large/Mid/Small cap
	
	# Stock identifiers
	bse_code = models.CharField(max_length=20, null=True, blank=True, verbose_name='BSE Code')
	nse_symbol = models.CharField(max_length=20, null=True, blank=True, verbose_name='NSE Symbol')
	isin = models.CharField(max_length=20, null=True, blank=True, verbose_name='ISIN')
	
	# Basic metrics from the sheet
	free_float = models.FloatField(null=True, blank=True, verbose_name='Free Float')
	revenue_6yr_cagr = models.FloatField(null=True, blank=True, verbose_name='Revenue 6 Year CAGR')
	revenue_ttm = models.FloatField(null=True, blank=True, verbose_name='Revenue TTM')
	pat_6yr_cagr = models.FloatField(null=True, blank=True, verbose_name='PAT 6 Year CAGR')
	pat_ttm = models.FloatField(null=True, blank=True, verbose_name='PAT TTM')
	current_value = models.FloatField(null=True, blank=True, verbose_name='Current Value')
	two_yr_avg = models.FloatField(null=True, blank=True, verbose_name='2 Year Average')
	reval_deval = models.FloatField(null=True, blank=True, verbose_name='Revaluation/Devaluation')
	
	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)
	
	def __str__(self):
		return f"{self.company_name} ({self.accord_code})"
		
	class Meta:
		verbose_name = "Stock"
		verbose_name_plural = "Stocks"

class StockMarketCap(models.Model):
	"""
	Market Cap time series data for stocks
	"""
	stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='market_cap_data')
	date = models.DateField(verbose_name='Date')
	market_cap = models.FloatField(null=True, blank=True, verbose_name='Market Cap (in crores)')
	market_cap_free_float = models.FloatField(null=True, blank=True, verbose_name='Market Cap Free Float (in crores)')
	
	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)
	
	class Meta:
		unique_together = ['stock', 'date']
		verbose_name = "Stock Market Cap"
		verbose_name_plural = "Stock Market Caps"
		ordering = ['-date']

class StockTTMData(models.Model):
	"""
	TTM (Trailing Twelve Months) financial data for stocks
	"""
	stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='ttm_data')
	period = models.CharField(max_length=6, verbose_name='Period (YYYYMM)')  # Format: 202506, 202503
	ttm_revenue = models.FloatField(null=True, blank=True, verbose_name='TTM Revenue')
	ttm_revenue_free_float = models.FloatField(null=True, blank=True, verbose_name='TTM Revenue Free Float')
	ttm_pat = models.FloatField(null=True, blank=True, verbose_name='TTM PAT')
	ttm_pat_free_float = models.FloatField(null=True, blank=True, verbose_name='TTM PAT Free Float')
	
	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)
	
	class Meta:
		unique_together = ['stock', 'period']
		verbose_name = "Stock TTM Data"
		verbose_name_plural = "Stock TTM Data"
		ordering = ['-period']

class StockQuarterlyData(models.Model):
	"""
	Quarterly financial data for stocks
	"""
	stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='quarterly_data')
	period = models.CharField(max_length=6, verbose_name='Period (YYYYMM)')  # Format: 202506, 202503
	quarterly_revenue = models.FloatField(null=True, blank=True, verbose_name='Quarterly Revenue')
	quarterly_revenue_free_float = models.FloatField(null=True, blank=True, verbose_name='Quarterly Revenue Free Float')
	quarterly_pat = models.FloatField(null=True, blank=True, verbose_name='Quarterly PAT')
	quarterly_pat_free_float = models.FloatField(null=True, blank=True, verbose_name='Quarterly PAT Free Float')
	
	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)
	
	class Meta:
		unique_together = ['stock', 'period']
		verbose_name = "Stock Quarterly Data"
		verbose_name_plural = "Stock Quarterly Data"
		ordering = ['-period']

class StockAnnualRatios(models.Model):
	"""
	Annual financial ratios for stocks
	"""
	stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='annual_ratios')
	financial_year = models.CharField(max_length=7, verbose_name='Financial Year')  # Format: 2024-25, 2023-24
	roce_percentage = models.FloatField(null=True, blank=True, verbose_name='ROCE (%)')
	roe_percentage = models.FloatField(null=True, blank=True, verbose_name='ROE (%)')
	retention_percentage = models.FloatField(null=True, blank=True, verbose_name='Retention (%)')
	
	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)
	
	class Meta:
		unique_together = ['stock', 'financial_year']
		verbose_name = "Stock Annual Ratios"
		verbose_name_plural = "Stock Annual Ratios"
		ordering = ['-financial_year']

class StockPrice(models.Model):
	"""
	Stock price and PE ratio data
	"""
	stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='price_data')
	price_date = models.DateField(verbose_name='Price Date')
	share_price = models.FloatField(null=True, blank=True, verbose_name='Share Price')
	pe_ratio = models.FloatField(null=True, blank=True, verbose_name='PE Ratio')

	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)

	class Meta:
		unique_together = ['stock', 'price_date']
		verbose_name = "Stock Price"
		verbose_name_plural = "Stock Prices"
		ordering = ['-price_date']

class FundHolding(models.Model):
	"""
	Links mutual fund schemes to their underlying stock holdings
	"""
	fund_holding_id = models.AutoField(primary_key=True)
	scheme = models.ForeignKey(AMCFundScheme, on_delete=models.CASCADE, related_name='holdings')
	stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='fund_holdings')
	holding_date = models.DateField(verbose_name='Holding Date')
	holding_percentage = models.FloatField(null=True, blank=True, verbose_name='Holding Percentage')
	market_value = models.FloatField(null=True, blank=True, verbose_name='Market Value')
	number_of_shares = models.FloatField(null=True, blank=True, verbose_name='Number of Shares')

	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)

	class Meta:
		unique_together = ['scheme', 'stock', 'holding_date']
		verbose_name = "Fund Holding"
		verbose_name_plural = "Fund Holdings"
		ordering = ['-holding_date', 'scheme', '-holding_percentage']
		indexes = [
			models.Index(fields=['scheme', 'holding_date']),
			models.Index(fields=['stock', 'holding_date']),
		]

	def __str__(self):
		return f"{self.scheme.name} holds {self.holding_percentage}% of {self.stock.company_name}"

class FundMetricsLog(models.Model):
	"""
	Stores calculated portfolio analysis metrics for each stock in each fund across all periods
	"""
	fund_metrics_log_id = models.AutoField(primary_key=True)
	scheme = models.ForeignKey(AMCFundScheme, on_delete=models.CASCADE, related_name='metrics_logs')
	stock = models.ForeignKey(Stock, on_delete=models.CASCADE, related_name='fund_metrics')
	period_date = models.DateField(verbose_name='Period Date')
	period_type = models.CharField(max_length=20, verbose_name='Period Type')  # 'quarterly', 'ttm', 'annual'

	# 22 Calculated Portfolio Analysis Metrics
	patm = models.FloatField(null=True, blank=True, verbose_name='Profit After Tax Margin')
	qoq_growth = models.FloatField(null=True, blank=True, verbose_name='Quarter over Quarter Growth')
	yoy_growth = models.FloatField(null=True, blank=True, verbose_name='Year over Year Growth')
	revenue_6yr_cagr = models.FloatField(null=True, blank=True, verbose_name='Revenue 6 Year CAGR')
	pat_6yr_cagr = models.FloatField(null=True, blank=True, verbose_name='PAT 6 Year CAGR')
	current_pe = models.FloatField(null=True, blank=True, verbose_name='Current PE')
	pe_2yr_avg = models.FloatField(null=True, blank=True, verbose_name='PE 2 Year Average')
	pe_5yr_avg = models.FloatField(null=True, blank=True, verbose_name='PE 5 Year Average')
	pe_2yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='PE 2 Year Reval/Deval')
	pe_5yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='PE 5 Year Reval/Deval')
	current_pr = models.FloatField(null=True, blank=True, verbose_name='Current Price to Revenue')
	pr_2yr_avg = models.FloatField(null=True, blank=True, verbose_name='PR 2 Year Average')
	pr_5yr_avg = models.FloatField(null=True, blank=True, verbose_name='PR 5 Year Average')
	pr_2yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='PR 2 Year Reval/Deval')
	pr_5yr_reval_deval = models.FloatField(null=True, blank=True, verbose_name='PR 5 Year Reval/Deval')
	pr_10q_low = models.FloatField(null=True, blank=True, verbose_name='PR 10 Quarter Low')
	pr_10q_high = models.FloatField(null=True, blank=True, verbose_name='PR 10 Quarter High')
	alpha_bond_cagr = models.FloatField(null=True, blank=True, verbose_name='Alpha over Bond CAGR')
	alpha_absolute = models.FloatField(null=True, blank=True, verbose_name='Alpha Absolute')
	pe_yield = models.FloatField(null=True, blank=True, verbose_name='PE Yield')
	growth_rate = models.FloatField(null=True, blank=True, verbose_name='Growth Rate')
	bond_rate = models.FloatField(null=True, blank=True, verbose_name='Bond Rate')

	created = models.DateTimeField(auto_now_add=True)
	modified = models.DateTimeField(auto_now=True)

	class Meta:
		unique_together = ['scheme', 'stock', 'period_date', 'period_type']
		verbose_name = "Fund Metrics Log"
		verbose_name_plural = "Fund Metrics Logs"
		# Removed default ordering to prevent Django from applying period_date ordering to wrong models
		# ordering = ['-period_date', 'scheme', 'stock']
		indexes = [
			models.Index(fields=['scheme', 'period_date']),
			models.Index(fields=['stock', 'period_date']),
			models.Index(fields=['scheme', 'stock', 'period_date']),
		]

	def __str__(self):
		return f"{self.scheme.name} - {self.stock.company_name} - {self.period_date}"

class MetricsCalculationSession(models.Model):
	"""
	Tracks progress of metrics calculation sessions for real-time updates
	"""
	session_id = models.CharField(max_length=255, unique=True, verbose_name='Session ID')
	user = models.ForeignKey(Customer, on_delete=models.CASCADE, related_name='calculation_sessions')
	total_funds = models.IntegerField(verbose_name='Total Funds')
	processed_funds = models.IntegerField(default=0, verbose_name='Processed Funds')
	current_fund_name = models.CharField(max_length=500, blank=True, verbose_name='Current Fund')
	current_stock_name = models.CharField(max_length=500, blank=True, verbose_name='Current Stock')
	status = models.CharField(max_length=50, verbose_name='Status')  # 'started', 'processing', 'completed', 'failed'
	progress_percentage = models.FloatField(default=0.0, verbose_name='Progress Percentage')
	error_message = models.TextField(blank=True, null=True, verbose_name='Error Message')
	started_at = models.DateTimeField(auto_now_add=True)
	completed_at = models.DateTimeField(null=True, blank=True)

	class Meta:
		verbose_name = "Metrics Calculation Session"
		verbose_name_plural = "Metrics Calculation Sessions"
		ordering = ['-started_at']

	def __str__(self):
		return f"Session {self.session_id} - {self.status} ({self.progress_percentage:.1f}%)"