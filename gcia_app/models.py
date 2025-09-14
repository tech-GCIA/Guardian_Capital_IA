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
	isin_div_or_growth_code = models.CharField(max_length=24, blank=True, null=True, db_index=True,verbose_name='ISIN Number')
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