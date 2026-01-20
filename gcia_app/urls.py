# urls.py
from django.urls import path
from gcia_app.views import (
    signup_view, login_view, home_view, logout_view, process_amcfs_nav_and_returns,
    process_portfolio_valuation, process_financial_planning, fund_analysis_metrics_view,
    download_fund_metrics, get_fund_holdings, trigger_metrics_calculation, get_calculation_progress,
    cancel_metrics_calculation, get_upload_progress
)

urlpatterns = [
    path('', login_view, name='login'),
    path('signup/', signup_view, name='signup'),
    path('login/', login_view, name='login'),
    path('home/', home_view, name='home'),
    path('logout/', logout_view, name='logout'),
    path('upload_scheme_data/', process_amcfs_nav_and_returns, name='process_amcfs_nav_and_returns'),
    path('portfolio_analysis/', process_portfolio_valuation, name='process_portfolio_valuation'),
    path('financial_planning/', process_financial_planning, name='process_financial_planning'),
    path('fund_analysis_metrics/', fund_analysis_metrics_view, name='fund_analysis_metrics'),
    path('download_fund_metrics/<int:scheme_id>/', download_fund_metrics, name='download_fund_metrics'),
    path('get_fund_holdings/<int:scheme_id>/', get_fund_holdings, name='get_fund_holdings'),

    # Portfolio Metrics Calculation URLs
    path('trigger_metrics_calculation/', trigger_metrics_calculation, name='trigger_metrics_calculation'),
    path('calculation_progress/<str:session_id>/', get_calculation_progress, name='calculation_progress'),
    path('cancel_calculation/<str:session_id>/', cancel_metrics_calculation, name='cancel_calculation'),

    # Stock Data Upload Progress URLs
    path('upload_progress/<str:session_id>/', get_upload_progress, name='upload_progress'),
]
