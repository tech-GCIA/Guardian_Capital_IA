# urls.py
from django.urls import path
# COMMENTED OUT - upload_stock_data and fund_analysis removed from imports
from gcia_app.views import signup_view, login_view, home_view, logout_view, process_amcfs_nav_and_returns, process_portfolio_valuation, process_financial_planning, mf_metrics_page, update_all_mf_metrics, mf_metrics_progress, download_portfolio_analysis
# COMMENTED OUT - fund_analysis, search_stocks, generate_fund_report_simple

urlpatterns = [
    path('', login_view, name='login'),
    path('signup/', signup_view, name='signup'),
    path('login/', login_view, name='login'),
    path('home/', home_view, name='home'),
    path('logout/', logout_view, name='logout'),
    path('upload_scheme_data/', process_amcfs_nav_and_returns, name='process_amcfs_nav_and_returns'),
    path('portfolio_analysis/', process_portfolio_valuation, name='process_portfolio_valuation'),
    path('financial_planning/', process_financial_planning, name='process_financial_planning'),
    
    # COMMENTED OUT - Upload Stock Data functionality disabled
    # path('upload_stock_data/', upload_stock_data, name='upload_stock_data'),
    
    # COMMENTED OUT - Fund Analysis functionality disabled
    # path('fund_analysis/', fund_analysis, name='fund_analysis'),
    # path('api/search_stocks/', search_stocks, name='search_stocks'),
    # path('generate_fund_report_simple/', generate_fund_report_simple, name='generate_fund_report_simple'),
    
    # MF Metrics URLs
    path('mf_metrics/', mf_metrics_page, name='mf_metrics'),
    path('mf_metrics/update_all/', update_all_mf_metrics, name='update_all_mf_metrics'),
    path('mf_metrics/progress/', mf_metrics_progress, name='mf_metrics_progress'),
    path('mf_metrics/download_analysis/', download_portfolio_analysis, name='download_portfolio_analysis'),
]
