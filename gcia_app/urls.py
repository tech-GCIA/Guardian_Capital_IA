# urls.py
from django.urls import path
from gcia_app.views import signup_view, login_view, home_view, logout_view, process_amcfs_nav_and_returns, process_portfolio_valuation, process_financial_planning, upload_stock_data

urlpatterns = [
    path('', login_view, name='login'),
    path('signup/', signup_view, name='signup'),
    path('login/', login_view, name='login'),
    path('home/', home_view, name='home'),
    path('logout/', logout_view, name='logout'),
    path('upload_scheme_data/', process_amcfs_nav_and_returns, name='process_amcfs_nav_and_returns'),
    path('portfolio_analysis/', process_portfolio_valuation, name='process_portfolio_valuation'),
    path('financial_planning/', process_financial_planning, name='process_financial_planning'),
    
    path('upload_stock_data/', upload_stock_data, name='upload_stock_data'),
]
