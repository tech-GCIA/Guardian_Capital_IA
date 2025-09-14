# views.py
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.utils.timezone import now
from datetime import timedelta
from gcia_app.forms import CustomerCreationForm
from django.contrib import messages
import openpyxl
from django.http import HttpResponse
from gcia_app.forms import ExcelUploadForm, MasterDataExcelUploadForm
from gcia_app.models import AMCFundScheme, AMCFundSchemeNavLog, Stock, StockMarketCap, StockTTMData, StockQuarterlyData, StockAnnualRatios, StockPrice, FundHolding
import pandas as pd
import os
import datetime
from django.db import transaction
from openpyxl import load_workbook
import traceback
from gcia_app.utils import update_avg_category_returns, calculate_scheme_age
from gcia_app.portfolio_analysis_ppt import create_fund_presentation
import re
from difflib import SequenceMatcher
from django.db import transaction
from gcia_app.index_scrapper_from_screener import get_bse500_pe_ratio

def signup_view(request):
    if request.method == 'POST':
        form = CustomerCreationForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('login')
    else:
        form = CustomerCreationForm()
    return render(request, 'gcia_app/signup.html', {'form': form})

def login_view(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)
            request.session['last_activity'] = now().isoformat()
            return redirect('home')
    return render(request, 'gcia_app/login.html')

@login_required
def home_view(request):
    last_activity = request.session.get('last_activity')
    if last_activity and now() - timedelta(minutes=30) > now().fromisoformat(last_activity):
        logout(request)
        return redirect('login')
    request.session['last_activity'] = now().isoformat()
    return render(request, 'gcia_app/home.html')

def logout_view(request):
    logout(request)
    return redirect('login')

def calculate_returns_for_index(scheme, index_data):
    """
    Calculate returns for various periods and update the scheme
    
    Args:
        scheme: The AMCFundScheme object to update
        index_data: DataFrame containing the index data with Date and Close Price columns
    """
    # Ensure data is sorted by date
    index_data = index_data.sort_values("Date")
    
    # Get latest date and NAV
    latest_date = index_data.iloc[-1]["Date"]
    latest_nav = float(index_data.iloc[-1]["Close Price"])
    
    # Function to calculate returns for a specific period
    def get_return_for_period(days):
        try:
            # Get data from the past
            past_date = latest_date - datetime.timedelta(days=days)
            past_data = index_data[index_data["Date"] <= past_date]
            
            if len(past_data) == 0:
                return None  # No data available for this period
            
            # Get the closest past NAV
            past_nav = float(past_data.iloc[-1]["Close Price"])
            
            # Calculate return as percentage
            if past_nav > 0:
                return ((latest_nav - past_nav) / past_nav) * 100
            return None
        except Exception:
            return None
    
    # Calculate returns for various periods
    scheme.returns_1_day = get_return_for_period(1)
    scheme.returns_7_day = get_return_for_period(7)
    scheme.returns_15_day = get_return_for_period(15)
    scheme.returns_1_mth = get_return_for_period(30)
    scheme.returns_3_mth = get_return_for_period(90)
    scheme.returns_6_mth = get_return_for_period(180)
    scheme.returns_1_yr = get_return_for_period(365)
    scheme.returns_2_yr = get_return_for_period(730)
    scheme.returns_3_yr = get_return_for_period(1095)
    scheme.returns_5_yr = get_return_for_period(1825)
    scheme.returns_7_yr = get_return_for_period(2555)
    scheme.returns_10_yr = get_return_for_period(3650)
    scheme.returns_15_yr = get_return_for_period(5475)
    scheme.returns_20_yr = get_return_for_period(7300)
    scheme.returns_25_yr = get_return_for_period(9125)
    
    # Calculate returns since inception (first available data point)
    try:
        first_data = index_data.iloc[0]
        first_nav = float(first_data["Close Price"])
        first_date = first_data["Date"]
        
        # Calculate annualized returns from inception
        if first_nav > 0:
            # Calculate days since inception
            days_since_inception = (latest_date - first_date).days
            if days_since_inception > 0:
                # Calculate simple return
                total_return = ((latest_nav - first_nav) / first_nav) * 100
                # Annualize the return
                years = days_since_inception / 365.0
                if years > 0:
                    annualized_return = ((1 + (total_return / 100)) ** (1 / years) - 1) * 100
                    scheme.returns_from_launch = annualized_return
                else:
                    scheme.returns_from_launch = total_return  # If less than a year, use simple return
    except Exception:
        scheme.returns_from_launch = None

def process_index_nav_file(excel_file):
    """
    Process INDEX NAV file to update daily prices and calculate returns
    
    Args:
        excel_file: The uploaded Excel file containing index NAV data
        
    Returns:
        dict: Statistics about the processing
    """
    # Read the Excel file, skipping the first 2 rows as header starts at row 3
    try:
        index_nav_df = pd.read_excel(excel_file, skiprows=2)
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {str(e)}")
    
    # Verify the expected columns exist
    expected_columns = ["Index Name", "Date", "Close Price"]
    for col in expected_columns:
        if col not in index_nav_df.columns:
            raise ValueError(f"Expected column '{col}' not found in the INDEX NAV file")
    
    # Drop any rows where Index Name or Date or Close Price is NaN
    index_nav_df = index_nav_df.dropna(subset=["Index Name", "Date", "Close Price"])
    
    # Ensure Date is in datetime format
    index_nav_df["Date"] = pd.to_datetime(index_nav_df["Date"]).dt.date
    
    # Get unique index names
    unique_indices = index_nav_df["Index Name"].unique()
    
    # Statistics for tracking the update process
    stats = {
        'total_indices': len(unique_indices),
        'indices_updated': 0,
        'indices_created': 0,
        'nav_entries_added': 0,
        'return_calculations_updated': 0
    }
    
    # Process each index
    with transaction.atomic():  # Use transaction to ensure data integrity
        for index_name in unique_indices:
            # Get rows for this index
            index_data = index_nav_df[index_nav_df["Index Name"] == index_name].sort_values("Date")
            
            if len(index_data) == 0:
                continue  # Skip if no data
            
            # Try to find this index in AMCFundScheme
            scheme = AMCFundScheme.objects.filter(name=index_name).first()
            
            if not scheme:
                # If not found, create a new entry
                scheme = AMCFundScheme.objects.create(
                    name=index_name,
                    is_active=True,
                    fund_name=index_name,
                    is_scheme_benchmark=True,  # This is an index, so mark it as a benchmark
                    latest_nav_as_on_date=index_data.iloc[-1]["Date"]  # Set the latest date
                )
                stats['indices_created'] += 1
            else:
                stats['indices_updated'] += 1
            
            # Process each date's NAV for this index
            for _, row in index_data.iterrows():
                date = row["Date"]
                close_price = float(row["Close Price"])
                
                # Update or create a nav log entry
                nav_log, created = AMCFundSchemeNavLog.objects.update_or_create(
                    amcfundscheme=scheme,
                    as_on_date=date,
                    nav=close_price
                )
                
                if created:
                    stats['nav_entries_added'] += 1
            
            # Update the scheme with latest NAV
            latest_data = index_data.iloc[-1]
            latest_date = latest_data["Date"]
            latest_nav = float(latest_data["Close Price"])
            
            scheme.latest_nav = latest_nav
            scheme.latest_nav_as_on_date = latest_date
            
            # Calculate returns for various periods
            calculate_returns_for_index(scheme, index_data)
            stats['return_calculations_updated'] += 1
            
            # Save the scheme with updated returns
            scheme.save()
    
    return stats

def transform_scheme_name(original_name):
    """
    Transform TopSchemes format to RATIOS PE format
    Optimized for the specific patterns in the provided files
    """
    transformed = original_name
    
    # Fix case sensitivity for "One" vs "ONE"
    transformed = transformed.replace(" One ", " ONE ")
    
    # Handle Direct Plan variations
    if " (G) Direct" in transformed:
        transformed = transformed.replace(" (G) Direct", "(G)-Direct Plan")
    elif " Direct" in transformed and not transformed.endswith(" Plan"):
        transformed = transformed.replace(" Direct", "-Direct Plan")
    
    # Handle Regular Plan variations
    if " Reg (G)" in transformed:
        transformed = transformed.replace(" Reg (G)", "-Reg(G)")
    elif " Reg " in transformed:
        transformed = transformed.replace(" Reg ", "-Reg")
    
    # Remove spaces around parentheses
    transformed = transformed.replace(" (", "(").replace(") ", ")")
    
    # Handle "FlexiCap" vs "Flexicap" difference
    transformed = transformed.replace("Flexi Cap", "Flexicap")
    
    return transformed

def find_closest_match(db_scheme_name, ratios_pe_schemes, threshold=0.85):
    """
    Find closest matching scheme name in RATIOS PE file for a given database scheme name
    Optimized for faster matching with early returns for exact matches
    """
    # Try direct transformation first
    transformed_name = transform_scheme_name(db_scheme_name)
    
    # Check for exact match
    exact_matches = [name for name in ratios_pe_schemes if name == transformed_name]
    if exact_matches:
        return exact_matches[0], 1.0  # Return with confidence score of 1.0
    
    # Extract fund family and scheme type for matching
    db_parts = db_scheme_name.lower().split()
    
    # Create a key from first few words (usually the fund family and type)
    key_parts = ' '.join(db_parts[:min(4, len(db_parts))])
    
    # Check if scheme has "Direct" or "Reg" in the name
    is_direct = "direct" in db_scheme_name.lower()
    is_reg = "reg" in db_scheme_name.lower()
    is_growth = "(g)" in db_scheme_name.lower() or "growth" in db_scheme_name.lower()
    
    # Find potential matches
    potential_matches = []
    for ratios_name in ratios_pe_schemes:
        # Skip header or empty rows
        if not isinstance(ratios_name, str) or ratios_name == 'Scheme Name':
            continue
            
        # Match fund family and type
        ratios_lower = ratios_name.lower()
        
        # Check if the key parts are in the ratios scheme name
        if key_parts in ratios_lower:
            # Check for Direct/Regular match
            ratios_is_direct = "direct" in ratios_lower
            ratios_is_reg = "reg" in ratios_lower
            ratios_is_growth = "(g)" in ratios_lower or "growth" in ratios_lower
            
            # Calculate base similarity ratio
            similarity = SequenceMatcher(None, transformed_name.lower(), ratios_lower).ratio()
            
            # Boost similarity if Direct/Regular and Growth/Income type matches
            if (is_direct and ratios_is_direct) or (is_reg and ratios_is_reg):
                similarity += 0.1
            if is_growth and ratios_is_growth:
                similarity += 0.1
                
            potential_matches.append((ratios_name, similarity))
    
    # Sort by similarity score
    potential_matches.sort(key=lambda x: x[1], reverse=True)
    
    # Return the highest scoring match if above threshold
    if potential_matches and potential_matches[0][1] >= threshold:
        return potential_matches[0]
    
    # If no good match found
    return None, 0.0

def process_ratios_pe_file(excel_file):
    """
    Process RATIOS PE file and update the database with TS_Total Count, RR3_Up Capture Ratio, and RR3_Down Capture Ratio
    This version assumes consistent file structure with headers at row 4 and specific column names
    """
    # Read the Excel file, skipping the first 3 rows as header starts at row 4
    ratios_pe_df = pd.read_excel(excel_file, skiprows=3)
    
    # Now the columns should have proper names:
    # - "Scheme Code"
    # - "Scheme Name"
    # - "TS_Total Count"
    # - "RR3_Up Capture Ratio"
    # - "RR3_Down Capture Ratio"
    
    # Verify the expected columns exist
    expected_columns = ["Scheme Code", "Scheme Name", "TS_Total Count", 
                        "RR3_Up Capture Ratio", "RR3_Down Capture Ratio"]
    
    for col in expected_columns:
        if col not in ratios_pe_df.columns:
            raise ValueError(f"Expected column '{col}' not found in the RATIOS PE file")
    
    # Drop any rows where Scheme Name is NaN
    ratios_pe_df = ratios_pe_df.dropna(subset=["Scheme Name"])
    
    # Get scheme names from database (AMCFundScheme table)
    db_schemes = AMCFundScheme.objects.all().values_list('name', 'amcfundscheme_id')
    
    # Get unique scheme names from the RATIOS PE file
    ratios_pe_schemes = ratios_pe_df["Scheme Name"].unique()
    
    # Statistics for tracking the update process
    stats = {
        'total_schemes': len(db_schemes),
        'updated': 0,
        'not_found': 0,
        'low_confidence': 0
    }
    
    # Process each scheme in the database
    with transaction.atomic():  # Use transaction to ensure data integrity
        for db_scheme_name, scheme_id in db_schemes:
            # Find matching scheme in RATIOS PE
            matched_scheme, confidence = find_closest_match(db_scheme_name, ratios_pe_schemes)
            
            if matched_scheme is not None and confidence >= 0.85:  # Only update with high confidence matches
                # Get the data for this scheme from RATIOS PE
                ratios_data = ratios_pe_df[ratios_pe_df["Scheme Name"] == matched_scheme]
                
                if not ratios_data.empty:
                    # Get the first matching row
                    data_row = ratios_data.iloc[0]
                    
                    # Update the AMCFundScheme object
                    scheme_obj = AMCFundScheme.objects.get(amcfundscheme_id=scheme_id)
                    print(scheme_obj.name, data_row)

                    # Update the fields
                    if pd.notna(data_row["TS_Total Count"]):
                        scheme_obj.number_of_underlying_stocks = float(data_row["TS_Total Count"])
                    
                    if pd.notna(data_row["RR3_Up Capture Ratio"]):
                        scheme_obj.up_capture = float(data_row["RR3_Up Capture Ratio"])
                    
                    if pd.notna(data_row["RR3_Down Capture Ratio"]):
                        scheme_obj.down_capture = float(data_row["RR3_Down Capture Ratio"])
                    
                    # Save the updated object
                    scheme_obj.save()
                    stats['updated'] += 1
            elif matched_scheme is not None:
                # Match found but confidence too low
                stats['low_confidence'] += 1
            else:
                # No match found
                stats['not_found'] += 1
    
    return stats

def process_top_schemes_data_uploaded(excel_file):
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel(excel_file)
    current_date = datetime.date.today()
    amcfs_list = AMCFundScheme.objects.filter()
    for index, row in df.iterrows():
        amcfs = amcfs_list.filter(name=row["SCHEMES"]).first()
        if not amcfs:
            amcfs = AMCFundScheme.objects.create(name=row["SCHEMES"],
                                                is_active=True,
                                                fund_name = row["FUND NAME"] if "FUND NAME" in row and not pd.isna(row["FUND NAME"]) else None,
                                                is_direct_fund = True if "SCHEMES" in row and not pd.isna(row["SCHEMES"]) and "direct" in row["SCHEMES"].lower() else False,
                                                assets_under_management = row[" A U M"] if " A U M" in row and not pd.isna(row[" A U M"]) else None,
                                                launch_date = datetime.datetime.strptime(str(row["LAUNCH DATE"]), "%d-%m-%Y") if "LAUNCH DATE" in row and not pd.isna(row["LAUNCH DATE"]) else None,
                                                returns_1_day = row["1 DAY"] if "1 DAY" in row and not pd.isna(row["1 DAY"]) else None,
                                                returns_7_day = row["7 DAY"] if "7 DAY" in row and not pd.isna(row["7 DAY"]) else None,
                                                returns_15_day = row["15 DAY"] if "15 DAY" in row and not pd.isna(row["15 DAY"]) else None,
                                                returns_1_mth = row["30 DAY"] if "30 DAY" in row and not pd.isna(row["30 DAY"]) else None,
                                                returns_3_mth = row["3 MONTH"] if "3 MONTH" in row and not pd.isna(row["3 MONTH"]) else None,
                                                returns_6_mth = row["6 MONTH"] if "6 MONTH" in row and not pd.isna(row["6 MONTH"]) else None,
                                                returns_1_yr = row["1 YEAR"] if "1 YEAR" in row and not pd.isna(row["1 YEAR"]) else None,
                                                returns_2_yr = row["2 YEAR"] if "2 YEAR" in row and not pd.isna(row["2 YEAR"]) else None,
                                                returns_3_yr = row["3 YEAR"] if "3 YEAR" in row and not pd.isna(row["3 YEAR"]) else None,
                                                returns_5_yr = row["5 YEAR"] if "5 YEAR" in row and not pd.isna(row["5 YEAR"]) else None,
                                                returns_7_yr = row["7 YEAR"] if "7 YEAR" in row and not pd.isna(row["7 YEAR"]) else None,
                                                returns_10_yr = row["10 YEAR"] if "10 YEAR" in row and not pd.isna(row["10 YEAR"]) else None,
                                                returns_15_yr = row["15 YEAR"] if "15 YEAR" in row and not pd.isna(row["15 YEAR"]) else None,
                                                returns_20_yr = row["20 YEAR"] if "20 YEAR" in row and not pd.isna(row["20 YEAR"]) else None,
                                                returns_25_yr = row["25 YEAR"] if "25 YEAR" in row and not pd.isna(row["25 YEAR"]) else None,
                                                returns_from_launch = row["SINCE INCEPTION RETURN"] if "SINCE INCEPTION RETURN" in row and not pd.isna(row["SINCE INCEPTION RETURN"]) else None,
                                                fund_rating = row["FUND RAT"] if "FUND RAT" in row and not pd.isna(row["FUND RAT"]) else None,
                                                fund_class = row["CATEGORY"] if "CATEGORY" in row and not pd.isna(row["CATEGORY"]) else None,
                                                latest_nav = row["CURRENT NAV"] if "CURRENT NAV" in row and not pd.isna(row["CURRENT NAV"]) else None,
                                                latest_nav_as_on_date = datetime.date.today(),
                                                alpha = row["ALPHA"] if "ALPHA" in row and not pd.isna(row["ALPHA"]) else None,
                                                beta = row["BETA"] if "BETA" in row and not pd.isna(row["BETA"]) else None,
                                                mean = row["MEAN"] if "MEAN" in row and not pd.isna(row["MEAN"]) else None,
                                                standard_dev = row["STANDARD DEV"] if "STANDARD DEV" in row and not pd.isna(row["STANDARD DEV"]) else None,
                                                sharpe_ratio = row["SHARPE RATIO"] if "SHARPE RATIO" in row and not pd.isna(row["SHARPE RATIO"]) else None,
                                                sorti_i_no = row["SORTI I NO"] if "SORTI I NO" in row and not pd.isna(row["SORTI I NO"]) else None,
                                                fund_manager = row["FUND MANAGER"] if "FUND MANAGER" in row and not pd.isna(row["FUND MANAGER"]) else None,
                                                avg_mat = row["AVG MAT"] if "AVG MAT" in row and not pd.isna(row["AVG MAT"]) else None,
                                                modified_duration = row["MODIFIED DURATION"] if "MODIFIED DURATION" in row and not pd.isna(row["MODIFIED DURATION"]) else None,
                                                ytm = row["YTM"] if "YTM" in row and not pd.isna(row["YTM"]) else None,
                                                purchase_minimum_amount = row["PURCHASE MIN AMOUNT"] if "PURCHASE MIN AMOUNT" in row and not pd.isna(row["PURCHASE MIN AMOUNT"]) else None,
                                                sip_minimum_amount = row["SIP MIN AMOUNT"] if "SIP MIN AMOUNT" in row and not pd.isna(row["SIP MIN AMOUNT"]) else None,
                                                large_cap = row["LARGE CAP"] if "LARGE CAP" in row and not pd.isna(row["LARGE CAP"]) else None,
                                                mid_cap = row["MID CAP"] if "MID CAP" in row and not pd.isna(row["MID CAP"]) else None,
                                                small_cap = row["SMALL CAP"] if "SMALL CAP" in row and not pd.isna(row["SMALL CAP"]) else None,
                                                pb_ratio = row["PB RATIO"] if "PB RATIO" in row and not pd.isna(row["PB RATIO"]) else None,
                                                pe_ratio = row["PE RATIO"] if "PE RATIO" in row and not pd.isna(row["PE RATIO"]) else None,
                                                exit_load = row["EXIT LOAD"] if "EXIT LOAD" in row and not pd.isna(row["EXIT LOAD"]) else None,
                                                equity_percentage = row[" EQUITY(%)"] if " EQUITY(%)" in row and not pd.isna(row[" EQUITY(%)"]) else None,
                                                debt_percentage = row[" DEBT(%)"] if " DEBT(%)" in row and not pd.isna(row[" DEBT(%)"]) else None,
                                                gold_percentage = row[" GOLD(%)"] if " GOLD(%)" in row and not pd.isna(row[" GOLD(%)"]) else None,
                                                global_equity_percentage = row[" GLOBAL EQUITY(%)"] if " GLOBAL EQUITY(%)" in row and not pd.isna(row[" GLOBAL EQUITY(%)"]) else None,
                                                other_percentage = row[" OTHER(%)"] if " OTHER(%)" in row and not pd.isna(row[" OTHER(%)"]) else None,
                                                rs_quard = row["RS QUARD"] if "RS QUARD" in row and not pd.isna(row["RS QUARD"]) else None,
                                                expense_ratio = row["EXPENSE RATIO"] if "EXPENSE RATIO" in row and not pd.isna(row["EXPENSE RATIO"]) else None,
                                                SOV = row["SOV"] if "SOV" in row and not pd.isna(row["SOV"]) else None,
                                                A = row["A"] if "A" in row and not pd.isna(row["A"]) else None,
                                                AA = row["AA"] if "AA" in row and not pd.isna(row["AA"]) else None,
                                                AAA = row["AAA"] if "AAA" in row and not pd.isna(row["AAA"]) else None,
                                                BIG = row["BIG"] if "BIG" in row and not pd.isna(row["BIG"]) else None,
                                                cash = row["CASH"] if "CASH" in row and not pd.isna(row["CASH"]) else None,
                                                downside_deviation = row["DOWNSIDE DEVIATION"] if "DOWNSIDE DEVIATION" in row and not pd.isna(row["DOWNSIDE DEVIATION"]) else None,
                                                downside_probability = row["DOWNSIDE PROBABILITY"] if "DOWNSIDE PROBABILITY" in row and not pd.isna(row["DOWNSIDE PROBABILITY"]) else None,
                                                scheme_benchmark = row["SCHEME BENCHMARK"] if "SCHEME BENCHMARK" in row and not pd.isna(row["SCHEME BENCHMARK"]) else None,
                                                is_scheme_benchmark = False if " A U M" in row and not pd.isna(row[" A U M"]) else True)
        
        if amcfs and amcfs.latest_nav_as_on_date==current_date:
            continue
        else:
            amcfs.is_active=True
            amcfs.fund_name = row["FUND NAME"] if "FUND NAME" in row and not pd.isna(row["FUND NAME"]) else None
            amcfs.is_direct_fund = True if "SCHEMES" in row and not pd.isna(row["SCHEMES"]) and "direct" in row["SCHEMES"].lower() else False
            amcfs.assets_under_management = row[" A U M"] if " A U M" in row and not pd.isna(row[" A U M"]) else None
            amcfs.launch_date = datetime.datetime.strptime(str(row["LAUNCH DATE"]), "%d-%m-%Y") if "LAUNCH DATE" in row and not pd.isna(row["LAUNCH DATE"]) else None
            amcfs.returns_1_day = row["1 DAY"] if "1 DAY" in row and not pd.isna(row["1 DAY"]) else None
            amcfs.returns_7_day = row["7 DAY"] if "7 DAY" in row and not pd.isna(row["7 DAY"]) else None
            amcfs.returns_15_day = row["15 DAY"] if "15 DAY" in row and not pd.isna(row["15 DAY"]) else None
            amcfs.returns_1_mth = row["30 DAY"] if "30 DAY" in row and not pd.isna(row["30 DAY"]) else None
            amcfs.returns_3_mth = row["3 MONTH"] if "3 MONTH" in row and not pd.isna(row["3 MONTH"]) else None
            amcfs.returns_6_mth = row["6 MONTH"] if "6 MONTH" in row and not pd.isna(row["6 MONTH"]) else None
            amcfs.returns_1_yr = row["1 YEAR"] if "1 YEAR" in row and not pd.isna(row["1 YEAR"]) else None
            amcfs.returns_2_yr = row["2 YEAR"] if "2 YEAR" in row and not pd.isna(row["2 YEAR"]) else None
            amcfs.returns_3_yr = row["3 YEAR"] if "3 YEAR" in row and not pd.isna(row["3 YEAR"]) else None
            amcfs.returns_5_yr = row["5 YEAR"] if "5 YEAR" in row and not pd.isna(row["5 YEAR"]) else None
            amcfs.returns_7_yr = row["7 YEAR"] if "7 YEAR" in row and not pd.isna(row["7 YEAR"]) else None
            amcfs.returns_10_yr = row["10 YEAR"] if "10 YEAR" in row and not pd.isna(row["10 YEAR"]) else None
            amcfs.returns_15_yr = row["15 YEAR"] if "15 YEAR" in row and not pd.isna(row["15 YEAR"]) else None
            amcfs.returns_20_yr = row["20 YEAR"] if "20 YEAR" in row and not pd.isna(row["20 YEAR"]) else None
            amcfs.returns_25_yr = row["25 YEAR"] if "25 YEAR" in row and not pd.isna(row["25 YEAR"]) else None
            amcfs.returns_from_launch = row["SINCE INCEPTION RETURN"] if "SINCE INCEPTION RETURN" in row and not pd.isna(row["SINCE INCEPTION RETURN"]) else None
            amcfs.fund_rating = row["FUND RAT"] if "FUND RAT" in row and not pd.isna(row["FUND RAT"]) else None
            amcfs.fund_class = row["CATEGORY"] if "CATEGORY" in row and not pd.isna(row["CATEGORY"]) else None
            amcfs.latest_nav = row["CURRENT NAV"] if "CURRENT NAV" in row and not pd.isna(row["CURRENT NAV"]) else None
            amcfs.latest_nav_as_on_date = datetime.date.today()
            amcfs.alpha = row["ALPHA"] if "ALPHA" in row and not pd.isna(row["ALPHA"]) else None
            amcfs.beta = row["BETA"] if "BETA" in row and not pd.isna(row["BETA"]) else None
            amcfs.mean = row["MEAN"] if "MEAN" in row and not pd.isna(row["MEAN"]) else None
            amcfs.standard_dev = row["STANDARD DEV"] if "STANDARD DEV" in row and not pd.isna(row["STANDARD DEV"]) else None
            amcfs.sharpe_ratio = row["SHARPE RATIO"] if "SHARPE RATIO" in row and not pd.isna(row["SHARPE RATIO"]) else None
            amcfs.sorti_i_no = row["SORTI I NO"] if "SORTI I NO" in row and not pd.isna(row["SORTI I NO"]) else None
            amcfs.fund_manager = row["FUND MANAGER"] if "FUND MANAGER" in row and not pd.isna(row["FUND MANAGER"]) else None
            amcfs.avg_mat = row["AVG MAT"] if "AVG MAT" in row and not pd.isna(row["AVG MAT"]) else None
            amcfs.modified_duration = row["MODIFIED DURATION"] if "MODIFIED DURATION" in row and not pd.isna(row["MODIFIED DURATION"]) else None
            amcfs.ytm = row["YTM"] if "YTM" in row and not pd.isna(row["YTM"]) else None
            amcfs.purchase_minimum_amount = row["PURCHASE MIN AMOUNT"] if "PURCHASE MIN AMOUNT" in row and not pd.isna(row["PURCHASE MIN AMOUNT"]) else None
            amcfs.sip_minimum_amount = row["SIP MIN AMOUNT"] if "SIP MIN AMOUNT" in row and not pd.isna(row["SIP MIN AMOUNT"]) else None
            amcfs.large_cap = row["LARGE CAP"] if "LARGE CAP" in row and not pd.isna(row["LARGE CAP"]) else None
            amcfs.mid_cap = row["MID CAP"] if "MID CAP" in row and not pd.isna(row["MID CAP"]) else None
            amcfs.small_cap = row["SMALL CAP"] if "SMALL CAP" in row and not pd.isna(row["SMALL CAP"]) else None
            amcfs.pb_ratio = row["PB RATIO"] if "PB RATIO" in row and not pd.isna(row["PB RATIO"]) else None
            amcfs.pe_ratio = row["PE RATIO"] if "PE RATIO" in row and not pd.isna(row["PE RATIO"]) else None
            amcfs.exit_load = row["EXIT LOAD"] if "EXIT LOAD" in row and not pd.isna(row["EXIT LOAD"]) else None
            amcfs.equity_percentage = row[" EQUITY(%)"] if " EQUITY(%)" in row and not pd.isna(row[" EQUITY(%)"]) else None
            amcfs.debt_percentage = row[" DEBT(%)"] if " DEBT(%)" in row and not pd.isna(row[" DEBT(%)"]) else None
            amcfs.gold_percentage = row[" GOLD(%)"] if " GOLD(%)" in row and not pd.isna(row[" GOLD(%)"]) else None
            amcfs.global_equity_percentage = row[" GLOBAL EQUITY(%)"] if " GLOBAL EQUITY(%)" in row and not pd.isna(row[" GLOBAL EQUITY(%)"]) else None
            amcfs.other_percentage = row[" OTHER(%)"] if " OTHER(%)" in row and not pd.isna(row[" OTHER(%)"]) else None
            amcfs.rs_quard = row["RS QUARD"] if "RS QUARD" in row and not pd.isna(row["RS QUARD"]) else None
            amcfs.expense_ratio = row["EXPENSE RATIO"] if "EXPENSE RATIO" in row and not pd.isna(row["EXPENSE RATIO"]) else None
            amcfs.SOV = row["SOV"] if "SOV" in row and not pd.isna(row["SOV"]) else None
            amcfs.A = row["A"] if "A" in row and not pd.isna(row["A"]) else None
            amcfs.AA = row["AA"] if "AA" in row and not pd.isna(row["AA"]) else None
            amcfs.AAA = row["AAA"] if "AAA" in row and not pd.isna(row["AAA"]) else None
            amcfs.BIG = row["BIG"] if "BIG" in row and not pd.isna(row["BIG"]) else None
            amcfs.cash = row["CASH"] if "CASH" in row and not pd.isna(row["CASH"]) else None
            amcfs.downside_deviation = row["DOWNSIDE DEVIATION"] if "DOWNSIDE DEVIATION" in row and not pd.isna(row["DOWNSIDE DEVIATION"]) else None
            amcfs.downside_probability = row["DOWNSIDE PROBABILITY"] if "DOWNSIDE PROBABILITY" in row and not pd.isna(row["DOWNSIDE PROBABILITY"]) else None
            amcfs.scheme_benchmark = row["SCHEME BENCHMARK"] if "SCHEME BENCHMARK" in row and not pd.isna(row["SCHEME BENCHMARK"]) else None
            amcfs.is_scheme_benchmark = False if " A U M" in row and not pd.isna(row[" A U M"]) else True
            amcfs.save()

    update_avg_category_returns()

    return "Done"

def parse_date_value(value):
    """
    Parse various date formats from Excel into Python date object
    """
    if pd.isna(value) or value is None:
        return None
    
    # If it's already a datetime object
    if isinstance(value, datetime.datetime):
        return value.date()
    elif isinstance(value, datetime.date):
        return value
    
    # If it's a string, try to parse it
    if isinstance(value, str):
        try:
            # Try different date formats
            for fmt in ['%m/%d/%Y', '%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                try:
                    return datetime.datetime.strptime(value, fmt).date()
                except ValueError:
                    continue
        except:
            pass
    
    return None

def parse_numeric_value(value):
    """
    Parse numeric values, handling Excel errors and formulas
    """
    if pd.isna(value) or value is None:
        return None
    
    # Handle Excel errors
    if isinstance(value, str) and value in ['#NAME?', '#DIV/0!', '#VALUE!', '#REF!']:
        return None
    
    try:
        return float(value)
    except (ValueError, TypeError):
        return None

def validate_excel_structure(excel_file):
    """
    Validate the Excel file structure and return diagnostic information
    """
    try:
        # Read the full Excel file including headers to validate structure
        df_full = pd.read_excel(excel_file, header=None)
        
        # Read just the data portion
        df_data = pd.read_excel(excel_file, skiprows=8)
        
        diagnostics = {
            'total_rows': len(df_full),
            'total_columns': len(df_full.columns),
            'data_rows': len(df_data),
            'data_columns': len(df_data.columns),
            'header_row_8': df_full.iloc[7].tolist() if len(df_full) > 7 else [],
            'sample_data_row': df_data.iloc[0].tolist() if len(df_data) > 0 else [],
            'columns_with_data': sum(1 for col in df_data.iloc[0] if not pd.isna(col) and col != ''),
            'null_columns': sum(1 for col in df_data.iloc[0] if pd.isna(col) or col == ''),
        }
        
        return diagnostics
    except Exception as e:
        return {'error': str(e)}

def process_underlying_stocks_file(excel_file):
    """
    Process underlying stocks holdings file with robust error handling and batch processing

    Args:
        excel_file: The uploaded Excel file containing holdings data

    Returns:
        dict: Comprehensive statistics about the processing
    """
    try:
        # Auto-detect header row by looking for expected column patterns
        holdings_df = None
        header_row = None

        for skip in range(10):
            try:
                temp_df = pd.read_excel(excel_file, skiprows=skip)
                if not temp_df.empty and len(temp_df.columns) > 10:
                    cols_str = ' '.join([str(col) for col in temp_df.columns])
                    if 'Scheme Code' in cols_str and 'Scheme Name' in cols_str and 'SD_Scheme ISIN' in cols_str:
                        holdings_df = temp_df
                        header_row = skip
                        break
            except Exception:
                continue

        if holdings_df is None:
            holdings_df = pd.read_excel(excel_file)

        if holdings_df.empty:
            raise ValueError("Holdings file contains no data")

        # Define column mappings
        expected_columns = {
            'scheme_code': ['Scheme Code', 'SCHEME CODE', 'scheme_code'],
            'scheme_name': ['Scheme Name', 'SCHEME NAME', 'scheme_name'],
            'sd_scheme_amfi_code': ['SD_Scheme AMFI Code', 'SD SCHEME AMFI CODE', 'sd_scheme_amfi_code'],
            'sd_scheme_isin': ['SD_Scheme ISIN', 'SD SCHEME ISIN', 'sd_scheme_isin'],
            'pd_month_end': ['PD_Month End', 'PD MONTH END', 'pd_month_end'],
            'pd_date': ['PD_Date', 'PD DATE', 'pd_date'],
            'pd_instrument_name': ['PD_Instrument Name', 'PD INSTRUMENT NAME', 'pd_instrument_name'],
            'pd_holding_percent': ['PD_Holding (%)', 'PD HOLDING (%)', 'pd_holding_percent'],
            'pd_market_value': ['PD_Market Value', 'PD MARKET VALUE', 'pd_market_value'],
            'pd_no_of_shares': ['PD_No of Shares', 'PD NO OF SHARES', 'pd_no_of_shares'],
            'pd_bse_code': ['PD_BSE Code', 'PD BSE CODE', 'pd_bse_code'],
            'pd_nse_symbol': ['PD_NSE Symbol', 'PD NSE SYMBOL', 'pd_nse_symbol'],
            'pd_company_isin': ['PD_Company ISIN no', 'PD COMPANY ISIN NO', 'pd_company_isin']
        }

        # Map columns
        column_mapping = {}
        available_columns = holdings_df.columns.tolist()

        for key, possible_names in expected_columns.items():
            found = False
            for possible_name in possible_names:
                if possible_name in available_columns:
                    column_mapping[key] = possible_name
                    found = True
                    break
            if not found:
                for col in available_columns:
                    if any(name.lower() in str(col).lower() for name in possible_names):
                        column_mapping[key] = col
                        found = True
                        break

        # Validate required columns
        required_columns = ['scheme_name', 'sd_scheme_isin', 'pd_instrument_name']
        missing_columns = [col for col in required_columns if col not in column_mapping]
        if missing_columns:
            raise ValueError(f"Missing required columns: {missing_columns}")

        # **PHASE 1: DATA CLEANING AND VALIDATION**
        print(f"Initial data: {len(holdings_df)} rows")

        # Remove rows with NaN in critical fields
        holdings_df = holdings_df.dropna(subset=[
            column_mapping['scheme_name'],
            column_mapping['sd_scheme_isin'],
            column_mapping['pd_instrument_name']
        ]).copy()
        print(f"After removing NaN critical fields: {len(holdings_df)} rows")

        # Convert ISIN column to string and filter out 'nan' strings
        holdings_df['clean_isin'] = holdings_df[column_mapping['sd_scheme_isin']].astype(str).str.strip()
        holdings_df = holdings_df[~holdings_df['clean_isin'].isin(['nan', 'None', ''])].copy()
        print(f"After removing invalid ISINs: {len(holdings_df)} rows")

        # **PHASE 2: DEDUPLICATION**
        original_count = len(holdings_df)

        # Create composite key for deduplication
        holdings_df['composite_key'] = (
            holdings_df['clean_isin'] + '|' +
            holdings_df[column_mapping['pd_instrument_name']].astype(str) + '|' +
            holdings_df[column_mapping['pd_date']].astype(str)
        )

        # Remove exact duplicates
        holdings_df = holdings_df.drop_duplicates(subset=['composite_key']).copy()
        duplicates_removed = original_count - len(holdings_df)
        print(f"Removed {duplicates_removed} duplicate holdings, {len(holdings_df)} unique records remain")

        # **PHASE 3: PREPARE UNIQUE SCHEMES DATA**
        unique_schemes = holdings_df.groupby('clean_isin').agg({
            column_mapping['scheme_name']: 'first',
            column_mapping.get('sd_scheme_amfi_code', column_mapping['scheme_name']): 'first'
        }).reset_index()
        print(f"Found {len(unique_schemes)} unique schemes to process")

        # Initialize comprehensive statistics
        stats = {
            'total_rows_in_file': original_count,
            'duplicates_removed': duplicates_removed,
            'unique_records_processed': len(holdings_df),
            'schemes_created': 0,
            'schemes_updated': 0,
            'schemes_failed': 0,
            'stocks_matched': 0,
            'stocks_created': 0,
            'stocks_failed': 0,
            'holdings_created': 0,
            'holdings_failed': 0,
            'batch_errors': 0
        }

        # **PHASE 4: BULK PROCESS SCHEMES FIRST**
        print("Processing schemes...")
        existing_schemes = {s.isin_number: s for s in AMCFundScheme.objects.all()}

        schemes_to_create = []
        schemes_to_update = []

        for _, scheme_row in unique_schemes.iterrows():
            scheme_isin = scheme_row['clean_isin']
            scheme_name = str(scheme_row[column_mapping['scheme_name']]).strip()

            if scheme_isin in existing_schemes:
                # Update existing scheme
                existing_scheme = existing_schemes[scheme_isin]
                existing_scheme.accord_mf_name = scheme_name
                existing_scheme.is_active = True
                schemes_to_update.append(existing_scheme)
                stats['schemes_updated'] += 1
            else:
                # Prepare new scheme
                unique_scheme_name = f"{scheme_name} ({scheme_isin})"
                amfi_code = None
                if column_mapping.get('sd_scheme_amfi_code'):
                    try:
                        amfi_val = scheme_row[column_mapping['sd_scheme_amfi_code']]
                        if pd.notna(amfi_val):
                            amfi_code = int(float(str(amfi_val)))
                    except (ValueError, TypeError):
                        pass

                schemes_to_create.append(AMCFundScheme(
                    name=unique_scheme_name,
                    accord_mf_name=scheme_name,
                    isin_number=scheme_isin,
                    amfi_scheme_code=amfi_code,
                    is_active=True
                ))
                stats['schemes_created'] += 1

        # Bulk create/update schemes
        try:
            if schemes_to_create:
                AMCFundScheme.objects.bulk_create(schemes_to_create, ignore_conflicts=True)
            if schemes_to_update:
                AMCFundScheme.objects.bulk_update(schemes_to_update, ['accord_mf_name', 'is_active'])
            print(f"Schemes processed: {stats['schemes_created']} created, {stats['schemes_updated']} updated")
        except Exception as e:
            print(f"Error in bulk scheme operations: {str(e)}")
            stats['schemes_failed'] = len(schemes_to_create) + len(schemes_to_update)

        # Refresh scheme lookup after bulk operations
        existing_schemes = {s.isin_number: s for s in AMCFundScheme.objects.all()}

        # **PHASE 5: BATCH PROCESS HOLDINGS**
        print("Processing holdings...")
        batch_size = 500
        total_batches = (len(holdings_df) + batch_size - 1) // batch_size

        for batch_num in range(total_batches):
            start_idx = batch_num * batch_size
            end_idx = min(start_idx + batch_size, len(holdings_df))
            batch_df = holdings_df.iloc[start_idx:end_idx]

            # Process each row in the batch individually to prevent rollback issues
            holdings_to_create = []

            for _, row in batch_df.iterrows():
                try:
                    # Get scheme
                    scheme_isin = row['clean_isin']
                    scheme = existing_schemes.get(scheme_isin)
                    if not scheme:
                        stats['holdings_failed'] += 1
                        continue

                    # Find or create stock (with individual transaction to prevent rollback)
                    try:
                        with transaction.atomic():
                            stock = find_or_create_stock(row, column_mapping, stats)
                    except Exception as e:
                        print(f"Error with stock operation: {str(e)}")
                        stats['holdings_failed'] += 1
                        continue

                    if not stock:
                        stats['holdings_failed'] += 1
                        continue

                    # Parse holding data
                    holding_date = parse_holding_date(row, column_mapping)
                    holding_percentage = safe_float_conversion(row.get(column_mapping.get('pd_holding_percent')))
                    market_value = safe_float_conversion(row.get(column_mapping.get('pd_market_value')))
                    number_of_shares = safe_float_conversion(row.get(column_mapping.get('pd_no_of_shares')))

                    holdings_to_create.append(FundHolding(
                        scheme=scheme,
                        stock=stock,
                        holding_date=holding_date,
                        holding_percentage=holding_percentage,
                        market_value=market_value,
                        number_of_shares=number_of_shares
                    ))

                except Exception as e:
                    print(f"Error preparing holding record: {str(e)}")
                    stats['holdings_failed'] += 1

            # Bulk create holdings for this batch
            if holdings_to_create:
                try:
                    with transaction.atomic():
                        created_holdings = FundHolding.objects.bulk_create(holdings_to_create, ignore_conflicts=True)
                        stats['holdings_created'] += len(created_holdings)
                        print(f"Batch {batch_num + 1}/{total_batches} completed: {len(created_holdings)} holdings created")
                except Exception as e:
                    print(f"Batch {batch_num + 1} bulk create failed: {str(e)}")
                    stats['batch_errors'] += 1
                    stats['holdings_failed'] += len(holdings_to_create)
            else:
                print(f"Batch {batch_num + 1}/{total_batches} completed: 0 valid holdings to create")

        print(f"Processing complete. Holdings created: {stats['holdings_created']}, failed: {stats['holdings_failed']}")
        return stats

    except Exception as e:
        raise ValueError(f"Error processing holdings file: {str(e)}")


def safe_float_conversion(value, default=None):
    """Safely convert value to float with fallback"""
    try:
        if pd.isna(value) or str(value).strip() in ['', 'nan', 'NaN', 'None']:
            return default
        return float(str(value).strip())
    except (ValueError, TypeError):
        return default


def safe_int_conversion(value, default=None):
    """Safely convert value to int with fallback"""
    try:
        if pd.isna(value) or str(value).strip() in ['', 'nan', 'NaN', 'None']:
            return default
        return int(float(str(value).strip()))
    except (ValueError, TypeError):
        return default


def parse_holding_date(row, column_mapping):
    """Parse holding date with fallback to today"""
    holding_date = datetime.date.today()
    if column_mapping.get('pd_date'):
        try:
            date_val = row[column_mapping.get('pd_date')]
            if pd.notna(date_val):
                if isinstance(date_val, datetime.datetime):
                    holding_date = date_val.date()
                elif isinstance(date_val, datetime.date):
                    holding_date = date_val
                else:
                    holding_date = pd.to_datetime(str(date_val)).date()
        except Exception:
            pass
    return holding_date


def find_or_create_stock(row, column_mapping, stats):
    """Find existing stock or create new one"""
    try:
        instrument_name = str(row[column_mapping['pd_instrument_name']]).strip()
        company_isin = str(row.get(column_mapping.get('pd_company_isin', ''), '')).strip()
        bse_code = str(row.get(column_mapping.get('pd_bse_code', ''), '')).strip()
        nse_symbol = str(row.get(column_mapping.get('pd_nse_symbol', ''), '')).strip()

        # Clean values
        company_isin = company_isin if company_isin not in ['nan', 'None', ''] else None
        bse_code = safe_int_conversion(bse_code)
        nse_symbol = nse_symbol if nse_symbol not in ['nan', 'None', ''] else None

        # Try to match existing stock
        stock = None
        if company_isin:
            stock = Stock.objects.filter(isin=company_isin).first()
        if not stock and bse_code:
            stock = Stock.objects.filter(bse_code=bse_code).first()
        if not stock and nse_symbol:
            stock = Stock.objects.filter(nse_symbol=nse_symbol).first()

        if stock:
            stats['stocks_matched'] += 1
            return stock
        else:
            # Create new stock with unique accord_code
            import uuid
            unique_code = f"AUTO_{uuid.uuid4().hex[:8]}"
            stock = Stock.objects.create(
                company_name=instrument_name,
                accord_code=unique_code,
                isin=company_isin,
                bse_code=bse_code,
                nse_symbol=nse_symbol,
                sector='Unknown'
            )
            stats['stocks_created'] += 1
            return stock

    except Exception as e:
        print(f"Error finding/creating stock: {str(e)}")
        stats['stocks_failed'] += 1
        return None

def process_stocks_base_sheet(excel_file):
    """
    Process the Stocks Base Sheet Excel file based on the multi-level header structure
    """
    try:
        # First, validate the Excel structure
        diagnostics = validate_excel_structure(excel_file)
        print(f"Excel Structure Diagnostics: {diagnostics}")
        
        # Read the Excel file - skip first 8 rows to get to actual data (row 9 onwards)
        df = pd.read_excel(excel_file, skiprows=8)
        
        # Get column headers (these should be from row 8 of the original file)
        headers = df.columns.tolist()
        print(f"Total columns after skiprows=8: {len(headers)}")
        print(f"First 20 headers: {headers[:20]}")
        print(f"Sample of first row data: {df.iloc[0][:20].tolist() if len(df) > 0 else 'No data'}")
        
        # Define column mapping based on the header structure analysis
        # Column positions are 0-based after reading with skiprows=8
        column_mapping = {
            'basic_info': {
                's_no': 0,           # S. No.
                'company_name': 1,   # Company Name
                'accord_code': 2,    # Accord Code
                'sector': 3,         # Sector
                'cap': 4,            # Cap
                'free_float': 5,     # Free Float
                'revenue_6yr_cagr': 6, # 6 Year CAGR Revenue
                'revenue_ttm': 7,    # TTM Revenue
                'pat_6yr_cagr': 8,   # 6 Year CAGR PAT
                'pat_ttm': 9,        # TTM PAT
                'current': 10,       # Current
                'two_yr_avg': 11,    # 2 Yr Avg
                'reval_deval': 12,   # Reval/deval
            }
        }
        
        # Identify time-series column ranges based on CORRECTED header analysis from Row 6
        # Market Cap dates: columns 14-41 (verified from actual data analysis)
        market_cap_start_col = 14
        market_cap_end_col = 41
        
        # Market Cap Free Float: columns 43-70 (Row 6 analysis shows FF at col 44+)
        market_cap_ff_start_col = 43
        market_cap_ff_end_col = 70
        
        # TTM Revenue: columns 72-107 (Row 6 analysis shows TTM Revenue at col 73+)
        ttm_revenue_start_col = 72
        ttm_revenue_end_col = 107
        
        # TTM Revenue Free Float: columns 108-143 (Row 6 analysis shows TTM Revenue FF at col 109+)
        ttm_revenue_ff_start_col = 108
        ttm_revenue_ff_end_col = 143
        
        # TTM PAT: columns 144-179 (Row 6 analysis shows TTM PAT at col 145+)
        ttm_pat_start_col = 144
        ttm_pat_end_col = 179
        
        # TTM PAT Free Float: columns 180-215 (Row 6 analysis shows TTM PAT FF at col 181+)
        ttm_pat_ff_start_col = 180
        ttm_pat_ff_end_col = 215
        
        # Quarterly Revenue: columns 216-251 (Row 6 analysis shows Quarterly Revenue at col 217+)
        qtr_revenue_start_col = 216
        qtr_revenue_end_col = 251
        
        # Quarterly Revenue Free Float: columns 252-287 (Row 6 analysis shows at col 253+)
        qtr_revenue_ff_start_col = 252
        qtr_revenue_ff_end_col = 287
        
        # Quarterly PAT: columns 288-323 (Row 6 analysis shows Quarterly PAT at col 289+)
        qtr_pat_start_col = 288
        qtr_pat_end_col = 323
        
        # Quarterly PAT Free Float: columns 324-359 (Row 6 analysis shows at col 325+)
        qtr_pat_ff_start_col = 324
        qtr_pat_ff_end_col = 359
        
        # Annual ratios: ROCE, ROE, Retention (12 years each) - based on Row 6 analysis
        roce_start_col = 361  # Row 6 shows ROCE at col 361, data starts at 361 (CORRECTED)
        roce_end_col = 372    # 12 years of data
        
        roe_start_col = 374   # Row 6 shows ROE at col 374, data starts at 374 (CORRECTED)
        roe_end_col = 385     # 12 years of data
        
        retention_start_col = 387  # Row 6 shows Retention at col 387, data starts at 387 (CORRECTED)
        retention_end_col = 398    # 12 years of data
        
        # Stock Price data (Row 6 shows Share Price at col 400)
        price_start_col = 399  # Data starts just before the header
        price_end_col = 409    # Various price dates
        
        pe_start_col = 410     # Row 6 shows PE around col 413
        pe_end_col = 420       # PE ratios
        
        # Identifiers at the end - fixed positions based on analysis
        bse_code_col = 422  # BSE Code at fixed position
        nse_code_col = 423  # NSE Symbol at fixed position  
        isin_col = 424      # ISIN at fixed position
        
        # Statistics for tracking
        stats = {
            'total_stocks': 0,
            'stocks_created': 0,
            'stocks_updated': 0,
            'market_cap_records': 0,
            'ttm_records': 0,
            'quarterly_records': 0,
            'annual_ratios_records': 0,
            'price_records': 0,
            'skipped_rows': 0
        }
        
        # Process each row
        print(f"Processing {len(df)} rows...")
        
        with transaction.atomic():
            for index, row in df.iterrows():
                try:
                    # Skip sample rows or empty rows
                    if (pd.isna(row.iloc[column_mapping['basic_info']['company_name']]) or 
                        str(row.iloc[column_mapping['basic_info']['s_no']]).upper() == 'XX' or
                        pd.isna(row.iloc[column_mapping['basic_info']['accord_code']])):
                        stats['skipped_rows'] += 1
                        continue
                    
                    company_name = str(row.iloc[column_mapping['basic_info']['company_name']]).strip()
                    accord_code = str(row.iloc[column_mapping['basic_info']['accord_code']]).strip()
                    
                    if not company_name or not accord_code:
                        stats['skipped_rows'] += 1
                        continue
                    
                    stats['total_stocks'] += 1
                    
                    # Create or update Stock record
                    stock, created = Stock.objects.update_or_create(
                        accord_code=accord_code,
                        defaults={
                            'company_name': company_name,
                            'sector': str(row.iloc[column_mapping['basic_info']['sector']]) if not pd.isna(row.iloc[column_mapping['basic_info']['sector']]) else '',
                            'cap': str(row.iloc[column_mapping['basic_info']['cap']]) if not pd.isna(row.iloc[column_mapping['basic_info']['cap']]) else '',
                            'free_float': parse_numeric_value(row.iloc[column_mapping['basic_info']['free_float']]),
                            'revenue_6yr_cagr': parse_numeric_value(row.iloc[column_mapping['basic_info']['revenue_6yr_cagr']]),
                            'revenue_ttm': parse_numeric_value(row.iloc[column_mapping['basic_info']['revenue_ttm']]),
                            'pat_6yr_cagr': parse_numeric_value(row.iloc[column_mapping['basic_info']['pat_6yr_cagr']]),
                            'pat_ttm': parse_numeric_value(row.iloc[column_mapping['basic_info']['pat_ttm']]),
                            'current_value': parse_numeric_value(row.iloc[column_mapping['basic_info']['current']]),
                            'two_yr_avg': parse_numeric_value(row.iloc[column_mapping['basic_info']['two_yr_avg']]),
                            'reval_deval': parse_numeric_value(row.iloc[column_mapping['basic_info']['reval_deval']]),
                            'bse_code': str(row.iloc[bse_code_col]) if bse_code_col and not pd.isna(row.iloc[bse_code_col]) else None,
                            'nse_symbol': str(row.iloc[nse_code_col]) if nse_code_col and not pd.isna(row.iloc[nse_code_col]) else None,
                            'isin': str(row.iloc[isin_col]) if isin_col and not pd.isna(row.iloc[isin_col]) else None,
                        }
                    )
                    
                    if created:
                        stats['stocks_created'] += 1
                    else:
                        stats['stocks_updated'] += 1
                    
                    # Add detailed logging for first row to confirm processing
                    if stats['total_stocks'] == 1:
                        print(f"=== DEBUGGING ROW {stats['total_stocks']} ===")
                        print(f"Company: {company_name}, Accord: {accord_code}")
                        print(f"Free Float raw value: {row.iloc[column_mapping['basic_info']['free_float']]}")
                        print(f"Free Float parsed: {parse_numeric_value(row.iloc[column_mapping['basic_info']['free_float']])}")
                        print(f"Revenue TTM raw: {row.iloc[column_mapping['basic_info']['revenue_ttm']]}")
                        print(f"Revenue TTM parsed: {parse_numeric_value(row.iloc[column_mapping['basic_info']['revenue_ttm']])}")
                        print(f"Market cap sample (col {market_cap_start_col}): {row.iloc[market_cap_start_col] if market_cap_start_col < len(row) else 'N/A'}")
                        print(f"TTM sample (col {ttm_revenue_start_col}): {row.iloc[ttm_revenue_start_col] if ttm_revenue_start_col < len(row) else 'N/A'}")
                        print("===========================")
                    
                    # Process Market Cap data (specific dates)
                    market_cap_dates = [
                        '2025-08-19', '2025-06-30', '2025-03-28', '2025-01-31', '2024-12-31',
                        '2024-09-30', '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29',
                        '2023-06-30', '2023-03-31', '2022-12-30', '2022-09-30', '2022-06-30',
                        '2022-03-31', '2021-12-31', '2021-09-30', '2021-06-30', '2021-03-31',
                        '2020-12-31', '2020-09-30', '2020-06-30', '2020-03-31', '2019-12-31',
                        '2019-09-30', '2019-06-28', '2019-04-01'
                    ]
                    
                    for i, date_str in enumerate(market_cap_dates):
                        if market_cap_start_col + i < len(row):
                            market_cap_val = parse_numeric_value(row.iloc[market_cap_start_col + i])
                            market_cap_ff_val = None
                            if market_cap_ff_start_col + i < len(row):
                                market_cap_ff_val = parse_numeric_value(row.iloc[market_cap_ff_start_col + i])
                            
                            if market_cap_val is not None or market_cap_ff_val is not None:
                                try:
                                    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                                    StockMarketCap.objects.update_or_create(
                                        stock=stock,
                                        date=date_obj,
                                        defaults={
                                            'market_cap': market_cap_val,
                                            'market_cap_free_float': market_cap_ff_val,
                                        }
                                    )
                                    stats['market_cap_records'] += 1
                                except ValueError:
                                    continue
                    
                    # Process TTM data (periods like 202506, 202503)
                    ttm_periods = [
                        '202506', '202503', '202412', '202409', '202406', '202403',
                        '202312', '202309', '202306', '202303', '202212', '202209',
                        '202206', '202203', '202112', '202109', '202106', '202103',
                        '202012', '202009', '202006', '202003', '201912', '201909',
                        '201906', '201903', '201812', '201809', '201806', '201803',
                        '201712', '201709', '201706', '201703', '201612'
                    ]
                    
                    for i, period in enumerate(ttm_periods):
                        if (ttm_revenue_start_col + i < len(row) and 
                            ttm_revenue_ff_start_col + i < len(row) and
                            ttm_pat_start_col + i < len(row) and 
                            ttm_pat_ff_start_col + i < len(row)):
                            
                            ttm_rev = parse_numeric_value(row.iloc[ttm_revenue_start_col + i])
                            ttm_rev_ff = parse_numeric_value(row.iloc[ttm_revenue_ff_start_col + i])
                            ttm_pat = parse_numeric_value(row.iloc[ttm_pat_start_col + i])
                            ttm_pat_ff = parse_numeric_value(row.iloc[ttm_pat_ff_start_col + i])
                            
                            if any(val is not None for val in [ttm_rev, ttm_rev_ff, ttm_pat, ttm_pat_ff]):
                                StockTTMData.objects.update_or_create(
                                    stock=stock,
                                    period=period,
                                    defaults={
                                        'ttm_revenue': ttm_rev,
                                        'ttm_revenue_free_float': ttm_rev_ff,
                                        'ttm_pat': ttm_pat,
                                        'ttm_pat_free_float': ttm_pat_ff,
                                    }
                                )
                                stats['ttm_records'] += 1
                    
                    # Process Quarterly data (same periods as TTM)
                    for i, period in enumerate(ttm_periods):
                        if (qtr_revenue_start_col + i < len(row) and 
                            qtr_revenue_ff_start_col + i < len(row) and
                            qtr_pat_start_col + i < len(row) and 
                            qtr_pat_ff_start_col + i < len(row)):
                            
                            qtr_rev = parse_numeric_value(row.iloc[qtr_revenue_start_col + i])
                            qtr_rev_ff = parse_numeric_value(row.iloc[qtr_revenue_ff_start_col + i])
                            qtr_pat = parse_numeric_value(row.iloc[qtr_pat_start_col + i])
                            qtr_pat_ff = parse_numeric_value(row.iloc[qtr_pat_ff_start_col + i])
                            
                            if any(val is not None for val in [qtr_rev, qtr_rev_ff, qtr_pat, qtr_pat_ff]):
                                StockQuarterlyData.objects.update_or_create(
                                    stock=stock,
                                    period=period,
                                    defaults={
                                        'quarterly_revenue': qtr_rev,
                                        'quarterly_revenue_free_float': qtr_rev_ff,
                                        'quarterly_pat': qtr_pat,
                                        'quarterly_pat_free_float': qtr_pat_ff,
                                    }
                                )
                                stats['quarterly_records'] += 1
                    
                    # Process Annual Ratios (financial years like 2024-25, 2023-24)
                    annual_years = [
                        '2024-25', '2023-24', '2022-23', '2021-22', '2020-21', '2019-20',
                        '2018-19', '2017-18', '2016-17', '2015-16', '2014-15', '2013-14'
                    ]
                    
                    for i, year in enumerate(annual_years):
                        if (roce_start_col + i < len(row) and 
                            roe_start_col + i < len(row) and
                            retention_start_col + i < len(row)):
                            
                            roce = parse_numeric_value(row.iloc[roce_start_col + i])
                            roe = parse_numeric_value(row.iloc[roe_start_col + i])
                            retention = parse_numeric_value(row.iloc[retention_start_col + i])
                            
                            if any(val is not None for val in [roce, roe, retention]):
                                StockAnnualRatios.objects.update_or_create(
                                    stock=stock,
                                    financial_year=year,
                                    defaults={
                                        'roce_percentage': roce,
                                        'roe_percentage': roe,
                                        'retention_percentage': retention,
                                    }
                                )
                                stats['annual_ratios_records'] += 1
                    
                    # Process Stock Price data
                    price_dates = [
                        '2025-08-19', '2025-06-30', '2025-03-28', '2024-12-31', '2024-09-30',
                        '2024-06-28', '2024-03-28', '2023-12-28', '2023-09-29', '2023-06-30'
                    ]
                    
                    for i, date_str in enumerate(price_dates):
                        if (price_start_col + i < len(row) and pe_start_col + i < len(row)):
                            price = parse_numeric_value(row.iloc[price_start_col + i])
                            pe = parse_numeric_value(row.iloc[pe_start_col + i])
                            
                            if price is not None or pe is not None:
                                try:
                                    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                                    StockPrice.objects.update_or_create(
                                        stock=stock,
                                        price_date=date_obj,
                                        defaults={
                                            'share_price': price,
                                            'pe_ratio': pe,
                                        }
                                    )
                                    stats['price_records'] += 1
                                except ValueError:
                                    continue
                    
                except Exception as e:
                    print(f"Error processing row {index}: {str(e)}")
                    stats['skipped_rows'] += 1
                    continue
        
        return stats
        
    except Exception as e:
        raise ValueError(f"Error processing stocks base sheet: {str(e)}")

@login_required
def process_amcfs_nav_and_returns(request):
    """
    View to process either TopSchemes data or RATIOS PE data 
    based on the selected file type in the form
    """
    form = MasterDataExcelUploadForm()
    
    if request.method == 'POST' and request.FILES.get('excel_file'):
        form = MasterDataExcelUploadForm(request.POST, request.FILES)
        
        if form.is_valid():
            excel_file = request.FILES['excel_file']
            file_type = form.cleaned_data['file_type']

            # Check if the file is an Excel file
            if not excel_file.name.endswith(('.xls', '.xlsx')):
                messages.error(request, "The uploaded file is not an Excel file. Please upload a valid Excel file.")
                return render(request, 'gcia_app/upload_amcfs_excel.html', {'form': form})

            try:
                if file_type == 'top_schemes':
                    # Process Top Schemes File
                    stats = process_top_schemes_data_uploaded(excel_file)
                elif file_type == 'ratios_pe':
                    # Process RATIOS PE file
                    stats = process_ratios_pe_file(excel_file)
                elif file_type == "index_nav":
                    # Process INDEX NAV file
                    stats = process_index_nav_file(excel_file)
                    try:
                        pe_ratio = get_bse500_pe_ratio()
                        if pe_ratio:
                            benchmark_index = AMCFundScheme.objects.filter(name="BSE 500 - TRI").first()
                            benchmark_index.pe_ratio = pe_ratio
                            benchmark_index.save()
                        messages.success(request, "BSE 500 TRI PE Updated Successfully")
                    except Exception as e:
                        print(str(e))
                        print(traceback.format_exc())
                elif file_type == 'stocks_base_sheet':
                    # Process Stocks Base Sheet file
                    stats = process_stocks_base_sheet(excel_file)
                    messages.success(request, f"Stocks data processed successfully! Created: {stats['stocks_created']}, Updated: {stats['stocks_updated']}, Total Records: {stats['market_cap_records'] + stats['ttm_records'] + stats['quarterly_records'] + stats['annual_ratios_records'] + stats['price_records']}")
                elif file_type == 'underlying_stocks':
                    # Process Underlying Stocks Holdings file
                    stats = process_underlying_stocks_file(excel_file)
                    success_msg = (
                        f"Holdings processed successfully! "
                        f"📊 Total: {stats['total_rows_in_file']} rows, "
                        f"✅ Processed: {stats['unique_records_processed']} unique, "
                        f"🔄 Duplicates removed: {stats['duplicates_removed']}, "
                        f"🏛️ Schemes: {stats['schemes_created']} created + {stats['schemes_updated']} updated, "
                        f"📈 Stocks: {stats['stocks_created']} created + {stats['stocks_matched']} matched, "
                        f"🎯 Holdings: {stats['holdings_created']} created"
                    )
                    if stats.get('holdings_failed', 0) > 0 or stats.get('batch_errors', 0) > 0:
                        success_msg += f", ⚠️ {stats.get('holdings_failed', 0)} failed"
                    messages.success(request, success_msg)

                messages.success(request, "File uploaded and data processed successfully!")

            except Exception as e:
                messages.error(request, f"Error processing file: {str(e)}")
                return render(request, 'gcia_app/upload_amcfs_excel.html', {'form': form})

    return render(request, 'gcia_app/upload_amcfs_excel.html', {'form': form})

def get_concentration_of_scheme(scheme):
    if not scheme.number_of_underlying_stocks:
        return "-"
    elif scheme.number_of_underlying_stocks < 40:
        return "High"
    elif scheme.number_of_underlying_stocks >= 40 and scheme.number_of_underlying_stocks <= 60:
        return "Medium"
    elif scheme.number_of_underlying_stocks > 60:
        return "Low"
    
def evaluate_fund_performance(scheme, index):
    """
    Evaluates fund performance against index and category averages using Django querysets
    
    Parameters:
    scheme (QuerySet object): Django QuerySet object for the fund scheme
    index (QuerySet object): Django QuerySet object for the index
        
    Returns:
    tuple: (performance_rating, valuation_status)
        - performance_rating: "High", "Medium", or "Low"
        - valuation_status: "Undervalued", "Fairly valued", or "Overvalued"
    """
    # Replace None values with 0 for all required fields
    scheme_returns_1_yr = scheme.returns_1_yr if scheme.returns_1_yr else 0
    scheme_returns_3_yr = scheme.returns_3_yr if scheme.returns_3_yr else 0
    scheme_returns_5_yr = scheme.returns_5_yr if scheme.returns_5_yr else 0
    
    scheme_fund_class_avg_1_yr = scheme.fund_class_avg_1_yr_returns if scheme.fund_class_avg_1_yr_returns else 0
    scheme_fund_class_avg_3_yr = scheme.fund_class_avg_3_yr_returns if scheme.fund_class_avg_3_yr_returns else 0
    scheme_fund_class_avg_5_yr = scheme.fund_class_avg_5_yr_returns if scheme.fund_class_avg_5_yr_returns else 0
    
    index_returns_1_yr = index.returns_1_yr if index.returns_1_yr else 0
    index_returns_3_yr = index.returns_3_yr if index.returns_3_yr else 0
    index_returns_5_yr = index.returns_5_yr if index.returns_5_yr else 0
    
    scheme_pe_ratio = scheme.pe_ratio if scheme.pe_ratio else 0
    index_pe_ratio = index.pe_ratio if index.pe_ratio else 0
    
    # Count periods where fund beats both index and category
    periods_beating_both = 0
    
    # 1-year comparison
    if (scheme_returns_1_yr > index_returns_1_yr and 
        scheme_returns_1_yr > scheme_fund_class_avg_1_yr):
        periods_beating_both += 1
    
    # 3-year comparison
    if (scheme_returns_3_yr > index_returns_3_yr and 
        scheme_returns_3_yr > scheme_fund_class_avg_3_yr):
        periods_beating_both += 1
    
    # 5-year comparison
    if (scheme_returns_5_yr > index_returns_5_yr and 
        scheme_returns_5_yr > scheme_fund_class_avg_5_yr):
        periods_beating_both += 1
    
    # Determine performance rating based on periods beating both index and category
    if periods_beating_both >= 2:
        performance_rating = "High"
    elif periods_beating_both >= 1:
        performance_rating = "Medium"
    else:
        performance_rating = "Low"
    
    # Determine valuation status based on PE ratio comparison
    # Handle special case where index PE is 0 to avoid division by zero
    if index_pe_ratio == 0:
        if scheme_pe_ratio == 0:
            valuation_status = "Fairly valued"  # Both are 0, consider them equal
        else:
            valuation_status = "Overvalued"  # Scheme has PE but index doesn't
    else:
        if scheme_pe_ratio < index_pe_ratio:
            valuation_status = "Undervalued"
        else:
            # Calculate percentage difference
            pe_difference_percentage = ((scheme_pe_ratio - index_pe_ratio) / index_pe_ratio) * 100
            
            if pe_difference_percentage > 10:
                valuation_status = "Overvalued"
            else:
                valuation_status = "Fairly valued"
    
    return performance_rating, valuation_status

@login_required
def process_portfolio_valuation(request):
    form = ExcelUploadForm()
    
    if request.method == 'POST' and request.FILES['excel_file']:
        excel_file = request.FILES['excel_file']

        # Check if the file is an Excel file
        if not excel_file.name.endswith(('.xls', '.xlsx')):
            messages.error(request, "The uploaded file is not an Excel file. Please upload a valid Excel file.")
            return render(request, 'gcia_app/portfolio_valuation.html', {'form': form})  # Return the form with error message

        try:
            # # Load the workbook and select the active sheet
            # wb = load_workbook(excel_file)
            # sheet = wb.active

            # # Initialize variables
            # headers = []
            # transaction_data = []
            # summary_data = []
            # current_scheme = None
            # scheme_folio = None
            # avg_nav = current_price = balance_units = None
            # client_details = {}
            # client_name = None

            # for row in sheet.iter_rows():
            #     values = [cell.value for cell in row]

            #     # Skip empty rows
            #     if all(v is None for v in values):
            #         continue

            #     # Detect headers for transactions
            #     if 'TRANSACTION TYPE' in values:
            #         headers = values
            #         continue

            #     if 'Client:' in values[0]:
            #         client_name, pan_number = values[0].replace("Client:","").replace(")", "").split("(")
            #         relationship_manager = values[2].split("Relationship Manager:")[1].strip()
            #         client_details["Client Name"] = client_name.strip()
            #         client_details["Client Pan"] = pan_number.strip()
            #         client_details["Relationship Manager"] = relationship_manager
                
            #     if 'Grand Total' in values[0]:
            #         client_details["Invested Amount"] = values[4]
            #         client_details["Current Value"] = values[6]
            #         client_details["Dividend"] = values[7]
            #         client_details["Gain"] = values[8]
            #         client_details["Holding Days"] = values[9]
            #         client_details["Absolute Return"] = values[10]
            #         client_details["CAGR(%)"] = values[11]

            #     # Detect scheme summaries (rows with 'Total' and not just category totals)
            #     if values[0] and 'Total' in values[0] and 'Fund' in values[0] and current_scheme:
            #         scheme_name = current_scheme
            #         summary_data.append({
            #             "Scheme": scheme_name,
            #             "Folio": scheme_folio,
            #             "Total Investment": values[4],
            #             "Current Value": values[6],
            #             "Gain": values[8],
            #             "Holding Days": values[9],
            #             "Absolute Return (%)": values[10],
            #             "CAGR (%)": values[11],
            #             "Avg. Purchase NAV": avg_nav,
            #             "Current Price": current_price,
            #             "Balance Units": balance_units,
            #         })
            #         continue

            #     # Capture transaction rows
            #     if headers and current_scheme:
            #         if any(values) and not 'Total' in values[0] and not 'Avg. Purchase' in values[0] and values[2]:  # Skip summary rows
            #             transaction_data.append({
            #                 "Scheme": current_scheme,
            #                 "Folio": scheme_folio,
            #                 **dict(zip(headers, values))
            #             })

            #     # Detect scheme names (rows containing the scheme details)
            #     if values[0] and "[" in values[0]:  # Assuming scheme names have this format
            #         current_scheme = values[0].split("[")[0].strip()
            #         scheme_folio = values[0].split("[")[-1].strip().split(" ")[0].strip()
                    

            #     # Capture Avg. Purchase NAV, Current Price, and Balance Units
            #     if values[0] and 'Avg. Purchase NAV' in values[0]:
            #         avg_nav = values[0].split(":")[1].split(",")[0].strip()
            #         current_price = values[0].split("Current Price:")[1].split(",")[0].strip()
            #         balance_units = values[0].split("Balance Units:")[1].strip()

            wb = load_workbook(excel_file)
            
            # Results containers
            all_clients = {}
            
            # Process each sheet (focusing on Mutual Fund sheet first)
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                
                # Skip processing if not the Mutual Fund sheet for now
                if not sheet_name.endswith("Mutual Fund"):
                    continue
                    
                # Initialize tracking variables
                headers = []
                current_client = None
                current_category = None
                current_scheme = None
                scheme_folio = None
                avg_nav = current_price = balance_units = None
                in_transaction_section = False
                
                # Process each row
                for row_idx, row in enumerate(sheet.iter_rows(), 1):
                    values = [cell.value for cell in row]
                    
                    # Skip empty rows
                    if all(v is None or v == "" for v in values):
                        continue
                        
                    # Detect client rows
                    if values[0] and "Client:" in str(values[0]):
                        client_text = str(values[0]).strip()
                        # Extract client name and PAN
                        if "(" in client_text and ")" in client_text:
                            client_parts = client_text.replace("Client:", "").strip()
                            client_name = client_parts.split("(")[0].strip()
                            pan_number = client_parts.split("(")[1].split(")")[0].strip()
                            
                            # Set as current client
                            current_client = client_name
                            
                            # Initialize client data if not exists
                            if current_client not in all_clients:
                                all_clients[current_client] = {
                                    "client_details": {
                                        "Client Name": client_name,
                                        "Client Pan": pan_number,
                                    },
                                    "summary_data": [],
                                    "transaction_data": []
                                }
                            
                            # Extract relationship manager
                            for cell_val in values:
                                if cell_val and "Relationship Manager:" in str(cell_val):
                                    manager = str(cell_val).split("Relationship Manager:")[1].strip()
                                    all_clients[current_client]["client_details"]["Relationship Manager"] = manager
                                    break
                        
                    # When we see a client name in a row by itself, switch to that client
                    elif values[0] and "(" in str(values[0]) and ")" in str(values[0]) and all(v is None or v == "" for v in values[1:]):
                        client_text = str(values[0]).strip()
                        for client_name in all_clients.keys():
                            if client_name in client_text:
                                current_client = client_name
                                break
                    
                    # Detect headers for transactions
                    if values[0] and "TRANSACTION TYPE" in str(values[0]):
                        headers = values
                        in_transaction_section = True
                        continue
                        
                    # Process Grand Total for current client
                    if values[0] and "Grand Total" in str(values[0]):
                        if current_client:
                            client_data = all_clients[current_client]
                            client_data["client_details"]["Invested Amount"] = values[4]
                            client_data["client_details"]["Current Value"] = values[6]
                            client_data["client_details"]["Dividend"] = values[7]
                            client_data["client_details"]["Gain"] = values[8]
                            client_data["client_details"]["Holding Days"] = values[9]
                            client_data["client_details"]["Absolute Return"] = values[10]
                            client_data["client_details"]["CAGR(%)"] = values[11]
                    
                    # Process individual client totals
                    if current_client and values[0] and str(values[0]).endswith("Total") and current_client in str(values[0]):
                        client_data = all_clients[current_client]
                        if "Total Investment" not in client_data["client_details"]:
                            client_data["client_details"]["Total Investment"] = values[4]
                            client_data["client_details"]["Current Value"] = values[6]
                            client_data["client_details"]["Gain"] = values[8]
                            client_data["client_details"]["Holding Days"] = values[9]
                            client_data["client_details"]["Absolute Return"] = values[10]
                            client_data["client_details"]["CAGR(%)"] = values[11]
                    
                    # Detect category rows (e.g., "Equity: ELSS")
                    if values[0] and ":" in str(values[0]) and not "[" in str(values[0]) and not "NAV:" in str(values[0]) and not "Client:" in str(values[0]):
                        current_category = str(values[0]).strip()
                        continue
                        
                    # Detect scheme rows
                    if values[0] and "[" in str(values[0]):
                        scheme_text = str(values[0]).strip()
                        current_scheme = scheme_text.split("[")[0].strip()
                        
                        # Extract folio number (handle cases with 'Comments:')
                        folio_section = scheme_text.split("[")[1].strip()
                        if "Comments:" in folio_section:
                            scheme_folio = folio_section.split("Comments:")[0].strip()
                            comments = folio_section.split("Comments:")[1].strip()
                        else:
                            scheme_folio = folio_section.rstrip("] ")
                            comments = ""
                            
                        # Ensure we're not capturing "]" at the end
                        scheme_folio = scheme_folio.rstrip("]").strip()
                        continue
                        
                    # Capture Avg. Purchase NAV, Current Price, and Balance Units
                    if values[0] and "Avg. Purchase NAV" in str(values[0]):
                        nav_text = str(values[0]).strip()
                        
                        # Extract the three values using regex
                        avg_nav_match = re.search(r'Avg\. Purchase NAV: ([\d\.]+)', nav_text)
                        current_price_match = re.search(r'Current Price: ([\d\.]+)', nav_text)
                        balance_units_match = re.search(r'Balance Units: ([\d\.]+)', nav_text)
                        
                        if avg_nav_match:
                            avg_nav = avg_nav_match.group(1)
                        if current_price_match:
                            current_price = current_price_match.group(1)
                        if balance_units_match:
                            balance_units = balance_units_match.group(1)
                        continue
                        
                    # Detect scheme totals (e.g., "Axis ELSS Tax Saver Fund (G) Total")
                    if values[0] and "Total" in str(values[0]) and current_scheme and current_scheme in str(values[0]):
                        if current_client and current_scheme:
                            all_clients[current_client]["summary_data"].append({
                                "Category": current_category,
                                "Scheme": current_scheme,
                                "Folio": scheme_folio,
                                "Total Investment": values[4],
                                "Current Value": values[6],
                                "Gain": values[8],
                                "Holding Days": values[9],
                                "Absolute Return (%)": values[10],
                                "CAGR (%)": values[11],
                                "Avg. Purchase NAV": avg_nav,
                                "Current Price": current_price,
                                "Balance Units": balance_units,
                            })
                        continue
                        
                    # Process transaction rows
                    if (headers and current_client and current_scheme and in_transaction_section and 
                        values[0] and any(x in str(values[0]) for x in ["Purchase", "SIP", "Switch In", "Redemption", "Switch Out"])):
                        
                        # Create a dictionary with header names as keys
                        transaction = {}
                        for i, header in enumerate(headers):
                            if i < len(values):
                                transaction[header] = values[i]
                            else:
                                transaction[header] = None
                                
                        # Add scheme info to transaction
                        transaction["Scheme"] = current_scheme
                        transaction["Folio"] = scheme_folio
                        transaction["Category"] = current_category
                        
                        # Add to the client's transactions
                        all_clients[current_client]["transaction_data"].append(transaction)
            
            for client_name, client_data in all_clients.items():
                print(f"Client: {client_name}")
                print(f"Details: {client_data['client_details']}")
                print(f"Number of schemes: {len(client_data['summary_data'])}")
                print(f"Number of transactions: {len(client_data['transaction_data'])}")

            client_details = client_data['client_details']
            summary_data = client_data['summary_data']
            transaction_data = client_data['transaction_data']

            ppt_data = []

            benchmark_index = AMCFundScheme.objects.filter(name="BSE 500 - TRI").first()

            for fund_data in summary_data:
                amcfs = AMCFundScheme.objects.filter(name=fund_data["Scheme"]).first()
                if amcfs:
                    concentration = get_concentration_of_scheme(amcfs)
                    quality, valuation = evaluate_fund_performance(amcfs, benchmark_index)

                    fund_data["Fund Category"] = amcfs.fund_class
                    fund_data["Is Direct Fund"] = amcfs.is_direct_fund
                    fund_data["Concentration"] = concentration
                    fund_data["Price"] = valuation
                    fund_data["Quality"] = quality
                    fund_data["Performance"] = quality

                    fund_data["No of Securities"] = int(amcfs.number_of_underlying_stocks) if amcfs.number_of_underlying_stocks else 0
                    fund_data["Age of the fund"] = calculate_scheme_age(amcfs.launch_date) if amcfs.launch_date else 0
                    fund_data["Equity Allocation"] = round(amcfs.equity_percentage, 1) if amcfs.equity_percentage else 0
                    fund_data["Debt Allocation"] = round(amcfs.debt_percentage, 1) if amcfs.debt_percentage else 0
                    fund_data["Cash Allocation"] = round(amcfs.cash, 1) if amcfs.cash else 0
                    fund_data["Expense ratio"] = round(amcfs.expense_ratio, 1) if amcfs.expense_ratio else 0
                    fund_data["Credit rating AAA"] = round(amcfs.AAA, 1) if amcfs.AAA else 0
                    fund_data["Credit rating AA"] = round(amcfs.AA, 1) if amcfs.AA else 0
                    fund_data["Credit rating A"] = round(amcfs.A, 1) if amcfs.A else 0
                    fund_data["Up Capture"] = round(amcfs.up_capture, 1) if amcfs.up_capture else 0
                    fund_data["Down Capture"] = round(amcfs.down_capture, 1) if amcfs.down_capture else 0

                    fund_data["Mod duration"] = round(amcfs.modified_duration, 1) if amcfs.modified_duration else 0
                    fund_data["YTM"] = round(amcfs.ytm, 1) if amcfs.ytm else 0

                    # benchmark_index = None
                    # if amcfs.scheme_benchmark:
                    #     benchmark_index = AMCFundScheme.objects.filter(name=amcfs.scheme_benchmark).first()
                    # if not benchmark_index:
                    #     benchmark_index = AMCFundScheme.objects.filter(name="BSE 500 TRI").first()
                    ppt_data.append({
                    "Scheme Name": [
                        amcfs.name,
                        benchmark_index.name,
                        amcfs.fund_class
                    ],
                    "Purchase Value": round(fund_data["Total Investment"], 1),
                    "Current Value": round(fund_data["Current Value"], 1),
                    "Gain": round(fund_data["Gain"], 1),
                    "Weight": round((fund_data["Current Value"]/client_details["Current Value"])*100,1),
                    "Stocks": int(amcfs.number_of_underlying_stocks) if amcfs.number_of_underlying_stocks else 0,
                    "SI(%)": [
                        round(amcfs.returns_from_launch, 1) if amcfs.returns_from_launch else 0,
                        round(benchmark_index.returns_from_launch, 1) if benchmark_index.returns_from_launch else 0,
                        round(amcfs.fund_class_avg_returns_from_launch, 1) if amcfs.fund_class_avg_returns_from_launch else 0
                    ],
                    "1YR(%)": [
                        round(amcfs.returns_1_yr, 1) if amcfs.returns_1_yr else 0,
                        round(benchmark_index.returns_1_yr, 1) if benchmark_index.returns_1_yr else 0,
                        round(amcfs.fund_class_avg_1_yr_returns, 1) if amcfs.fund_class_avg_1_yr_returns else 0
                    ],
                    "3YR(%)": [
                        round(amcfs.returns_3_yr, 1) if amcfs.returns_3_yr else 0,
                        round(benchmark_index.returns_3_yr, 1) if benchmark_index.returns_3_yr else 0,
                        round(amcfs.fund_class_avg_3_yr_returns, 1) if amcfs.fund_class_avg_3_yr_returns else 0
                    ],
                    "5YR(%)": [
                        round(amcfs.returns_5_yr, 1) if amcfs.returns_5_yr else 0,
                        round(benchmark_index.returns_5_yr, 1) if benchmark_index.returns_5_yr else 0,
                        round(amcfs.fund_class_avg_5_yr_returns, 1) if amcfs.fund_class_avg_5_yr_returns else 0
                    ],
                    "PE": round(amcfs.pe_ratio, 1) if amcfs.pe_ratio else 0,
                    "Index PE": round(benchmark_index.pe_ratio, 1) if benchmark_index.pe_ratio else 0,
                    "Comments": [
                        "Concentration: " + concentration,
                        "Performance: " + quality,
                        "Valuation: " + valuation
                    ]
                })
                    
            file = create_fund_presentation(ppt_data, client_details, summary_data, transaction_data)
            
            # Create response for file download
            if os.path.exists(file):
                with open(file, 'rb') as f:
                    response = HttpResponse(
                        f.read(),
                        content_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
                    )
                    response['Content-Disposition'] = f'attachment; filename="{os.path.basename(file)}"'
                
                # Clean up the file after sending
                try:
                    os.remove(file)
                except Exception as e:
                    print(f"Error removing file: {e}")
                
                return response
            
            messages.success(request, "PPT File Generated Successfully")
            return render(request, 'gcia_app/portfolio_valuation.html', {'form': form})
        except Exception as e:
            print(traceback.format_exc())
            messages.error(request, f"Error processing file: {e}")
            return render(request, 'gcia_app/portfolio_valuation.html')

    return render(request, 'gcia_app/portfolio_valuation.html', {'form': form})

@login_required
def process_financial_planning(request):
    return render(request, 'gcia_app/financial_planning.html')

@login_required
def fund_analysis_metrics_view(request):
    """
    Display page with dropdown to select mutual funds for analysis
    """
    # Get all active mutual fund schemes
    funds = AMCFundScheme.objects.filter(is_active=True).order_by('name')

    context = {
        'funds': funds,
        'page_title': 'Fund Analysis Metrics'
    }

    return render(request, 'gcia_app/fund_analysis_metrics.html', context)

@login_required
def download_fund_metrics(request, scheme_id):
    """
    Download Excel file with fund holdings integrated into stock structure (Portfolio Analysis format)
    Uses complete 431-column structure: ALL 426 stock columns + 5 fund columns
    Preserves every single stock data column without loss
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from django.utils import timezone
    from io import BytesIO
    from .header_mapping import get_full_fund_integrated_headers, get_complete_fund_column_mapping
    from datetime import datetime

    try:
        # Get the selected fund scheme
        scheme = AMCFundScheme.objects.get(amcfundscheme_id=scheme_id, is_active=True)

        # Get fund holdings with related stock data
        holdings = FundHolding.objects.filter(scheme=scheme).select_related('stock').prefetch_related(
            'stock__market_cap_data', 'stock__ttm_data', 'stock__quarterly_data',
            'stock__annual_ratios', 'stock__price_data'
        ).order_by('-holding_percentage')

        if not holdings.exists():
            messages.error(request, f"No holdings data found for {scheme.name}")
            return redirect('fund_analysis_metrics')

        # Create workbook with Portfolio Analysis format
        wb = Workbook()
        ws = wb.active
        ws.title = "Portfolio Analysis"  # Generic title, not fund-specific

        # Get COMPLETE header structure (431 columns: ALL stock data + fund columns)
        integrated_headers = get_full_fund_integrated_headers()
        column_mapping = get_complete_fund_column_mapping()

        # ROW 1: Fund name only
        row_1 = integrated_headers['row_1'].copy()
        row_1[0] = scheme.name
        ws.append(row_1)

        # ROW 2: Column numbers (professional format)
        ws.append(integrated_headers['row_2'])

        # ROW 3: Portfolio date + header content
        row_3 = integrated_headers['row_3'].copy()
        portfolio_date = holdings.first().holding_date if holdings.first().holding_date else datetime.now().date()
        row_3[0] = f"Portfolio as on: {portfolio_date.strftime('%dth %B %Y')}"
        ws.append(row_3)

        # ROW 4-7: Professional header structure (ALL preserved)
        ws.append(integrated_headers['row_4'])
        ws.append(integrated_headers['row_5'])
        ws.append(integrated_headers['row_6'])
        ws.append(integrated_headers['row_7'])

        # ROW 8: Build comprehensive column headers (431 columns)
        headers_row = [''] * 431

        # Fund-integrated basic columns (0-16)
        fund_columns = column_mapping['fund_integrated_columns']
        for i, header in enumerate(fund_columns):
            if i < len(headers_row):
                headers_row[i] = header

        # Market cap dates - ALL preserved
        market_cap_dates = column_mapping['market_cap_dates']
        start_pos = column_mapping['market_cap_start']
        for i, date in enumerate(market_cap_dates):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = date

        # Market cap free float dates - ALL preserved
        start_pos = column_mapping['market_cap_ff_start']
        for i, date in enumerate(market_cap_dates):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = date

        # TTM periods for various sections - ALL preserved
        ttm_periods = column_mapping['ttm_periods']

        # TTM Revenue - ALL periods preserved
        start_pos = column_mapping['ttm_revenue_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        # TTM Revenue Free Float - ALL periods preserved
        start_pos = column_mapping['ttm_revenue_ff_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        # TTM PAT - ALL periods preserved
        start_pos = column_mapping['ttm_pat_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        # TTM PAT Free Float - ALL periods preserved
        start_pos = column_mapping['ttm_pat_ff_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        # Quarterly sections - ALL periods preserved
        start_pos = column_mapping['qtr_revenue_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        start_pos = column_mapping['qtr_revenue_ff_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        start_pos = column_mapping['qtr_pat_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        start_pos = column_mapping['qtr_pat_ff_start']
        for i, period in enumerate(ttm_periods):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = period

        # Annual ratios - ALL years preserved
        annual_years = column_mapping['annual_years']

        # ROCE - ALL years preserved
        start_pos = column_mapping['roce_start']
        for i, year in enumerate(annual_years):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = year

        # ROE - ALL years preserved
        start_pos = column_mapping['roe_start']
        for i, year in enumerate(annual_years):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = year

        # Retention - ALL years preserved
        start_pos = column_mapping['retention_start']
        for i, year in enumerate(annual_years):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = year

        # Price data - ALL dates preserved
        price_dates = column_mapping['price_dates']

        # Share prices - ALL dates preserved
        start_pos = column_mapping['share_price_start']
        for i, date in enumerate(price_dates):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = date

        # PE ratios - ALL dates preserved
        start_pos = column_mapping['pe_ratio_start']
        for i, date in enumerate(price_dates):
            pos = start_pos + i
            if pos < len(headers_row):
                headers_row[pos] = f"PE_{date}"

        # Identifiers - ALL preserved
        start_pos = column_mapping['identifiers_start']
        if start_pos < len(headers_row):
            headers_row[start_pos] = 'BSE Code'
        if start_pos + 1 < len(headers_row):
            headers_row[start_pos + 1] = 'NSE Code'
        if start_pos + 2 < len(headers_row):
            headers_row[start_pos + 2] = 'ISIN'

        ws.append(headers_row)

        # Add fund holding data with integrated stock metrics
        for holding in holdings:
            stock = holding.stock
            row_data = [''] * 431  # 431 columns: 426 stock + 5 fund

            # Calculate fund-specific values
            weights = (holding.holding_percentage / 100) if holding.holding_percentage else 0
            value = holding.market_value if holding.market_value else 0
            num_shares = holding.number_of_shares if holding.number_of_shares else 0

            # Get latest stock price
            latest_price = stock.price_data.order_by('-price_date').first()
            share_price = latest_price.share_price if latest_price else 0

            # Calculate market cap (if available from stock data)
            latest_market_cap = stock.market_cap_data.order_by('-date').first()
            market_cap = latest_market_cap.market_cap if latest_market_cap else 0

            # Calculate factor (example formula - adjust as needed)
            factor = 0
            if market_cap and market_cap > 0:
                factor = weights / (market_cap / 1000000)  # Adjust scale as needed

            # Populate fund-integrated basic columns (0-16)
            row_data[0] = stock.company_name              # Company Name
            row_data[1] = stock.accord_code               # Accord Code
            row_data[2] = stock.sector                    # Sector
            row_data[3] = stock.cap                       # Cap
            row_data[4] = market_cap                      # Market Cap (FUND INTEGRATED)
            row_data[5] = weights                         # Weights (FUND INTEGRATED)
            row_data[6] = factor                          # Factor (FUND INTEGRATED)
            row_data[7] = value                           # Value (FUND INTEGRATED)
            row_data[8] = num_shares                      # No.of shares (FUND INTEGRATED)
            row_data[9] = stock.free_float                # Free Float (shifted by +5)
            row_data[10] = stock.revenue_6yr_cagr         # 6 Year CAGR Revenue (shifted by +5)
            row_data[11] = stock.revenue_ttm              # TTM Revenue (shifted by +5)
            row_data[12] = stock.pat_6yr_cagr             # 6 Year CAGR PAT (shifted by +5)
            row_data[13] = stock.pat_ttm                  # TTM PAT (shifted by +5)
            row_data[14] = stock.current_value            # Current (shifted by +5)
            row_data[15] = stock.two_yr_avg               # 2 Yr Avg (shifted by +5)
            row_data[16] = stock.reval_deval              # Reval/deval (shifted by +5)

            # Populate ALL stock data - NO LIMITS, ALL PRESERVED
            # Market cap data - ALL dates preserved
            market_cap_data = {}
            for mc in stock.market_cap_data.all():
                market_cap_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap

            start_pos = column_mapping['market_cap_start']
            for idx, date_str in enumerate(market_cap_dates):
                pos = start_pos + idx
                if pos < len(row_data):
                    row_data[pos] = market_cap_data.get(date_str)

            # Market cap free float data - ALL dates preserved
            market_cap_ff_data = {}
            for mc in stock.market_cap_data.all():
                market_cap_ff_data[mc.date.strftime('%Y-%m-%d')] = mc.market_cap_free_float

            start_pos = column_mapping['market_cap_ff_start']
            for idx, date_str in enumerate(market_cap_dates):
                pos = start_pos + idx
                if pos < len(row_data):
                    row_data[pos] = market_cap_ff_data.get(date_str)

            # TTM data - ALL periods preserved
            ttm_data = {}
            for ttm in stock.ttm_data.all():
                ttm_data[ttm.period] = ttm

            # TTM Revenue - ALL periods preserved
            start_pos = column_mapping['ttm_revenue_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    ttm_obj = ttm_data.get(period)
                    row_data[pos] = ttm_obj.ttm_revenue if ttm_obj else None

            # TTM Revenue Free Float - ALL periods preserved
            start_pos = column_mapping['ttm_revenue_ff_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    ttm_obj = ttm_data.get(period)
                    row_data[pos] = ttm_obj.ttm_revenue_free_float if ttm_obj else None

            # TTM PAT - ALL periods preserved
            start_pos = column_mapping['ttm_pat_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    ttm_obj = ttm_data.get(period)
                    row_data[pos] = ttm_obj.ttm_pat if ttm_obj else None

            # TTM PAT Free Float - ALL periods preserved
            start_pos = column_mapping['ttm_pat_ff_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    ttm_obj = ttm_data.get(period)
                    row_data[pos] = ttm_obj.ttm_pat_free_float if ttm_obj else None

            # Quarterly data - ALL periods preserved
            quarterly_data = {}
            for qtr in stock.quarterly_data.all():
                quarterly_data[qtr.period] = qtr

            # Quarterly Revenue - ALL periods preserved
            start_pos = column_mapping['qtr_revenue_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    qtr_obj = quarterly_data.get(period)
                    row_data[pos] = qtr_obj.quarterly_revenue if qtr_obj else None

            # Quarterly Revenue Free Float - ALL periods preserved
            start_pos = column_mapping['qtr_revenue_ff_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    qtr_obj = quarterly_data.get(period)
                    row_data[pos] = qtr_obj.quarterly_revenue_free_float if qtr_obj else None

            # Quarterly PAT - ALL periods preserved
            start_pos = column_mapping['qtr_pat_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    qtr_obj = quarterly_data.get(period)
                    row_data[pos] = qtr_obj.quarterly_pat if qtr_obj else None

            # Quarterly PAT Free Float - ALL periods preserved
            start_pos = column_mapping['qtr_pat_ff_start']
            for idx, period in enumerate(ttm_periods):
                pos = start_pos + idx
                if pos < len(row_data):
                    qtr_obj = quarterly_data.get(period)
                    row_data[pos] = qtr_obj.quarterly_pat_free_float if qtr_obj else None

            # Annual ratios - ALL years preserved
            annual_data = {}
            for ratio in stock.annual_ratios.all():
                annual_data[ratio.financial_year] = ratio

            # ROCE - ALL years preserved
            start_pos = column_mapping['roce_start']
            for idx, year in enumerate(annual_years):
                pos = start_pos + idx
                if pos < len(row_data):
                    ratio_obj = annual_data.get(year)
                    row_data[pos] = ratio_obj.roce_percentage if ratio_obj else None

            # ROE - ALL years preserved
            start_pos = column_mapping['roe_start']
            for idx, year in enumerate(annual_years):
                pos = start_pos + idx
                if pos < len(row_data):
                    ratio_obj = annual_data.get(year)
                    row_data[pos] = ratio_obj.roe_percentage if ratio_obj else None

            # Retention - ALL years preserved
            start_pos = column_mapping['retention_start']
            for idx, year in enumerate(annual_years):
                pos = start_pos + idx
                if pos < len(row_data):
                    ratio_obj = annual_data.get(year)
                    row_data[pos] = ratio_obj.retention_percentage if ratio_obj else None

            # Price data - ALL dates preserved
            price_data = {}
            pe_data = {}
            for price in stock.price_data.all():
                price_data[price.price_date.strftime('%Y-%m-%d')] = price.share_price
                pe_data[price.price_date.strftime('%Y-%m-%d')] = price.pe_ratio

            # Share prices - ALL dates preserved
            start_pos = column_mapping['share_price_start']
            for idx, date_str in enumerate(price_dates):
                pos = start_pos + idx
                if pos < len(row_data):
                    row_data[pos] = price_data.get(date_str)

            # PE ratios - ALL dates preserved
            start_pos = column_mapping['pe_ratio_start']
            for idx, date_str in enumerate(price_dates):
                pos = start_pos + idx
                if pos < len(row_data):
                    row_data[pos] = pe_data.get(date_str)

            # Identifiers - ALL preserved
            start_pos = column_mapping['identifiers_start']
            if start_pos < len(row_data):
                row_data[start_pos] = stock.bse_code
            if start_pos + 1 < len(row_data):
                row_data[start_pos + 1] = stock.nse_symbol
            if start_pos + 2 < len(row_data):
                row_data[start_pos + 2] = stock.isin

            ws.append(row_data)

        # Save to BytesIO buffer
        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        excel_content = excel_buffer.getvalue()
        excel_buffer.close()

        # Create response with generic filename
        response = HttpResponse(
            excel_content,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        filename = f"portfolio_analysis_{timezone.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        return response

    except AMCFundScheme.DoesNotExist:
        messages.error(request, "Fund not found or inactive")
        return redirect('fund_analysis_metrics')
    except Exception as e:
        print(f"Error in download_fund_metrics: {e}")
        print(traceback.format_exc())
        messages.error(request, f"Error generating fund metrics file: {e}")
        return redirect('fund_analysis_metrics')



