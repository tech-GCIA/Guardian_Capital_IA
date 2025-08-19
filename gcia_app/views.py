# views.py
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.utils.timezone import now
from datetime import timedelta
from gcia_app.forms import CustomerCreationForm
from django.contrib import messages
import openpyxl
from django.http import HttpResponse, JsonResponse
from gcia_app.forms import ExcelUploadForm, MasterDataExcelUploadForm
from gcia_app.models import AMCFundScheme, AMCFundSchemeNavLog, Stock
import pandas as pd
import os
import datetime
from openpyxl import load_workbook
import traceback
from gcia_app.utils import update_avg_category_returns, calculate_scheme_age
from gcia_app.portfolio_analysis_ppt import create_fund_presentation
import re
from difflib import SequenceMatcher
from django.db import transaction
from gcia_app.index_scrapper_from_screener import get_bse500_pe_ratio
import json
from django.db.models import Q
from django.core.exceptions import ValidationError

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

# Add this to gcia_app/views.py

import os
import tempfile
from django.shortcuts import render, redirect
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.conf import settings
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from gcia_app.forms import StockDataUploadForm
from gcia_app.models import StockUploadLog
from gcia_app.stock_data_processor import process_stock_data_file
import traceback
import logging

logger = logging.getLogger(__name__)

@login_required
def upload_stock_data(request):
    """
    View to handle stock data Excel file upload and processing
    """
    form = StockDataUploadForm()
    recent_uploads = []
    
    # Get recent uploads for the current user
    if request.user.is_authenticated:
        recent_uploads = StockUploadLog.objects.filter(
            uploaded_by=request.user
        ).order_by('-uploaded_at')[:10]
    
    if request.method == 'POST' and request.FILES.get('excel_file'):
        form = StockDataUploadForm(request.POST, request.FILES)
        
        if form.is_valid():
            excel_file = request.FILES['excel_file']
            
            try:
                # Create uploads directory if it doesn't exist
                uploads_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
                os.makedirs(uploads_dir, exist_ok=True)
                
                # Save file temporarily
                file_path = default_storage.save(
                    f'uploads/{excel_file.name}',
                    ContentFile(excel_file.read())
                )
                
                try:
                    # Process the file
                    with default_storage.open(file_path, 'rb') as stored_file:
                        # Create a temporary file for processing
                        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
                            temp_file.write(stored_file.read())
                            temp_file.flush()
                            
                            # Reset file pointer
                            excel_file.seek(0)
                            
                            # Process the data
                            stats, upload_log = process_stock_data_file(excel_file, request.user)
                            
                            # Show success message with statistics
                            success_message = (
                                f"File processed successfully! "
                                f"Stocks added: {stats['stocks_added']}, "
                                f"Stocks updated: {stats['stocks_updated']}, "
                                f"Quarterly records added: {stats['quarterly_records_added']}, "
                                f"Quarterly records updated: {stats['quarterly_records_updated']}"
                            )
                            
                            if stats['errors']:
                                success_message += f" (with {len(stats['errors'])} warnings)"
                            
                            messages.success(request, success_message)
                            
                            # Log any warnings
                            if stats['errors']:
                                for error in stats['errors'][:5]:  # Show first 5 errors
                                    messages.warning(request, f"Warning: {error}")
                                
                                if len(stats['errors']) > 5:
                                    messages.info(request, f"... and {len(stats['errors']) - 5} more warnings")
                
                except Exception as e:
                    logger.error(f"Error processing stock data file: {str(e)}")
                    logger.error(traceback.format_exc())
                    messages.error(request, f"Error processing file: {str(e)}")
                
                finally:
                    # Clean up temporary files
                    try:
                        default_storage.delete(file_path)
                        if 'temp_file' in locals():
                            os.unlink(temp_file.name)
                    except Exception as cleanup_error:
                        logger.warning(f"Error cleaning up files: {cleanup_error}")
                
            except Exception as e:
                logger.error(f"Error handling file upload: {str(e)}")
                logger.error(traceback.format_exc())
                messages.error(request, f"Error uploading file: {str(e)}")
            
            # Redirect to prevent form resubmission
            return redirect('upload_stock_data')
    
    context = {
        'form': form,
        'recent_uploads': recent_uploads,
    }
    
    return render(request, 'gcia_app/upload_stock_data.html', context)

# Add this view to your existing gcia_app/views.py

# Replace your search_stocks view in gcia_app/views.py with this improved version:

@login_required
def search_stocks(request):
    """
    AJAX endpoint for stock search functionality with debugging
    """
    query = request.GET.get('q', '').strip()
    
    print(f"Search request received: query='{query}'")  # Debug log
    
    if len(query) < 2:
        print("Query too short, returning empty results")  # Debug log
        return JsonResponse({'stocks': []})
    
    try:
        # Search in both name and symbol
        stocks = Stock.objects.filter(
            Q(name__icontains=query) | Q(symbol__icontains=query),
            is_active=True
        ).order_by('name')[:20]  # Limit to 20 results for performance
        
        print(f"Found {stocks.count()} stocks matching '{query}'")  # Debug log
        
        stock_data = [
            {
                'id': stock.stock_id,
                'name': stock.name,
                'symbol': stock.symbol,
                'sector': stock.sector or '',
                'display_name': f"{stock.name} ({stock.symbol})"
            }
            for stock in stocks
        ]
        
        print(f"Returning {len(stock_data)} stock results")  # Debug log
        return JsonResponse({'stocks': stock_data})
        
    except Exception as e:
        print(f"Error in search_stocks: {e}")  # Debug log
        return JsonResponse({'stocks': [], 'error': str(e)})

# Also update your fund_analysis view to include debug info:

@login_required
def fund_analysis(request):
    """
    Fund Analysis page with simple stock selection and table management
    """
    # Get all active stocks for the search dropdown
    all_stocks = Stock.objects.filter(is_active=True).order_by('name')
    
    print(f"Loading fund_analysis page with {all_stocks.count()} active stocks")  # Debug log
    
    context = {
        'all_stocks': all_stocks,
        'all_stocks_json': json.dumps([
            {
                'id': stock.stock_id,
                'name': stock.name,
                'symbol': stock.symbol,
                'sector': stock.sector or '',
                'display_name': f"{stock.name} ({stock.symbol})"
            }
            for stock in all_stocks
        ])
    }
    
    return render(request, 'gcia_app/fund_analysis.html', context)

# Add this to your existing gcia_app/views.py file
import os
import json
import tempfile
from django.http import HttpResponse, JsonResponse
from django.contrib.auth.decorators import login_required
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods
from gcia_app.fund_report_generator import FundReportGenerator
import logging
from datetime import datetime
logger = logging.getLogger(__name__)

# Replace the existing generate_fund_report_simple function with this:

@login_required
@require_http_methods(["POST"])
def generate_fund_report_simple(request):
    """
    Generate and download Excel report directly when "Generate Report" is clicked
    Fixed version that properly handles .xlsx format
    """
    try:
        # Parse JSON data from request
        data = json.loads(request.body)
        fund_name = data.get('fund_name', '').strip()
        selected_stocks = data.get('selected_stocks', [])
        
        # Validation
        if not fund_name:
            return JsonResponse({
                'success': False,
                'message': 'Please enter a fund name'
            })
        
        if not selected_stocks:
            return JsonResponse({
                'success': False,
                'message': 'Please select at least one stock'
            })
        
        # Validate stock data
        total_weightage = 0
        for stock in selected_stocks:
            try:
                weightage = float(stock.get('weightage', 0))
                shares = int(stock.get('shares', 0))
                
                if weightage <= 0:
                    return JsonResponse({
                        'success': False,
                        'message': f'Invalid weightage for stock {stock.get("name", "Unknown")}: {weightage}'
                    })
                
                if shares <= 0:
                    return JsonResponse({
                        'success': False,
                        'message': f'Invalid number of shares for stock {stock.get("name", "Unknown")}: {shares}'
                    })
                
                total_weightage += weightage
                
            except (ValueError, TypeError):
                return JsonResponse({
                    'success': False,
                    'message': f'Invalid numeric data for stock {stock.get("name", "Unknown")}'
                })
        
        # Check total weightage
        if total_weightage > 100:
            return JsonResponse({
                'success': False,
                'message': f'Total weightage ({total_weightage:.2f}%) cannot exceed 100%'
            })
        
        if total_weightage <= 0:
            return JsonResponse({
                'success': False,
                'message': 'Total weightage must be greater than 0%'
            })
        
        # Generate the Excel report
        logger.info(f"Generating Excel report for fund: {fund_name}")
        logger.info(f"Selected stocks: {len(selected_stocks)}")
        
        try:
            report_generator = FundReportGenerator()
            output_path = report_generator.generate_report(fund_name, selected_stocks)
        except FileNotFoundError as e:
            logger.error(f"Template file not found: {str(e)}")
            # Since we're not using template anymore, this shouldn't happen
            # But keeping for backward compatibility
            return JsonResponse({
                'success': False,
                'message': f'Error generating report: {str(e)}'
            })
        except Exception as e:
            logger.error(f"Error generating report: {str(e)}")
            logger.error(traceback.format_exc())
            return JsonResponse({
                'success': False,
                'message': f'Error generating report: {str(e)}'
            })
        
        # Check if file was generated successfully
        if not os.path.exists(output_path):
            return JsonResponse({
                'success': False,
                'message': 'Failed to generate Excel report'
            })
        
        # Read the file and create HTTP response
        try:
            with open(output_path, 'rb') as excel_file:
                file_content = excel_file.read()
                
                # Create response with correct MIME type for .xlsx files
                response = HttpResponse(
                    file_content,
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
                # Set filename for download - IMPORTANT: Use .xlsx extension
                safe_fund_name = "".join(c for c in fund_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
                safe_fund_name = safe_fund_name.replace(' ', '_')
                current_date = datetime.now().strftime('%Y%m%d')
                filename = f"{safe_fund_name}_MF_Analysis_{current_date}.xlsx"  # Changed to .xlsx
                
                response['Content-Disposition'] = f'attachment; filename="{filename}"'
                response['Content-Length'] = len(file_content)
                
                # Add additional headers to ensure proper file handling
                response['Content-Transfer-Encoding'] = 'binary'
                response['Cache-Control'] = 'no-cache, no-store, must-revalidate'
                response['Pragma'] = 'no-cache'
                response['Expires'] = '0'
        
        except Exception as e:
            logger.error(f"Error reading generated file: {str(e)}")
            return JsonResponse({
                'success': False,
                'message': f'Error reading generated file: {str(e)}'
            })
        
        finally:
            # Clean up temporary file
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
                    logger.info(f"Cleaned up temporary file: {output_path}")
            except Exception as e:
                logger.warning(f"Could not delete temporary file {output_path}: {e}")
        
        return response
        
    except json.JSONDecodeError:
        return JsonResponse({
            'success': False,
            'message': 'Invalid JSON data received'
        })
    
    except Exception as e:
        logger.error(f"Unexpected error in generate_fund_report_simple: {str(e)}")
        logger.error(traceback.format_exc())
        return JsonResponse({
            'success': False,
            'message': f'An unexpected error occurred: {str(e)}'
        })