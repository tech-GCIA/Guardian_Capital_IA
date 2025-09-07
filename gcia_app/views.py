# views.py
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.utils.timezone import now
from datetime import timedelta, datetime
from gcia_app.forms import CustomerCreationForm
from django.contrib import messages
import openpyxl
from django.http import HttpResponse, JsonResponse
from gcia_app.forms import ExcelUploadForm, MasterDataExcelUploadForm
from gcia_app.models import AMCFundScheme, AMCFundSchemeNavLog, Stock, StockQuarterlyData, StockUploadLog
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
from decimal import Decimal, InvalidOperation
from gcia_app.index_scrapper_from_screener import get_bse500_pe_ratio
import json
from django.db.models import Q, Count
from django.core.exceptions import ValidationError
from typing import Dict, List
import logging

# Set up logging
logger = logging.getLogger(__name__)
data_quality_logger = logging.getLogger('gcia_app.data_quality')

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

def process_underlying_holdings_file(excel_file):
    data_quality_logger.info("=== STARTING HOLDINGS UPLOAD PROCESS ===")
    data_quality_logger.info(f"Processing file: {excel_file.name}")
    
    # Read the Excel file, skipping the first 2 rows as header starts at row 3
    try:
        holdings_df = pd.read_excel(excel_file, skiprows=3)
        data_quality_logger.info(f"Successfully loaded {len(holdings_df)} records from Excel file")
    except Exception as e:
        data_quality_logger.error(f"Error reading Excel file: {str(e)}")
        raise ValueError(f"Error reading Excel file: {str(e)}")

    holdings_df['PD_Date'] = pd.to_datetime(holdings_df['PD_Date'], dayfirst=True)

    # with transaction.atomic():  # Use transaction to ensure data integrity
    # Update prev Holding to Inactive state
    prev_holdings_count = SchemeUnderlyingHoldings.objects.filter(is_active=True).count()
    SchemeUnderlyingHoldings.objects.all().update(is_active=False)
    data_quality_logger.info(f"Deactivated {prev_holdings_count} existing holdings records")
    
    holdings_df = holdings_df.dropna(subset=['SD_Scheme ISIN'])
    data_quality_logger.info(f"Processing {len(holdings_df)} holdings records after cleaning")
    
    underlying_holdings_create_list = []
    funds_created = 0
    funds_activated = 0
    
    for index, row in holdings_df.iterrows():
        print(index, row["SD_Scheme AMFI Code"], pd.isna(row["SD_Scheme AMFI Code"]))
        if pd.isna(row["SD_Scheme AMFI Code"]):
            mf = AMCFundScheme.objects.filter(isin_number=row["SD_Scheme ISIN"]).first()
        else:
            mf = AMCFundScheme.objects.filter(amfi_scheme_code=row["SD_Scheme AMFI Code"], isin_number=row["SD_Scheme ISIN"]).first()
        
        if not mf:
            # Create new fund and auto-activate it since it has holdings data
            mf = AMCFundScheme.objects.create(
                name=row["Scheme Name"], 
                accord_scheme_name=row["Scheme Name"], 
                amfi_scheme_code=0 if pd.isna(row["SD_Scheme AMFI Code"]) else row["SD_Scheme AMFI Code"], 
                isin_number=row["SD_Scheme ISIN"],
                is_active=True  # Auto-activate funds that have holdings data
            )
            funds_created += 1
            print(f"Created new ACTIVE fund: {mf.name} (ID: {mf.amcfundscheme_id})")
            data_quality_logger.info(f"FUND_CREATED: {mf.name} (ID: {mf.amcfundscheme_id}, ISIN: {row['SD_Scheme ISIN']}) - AUTO-ACTIVATED")
        
        # Ensure existing fund is also active if it has holdings data
        elif not mf.is_active:
            mf.is_active = True
            mf.save()
            funds_activated += 1
            print(f"Activated existing fund: {mf.name} (ID: {mf.amcfundscheme_id})")
            data_quality_logger.info(f"FUND_ACTIVATED: {mf.name} (ID: {mf.amcfundscheme_id}, ISIN: {row['SD_Scheme ISIN']}) - Had holdings but was inactive")
        
        # Use helper function to find or create stock (consistent with Base Sheet processing)
        holding, _ = find_or_create_stock(
            name=row["PD_Instrument Name"],
            isin=row.get("PD_Company ISIN no") if not pd.isna(row.get("PD_Company ISIN no")) else None,
            bse_code=row.get("PD_BSE Code") if not pd.isna(row.get("PD_BSE Code")) else None,
            nse_code=row.get("PD_NSE Symbol") if not pd.isna(row.get("PD_NSE Symbol")) else None
        )
        
        holding_data = SchemeUnderlyingHoldings.objects.filter(amcfundscheme=mf, holding=holding, as_on_date=row["PD_Date"]).first()
        if holding_data:
            continue
        data = SchemeUnderlyingHoldings()
        data.amcfundscheme = mf
        data.holding = holding
        data.as_on_date = row["PD_Date"]
        data.as_on_month_end = row["PD_Month End"]
        data.weightage = row["PD_Holding (%)"]
        data.no_of_shares = 0 if pd.isna(row["PD_No of Shares"]) else row["PD_No of Shares"]
        underlying_holdings_create_list.append(data)
    if underlying_holdings_create_list:
        SchemeUnderlyingHoldings.objects.bulk_create(underlying_holdings_create_list)     
    
    # Data Quality Validation: Ensure all funds with holdings are active
    print("\n=== POST-UPLOAD VALIDATION ===")
    data_quality_logger.info("=== POST-UPLOAD DATA QUALITY VALIDATION ===")
    
    funds_with_holdings_ids = SchemeUnderlyingHoldings.objects.values_list('amcfundscheme_id', flat=True).distinct()
    inactive_funds_with_holdings = AMCFundScheme.objects.filter(
        amcfundscheme_id__in=funds_with_holdings_ids,
        is_active=False
    )
    
    if inactive_funds_with_holdings.exists():
        # This should not happen after our fix, but let's handle it just in case
        activated_count = inactive_funds_with_holdings.update(is_active=True)
        print(f"WARNING: Found {activated_count} inactive funds with holdings - automatically activated them")
        data_quality_logger.warning(f"POST_UPLOAD_FIX: Found {activated_count} inactive funds with holdings - automatically activated them")
        for fund in inactive_funds_with_holdings[:10]:  # Log first 10 for detail
            data_quality_logger.warning(f"POST_UPLOAD_ACTIVATION: {fund.name} (ID: {fund.amcfundscheme_id})")
    
    # Summary statistics
    total_holdings_created = len(underlying_holdings_create_list)
    unique_funds_with_holdings = len(set(funds_with_holdings_ids))
    active_funds_with_holdings = AMCFundScheme.objects.filter(
        amcfundscheme_id__in=funds_with_holdings_ids,
        is_active=True
    ).count()
    
    print(f"✓ Created {total_holdings_created} new holdings records")
    print(f"✓ Processed {unique_funds_with_holdings} unique funds with holdings")
    print(f"✓ All {active_funds_with_holdings} funds with holdings are ACTIVE")
    print("=== UPLOAD COMPLETED SUCCESSFULLY ===\n")
    
    # Log comprehensive summary
    data_quality_logger.info("=== UPLOAD SUMMARY STATISTICS ===")
    data_quality_logger.info(f"Holdings created: {total_holdings_created}")
    data_quality_logger.info(f"New funds created: {funds_created}")
    data_quality_logger.info(f"Existing funds activated: {funds_activated}")
    data_quality_logger.info(f"Total unique funds with holdings: {unique_funds_with_holdings}")
    data_quality_logger.info(f"All funds with holdings are active: {active_funds_with_holdings == unique_funds_with_holdings}")
    data_quality_logger.info("=== HOLDINGS UPLOAD PROCESS COMPLETED SUCCESSFULLY ===")
    
    return f"Success: {total_holdings_created} holdings created for {unique_funds_with_holdings} active funds"       


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

def process_stocks_base_sheet(excel_file, user):
    """
    Process Base Sheet Excel file to update Stock and StockQuarterlyData models
    
    Args:
        excel_file: The uploaded Excel file containing Base Sheet data
        user: The user uploading the file
        
    Returns:
        tuple: (stats, upload_log) containing processing statistics and upload log
    """
    from datetime import datetime as dt
    
    # Initialize statistics
    stats = {
        'stocks_added': 0,
        'stocks_updated': 0,
        'quarterly_records_added': 0,
        'quarterly_records_updated': 0,
        'errors': [],
        'total_rows_processed': 0
    }
    
    # Create upload log entry
    upload_log = StockUploadLog.objects.create(
        uploaded_by=user,
        filename=excel_file.name,
        file_size=excel_file.size,
        status='processing',
        processing_started_at=dt.now()
    )
    
    try:
        # Read Excel file with proper header handling
        df_raw = pd.read_excel(excel_file, sheet_name='App-Base Sheet', header=None)
        
        # Set header row (row 8 in 1-indexed, which is index 7 in 0-indexed)
        df_data = pd.read_excel(excel_file, sheet_name='App-Base Sheet', header=7)
        
        if df_data.empty:
            raise ValueError("No data found in the Excel file")
        
        # Remove sample/test rows (rows with 'XX' or 'Sample' in S. No. column)
        df_data = df_data[~df_data['S. No.'].astype(str).str.contains('XX|Sample', na=False, case=False)]
        
        stats['total_rows_processed'] = len(df_data)
        
        with transaction.atomic():
            for index, row in df_data.iterrows():
                try:
                    # Process Stock data (static company information)
                    stock_data = {
                        'name': str(row['Company Name']).strip() if pd.notna(row['Company Name']) else '',
                        'accord_code': str(row['Accord Code']).strip() if pd.notna(row['Accord Code']) else None,
                        'sector': str(row['Sector']).strip() if pd.notna(row['Sector']) else None,
                        'cap': str(row['Cap']).strip() if pd.notna(row['Cap']) else None,
                        'free_float': _safe_decimal(row['Free Float']),
                        'revenue_6yr_cagr': _safe_decimal(row['6 Year CAGR']),  # Column G
                        'revenue_ttm': _safe_decimal(row['TTM']),  # Column H
                        'pat_6yr_cagr': _safe_decimal(row['6 Year CAGR.1']),  # Column I
                        'pat_ttm': _safe_decimal(row['TTM.1']),  # Column J
                        'current_pr': _safe_decimal(row['Current']),  # Column K
                        'pr_2yr_avg': _safe_decimal(row['2 Yr Avg']),  # Column L
                        'reval_deval': _safe_decimal(row['Reval/deval']),  # Column M
                    }
                    
                    # Extract codes from end columns
                    if 'BSE Code' in df_data.columns:
                        stock_data['bse_code'] = str(row['BSE Code']).strip() if pd.notna(row['BSE Code']) else None
                    if 'NSE Code' in df_data.columns:
                        stock_data['nse_code'] = str(row['NSE Code']).strip() if pd.notna(row['NSE Code']) else None
                    if 'ISIN' in df_data.columns:
                        stock_data['isin'] = str(row['ISIN']).strip() if pd.notna(row['ISIN']) else None
                    
                    # Use helper function to find or create stock
                    stock, created = find_or_create_stock(
                        name=stock_data['name'],
                        isin=stock_data.get('isin'),
                        bse_code=stock_data.get('bse_code'),
                        nse_code=stock_data.get('nse_code'),
                        symbol=stock_data.get('symbol'),
                        # Pass all additional data fields
                        accord_code=stock_data.get('accord_code'),
                        sector=stock_data.get('sector'),
                        cap=stock_data.get('cap'),
                        free_float=stock_data.get('free_float'),
                        revenue_6yr_cagr=stock_data.get('revenue_6yr_cagr'),
                        revenue_ttm=stock_data.get('revenue_ttm'),
                        pat_6yr_cagr=stock_data.get('pat_6yr_cagr'),
                        pat_ttm=stock_data.get('pat_ttm'),
                        current_pr=stock_data.get('current_pr'),
                        pr_2yr_avg=stock_data.get('pr_2yr_avg'),
                        reval_deval=stock_data.get('reval_deval')
                    )
                    
                    if created:
                        stats['stocks_added'] += 1
                    else:
                        stats['stocks_updated'] += 1
                    
                    # Process quarterly data (all date columns)
                    date_columns = [col for col in df_data.columns if isinstance(col, dt)]
                    
                    for date_col in date_columns:
                        if pd.notna(row[date_col]):
                            # Determine which metric this date column represents
                            col_index = df_data.columns.get_loc(date_col)
                            metric_type = _determine_metric_type(col_index)
                            
                            # Create quarterly data entry
                            quarterly_data = {
                                'stock': stock,
                                'quarter_date': date_col.date() if hasattr(date_col, 'date') else date_col,
                                metric_type: _safe_decimal(row[date_col])
                            }
                            
                            quarterly_obj, q_created = StockQuarterlyData.objects.update_or_create(
                                stock=stock,
                                quarter_date=quarterly_data['quarter_date'],
                                defaults=quarterly_data
                            )
                            
                            if q_created:
                                stats['quarterly_records_added'] += 1
                            else:
                                stats['quarterly_records_updated'] += 1
                    
                    # Handle period format columns (YYYYMM format)
                    period_columns = [col for col in df_data.columns if 
                                    isinstance(col, str) and col.replace('.', '').isdigit() and len(col.replace('.', '')) == 6]
                    
                    for period_col in period_columns:
                        if pd.notna(row[period_col]):
                            # Convert period format to date (YYYYMM -> YYYY-MM-last_day)
                            period_str = str(period_col).split('.')[0]  # Remove .1, .2 etc.
                            if len(period_str) == 6:
                                year = int(period_str[:4])
                                month = int(period_str[4:6])
                                
                                # Create last day of month
                                import calendar
                                last_day = calendar.monthrange(year, month)[1]
                                quarter_date = dt(year, month, last_day).date()
                                
                                col_index = df_data.columns.get_loc(period_col)
                                metric_type = _determine_metric_type(col_index)
                                
                                quarterly_data = {
                                    'stock': stock,
                                    'quarter_date': quarter_date,
                                    metric_type: _safe_decimal(row[period_col])
                                }
                                
                                quarterly_obj, q_created = StockQuarterlyData.objects.update_or_create(
                                    stock=stock,
                                    quarter_date=quarter_date,
                                    defaults=quarterly_data
                                )
                                
                                if q_created:
                                    stats['quarterly_records_added'] += 1
                                else:
                                    stats['quarterly_records_updated'] += 1
                    
                except Exception as e:
                    error_msg = f"Error processing row {index + 1}: {str(e)}"
                    stats['errors'].append(error_msg)
                    continue
        
        # Update upload log with success
        upload_log.status = 'completed'
        upload_log.stocks_added = stats['stocks_added']
        upload_log.stocks_updated = stats['stocks_updated']
        upload_log.quarterly_records_added = stats['quarterly_records_added']
        upload_log.quarterly_records_updated = stats['quarterly_records_updated']
        upload_log.processing_completed_at = dt.now()
        upload_log.processing_time = upload_log.processing_completed_at - upload_log.processing_started_at
        upload_log.save()
        
    except Exception as e:
        # Update upload log with failure
        upload_log.status = 'failed'
        upload_log.error_message = str(e)
        upload_log.processing_completed_at = dt.now()
        upload_log.processing_time = upload_log.processing_completed_at - upload_log.processing_started_at
        upload_log.save()
        raise e
    
    return stats, upload_log

def find_or_create_stock(name, isin=None, bse_code=None, nse_code=None, symbol=None, **additional_data):
    """
    Find existing stock or create new one using multiple matching criteria
    Priority order: ISIN -> BSE Code -> NSE Code -> Name
    """
    stock = None
    created = False
    
    # Clean up input data
    name = str(name).strip() if name else ''
    isin = str(isin).strip() if isin and not pd.isna(isin) else None
    bse_code = str(bse_code).strip() if bse_code and not pd.isna(bse_code) else None
    nse_code = str(nse_code).strip() if nse_code and not pd.isna(nse_code) else None
    symbol = str(symbol).strip() if symbol and not pd.isna(symbol) else None
    
    # Try to find existing stock using multiple criteria in priority order
    if isin:
        stock = Stock.objects.filter(isin=isin).first()
    
    if not stock and bse_code:
        stock = Stock.objects.filter(bse_code=bse_code).first()
    
    if not stock and nse_code:
        stock = Stock.objects.filter(nse_code=nse_code).first()
    
    if not stock and name:
        stock = Stock.objects.filter(name__iexact=name).first()
    
    # Prepare stock data
    stock_data = {
        'name': name,
        'isin': isin,
        'bse_code': bse_code,
        'nse_code': nse_code,
        'symbol': symbol or nse_code or bse_code,  # Use NSE code, then BSE code as fallback
        **additional_data
    }
    
    # Remove None/empty values
    stock_data = {k: v for k, v in stock_data.items() if v is not None and v != ''}
    
    if stock:
        # Update existing stock with new data
        for key, value in stock_data.items():
            if value is not None and value != '':
                setattr(stock, key, value)
        stock.save()
        created = False
    else:
        # Create new stock
        stock = Stock.objects.create(**stock_data)
        created = True
    
    return stock, created

def _safe_decimal(value):
    """Safely convert value to Decimal, return None if invalid"""
    if pd.isna(value) or value == '' or value is None:
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return None

def _determine_metric_type(col_index):
    """Determine which metric type based on column index"""
    # Based on the Excel structure analysis
    if 14 <= col_index <= 41:  # Market Cap columns
        return 'mcap'
    elif 43 <= col_index <= 71:  # Free Float Market Cap columns  
        return 'free_float_mcap'
    elif 72 <= col_index <= 107:  # TTM Revenue columns
        return 'ttm_revenue'
    elif 108 <= col_index <= 143:  # TTM Revenue Free Float columns
        return 'ttm_revenue_free_float'
    elif 144 <= col_index <= 179:  # TTM PAT columns
        return 'pat'
    elif 180 <= col_index <= 215:  # TTM PAT Free Float columns
        return 'ttm_pat_free_float'
    elif 216 <= col_index <= 251:  # Quarterly Revenue columns
        return 'quarterly_revenue'
    elif 252 <= col_index <= 287:  # Quarterly Revenue Free Float columns
        return 'quarterly_revenue_free_float'
    elif 288 <= col_index <= 323:  # Quarterly PAT columns
        return 'quarterly_pat'
    elif 324 <= col_index <= 359:  # Quarterly PAT Free Float columns
        return 'quarterly_pat_free_float'
    elif 360 <= col_index <= 371:  # ROCE columns
        return 'roce'
    elif 372 <= col_index <= 383:  # ROE columns
        return 'roe'
    elif 384 <= col_index <= 395:  # Retention columns
        return 'retention'
    elif col_index == 396:  # Share Price column
        return 'share_price'
    elif 397 <= col_index <= 408:  # PR columns
        return 'pr_quarterly'
    elif 409 <= col_index <= 420:  # PE columns
        return 'pe_quarterly'
    else:
        return 'mcap'  # Default fallback

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
                elif file_type == "underlying_holdings":
                    # Process Underlying Holdings File
                    stats = process_underlying_holdings_file(excel_file)
                elif file_type == "stocks_base":
                    # Process Stocks Base Sheet
                    stats, upload_log = process_stocks_base_sheet(excel_file, request.user)
                    
                    # Create detailed success message
                    success_message = (
                        f"Base Sheet processed successfully! "
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
                else:
                    messages.success(request, "File uploaded and data processed successfully!")

            except Exception as e:
                print(str(e))
                print(traceback.format_exc())
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

# COMMENTED OUT - Upload Stock Data functionality disabled - duplicate imports removed
# The required imports are already included at the top of the file

# COMMENTED OUT - Upload Stock Data functionality disabled
# @login_required
# def upload_stock_data(request):
#     """
#     View to handle stock data Excel file upload and processing
#     """
#     form = StockDataUploadForm()
#     recent_uploads = []
#     
#     # Get recent uploads for the current user
#     if request.user.is_authenticated:
#         recent_uploads = StockUploadLog.objects.filter(
#             uploaded_by=request.user
#         ).order_by('-uploaded_at')[:10]
#     
#     if request.method == 'POST' and request.FILES.get('excel_file'):
#         form = StockDataUploadForm(request.POST, request.FILES)
#         
#         if form.is_valid():
#             excel_file = request.FILES['excel_file']
#             
#             try:
#                 # Create uploads directory if it doesn't exist
#                 uploads_dir = os.path.join(settings.MEDIA_ROOT, 'uploads')
#                 os.makedirs(uploads_dir, exist_ok=True)
#                 
#                 # Save file temporarily
#                 file_path = default_storage.save(
#                     f'uploads/{excel_file.name}',
#                     ContentFile(excel_file.read())
#                 )
#                 
#                 try:
#                     # Process the file
#                     with default_storage.open(file_path, 'rb') as stored_file:
#                         # Create a temporary file for processing
#                         with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
#                             temp_file.write(stored_file.read())
#                             temp_file.flush()
#                             
#                             # Reset file pointer
#                             excel_file.seek(0)
#                             
#                             # Process the data
#                             stats, upload_log = process_stock_data_file(excel_file, request.user)
#                             
#                             # Show success message with statistics
#                             success_message = (
#                                 f"File processed successfully! "
#                                 f"Stocks added: {stats['stocks_added']}, "
#                                 f"Stocks updated: {stats['stocks_updated']}, "
#                                 f"Quarterly records added: {stats['quarterly_records_added']}, "
#                                 f"Quarterly records updated: {stats['quarterly_records_updated']}"
#                             )
#                             
#                             if stats['errors']:
#                                 success_message += f" (with {len(stats['errors'])} warnings)"
#                             
#                             messages.success(request, success_message)
#                             
#                             # Log any warnings
#                             if stats['errors']:
#                                 for error in stats['errors'][:5]:  # Show first 5 errors
#                                     messages.warning(request, f"Warning: {error}")
#                                 
#                                 if len(stats['errors']) > 5:
#                                     messages.info(request, f"... and {len(stats['errors']) - 5} more warnings")
#                 
#                 except Exception as e:
#                     logger.error(f"Error processing stock data file: {str(e)}")
#                     logger.error(traceback.format_exc())
#                     messages.error(request, f"Error processing file: {str(e)}")
#                 
#                 finally:
#                     # Clean up temporary files
#                     try:
#                         default_storage.delete(file_path)
#                         if 'temp_file' in locals():
#                             os.unlink(temp_file.name)
#                     except Exception as cleanup_error:
#                         logger.warning(f"Error cleaning up files: {cleanup_error}")
#                 
#             except Exception as e:
#                 logger.error(f"Error handling file upload: {str(e)}")
#                 logger.error(traceback.format_exc())
#                 messages.error(request, f"Error uploading file: {str(e)}")
#             
#             # Redirect to prevent form resubmission
#             return redirect('upload_stock_data')
#     
#     context = {
#         'form': form,
#         'recent_uploads': recent_uploads,
#     }
#     
#     return render(request, 'gcia_app/upload_stock_data.html', context)

# Add this view to your existing gcia_app/views.py

# COMMENTED OUT - Fund Analysis functionality disabled
"""
# Replace your search_stocks view in gcia_app/views.py with this improved version:

@login_required
def search_stocks(request):
    \"\"\"
    AJAX endpoint for stock search functionality with debugging
    \"\"\"
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
"""

# COMMENTED OUT - Fund Analysis functionality disabled
"""
# Also update your fund_analysis view to include debug info:

@login_required
def fund_analysis(request):
    \"\"\"
    Fund Analysis page with simple stock selection and table management
    \"\"\"
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
"""

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

# COMMENTED OUT - Fund Analysis functionality disabled
"""
# Replace the existing generate_fund_report_simple function with this:

@login_required
@require_http_methods(["POST"])
def generate_fund_report_simple(request):
    # Generate and download Excel report directly when "Generate Report" is clicked
    # Fixed version that properly handles .xlsx format
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


# MF Metrics Views
"""
# End of generate_fund_report_simple function comment block

import threading
import time
import tempfile
from gcia_app.mf_metrics_calculator import MFMetricsCalculator
from gcia_app.models import MutualFundMetrics, MetricsCalculationLog
from django.views.decorators.http import require_http_methods

# Simple progress tracking - always starts fresh
calculation_progress = {
    'status': 'idle',
    'total_funds': 0,
    'processed_funds': 0,
    'successful_funds': 0,
    'partial_funds': 0,
    'failed_funds': 0,
    'current_fund': '',
    'error_message': '',
    'log_id': None
}

@login_required
def mf_metrics_page(request):
    """
    Main MF Metrics management page - always starts with fresh state
    """
    # Force reset calculation progress for fresh start every time
    global calculation_progress
    calculation_progress.update({
        'status': 'idle',
        'total_funds': 0,
        'processed_funds': 0,
        'successful_funds': 0,
        'partial_funds': 0,
        'failed_funds': 0,
        'current_fund': '',
        'error_message': '',
        'log_id': None
    })
    
    # Get last calculation log
    last_calculation = MetricsCalculationLog.objects.filter(
        initiated_by=request.user
    ).order_by('-started_at').first()
    
    # Get active funds for dropdown
    active_funds = AMCFundScheme.objects.filter(
        is_active=True
    ).annotate(
        holdings_count=Count('schemeunderlyingholdings')
    ).order_by('name')
    
    # Get recent metrics summary (top 20 by market cap)
    metrics_summary = MutualFundMetrics.objects.filter(
        calculation_status__in=['success', 'partial']
    ).select_related(
        'amcfundscheme'
    ).order_by(
        '-portfolio_market_cap'
    )[:20]
    
    context = {
        'last_calculation': last_calculation,
        'active_funds': active_funds,
        'metrics_summary': metrics_summary,
        'total_active_funds': active_funds.count(),
    }
    
    return render(request, 'gcia_app/mf_metrics.html', context)


@login_required
@require_http_methods(["POST"])
def update_all_mf_metrics(request):
    """
    Start bulk calculation of all fund metrics (async) - simplified tracking
    """
    global calculation_progress
    
    try:
        # Check if calculation is already running
        if calculation_progress['status'] == 'running':
            return JsonResponse({
                'success': False,
                'message': 'Calculation is already in progress'
            })
        
        # Parse request data
        data = json.loads(request.body)
        force_recalculate = data.get('force_recalculate', False)
        
        # Initialize progress tracking
        active_funds_count = AMCFundScheme.objects.filter(is_active=True).count()
        calculation_progress.update({
            'status': 'running',
            'total_funds': active_funds_count,
            'processed_funds': 0,
            'successful_funds': 0,
            'partial_funds': 0,
            'failed_funds': 0,
            'current_fund': 'Initializing...',
            'error_message': '',
            'log_id': None
        })
        
        # Start calculation in separate thread
        calculation_thread = threading.Thread(
            target=run_metrics_calculation,
            args=(request.user, force_recalculate)
        )
        calculation_thread.daemon = True
        calculation_thread.start()
        
        return JsonResponse({
            'success': True,
            'message': f'Metrics calculation started for {active_funds_count} funds',
            'total_funds': active_funds_count
        })
        
    except json.JSONDecodeError:
        return JsonResponse({
            'success': False,
            'message': 'Invalid JSON data received'
        })
    except Exception as e:
        logger.error(f"Error starting metrics calculation: {str(e)}")
        return JsonResponse({
            'success': False,
            'message': f'Error starting calculation: {str(e)}'
        })


def run_metrics_calculation(user, force_recalculate=False):
    """
    Run the metrics calculation in background thread
    """
    global calculation_progress
    
    try:
        calculator = MFMetricsCalculator()
        
        # Get all active funds
        active_funds = AMCFundScheme.objects.filter(is_active=True).order_by('name')
        
        calculation_progress['total_funds'] = active_funds.count()
        
        # Create calculation log
        calc_log = MetricsCalculationLog.objects.create(
            initiated_by=user,
            calculation_type='bulk_all',
            status='running',
            total_funds_targeted=active_funds.count()
        )
        calculation_progress['log_id'] = calc_log.log_id
        
        # Process each fund
        for index, fund in enumerate(active_funds, 1):
            try:
                calculation_progress['current_fund'] = fund.name
                calculation_progress['processed_funds'] = index
                
                result = calculator.calculate_single_fund_metrics(fund, force_recalculate)
                
                if result['status'] == 'success':
                    calculation_progress['successful_funds'] += 1
                elif result['status'] == 'partial':
                    calculation_progress['partial_funds'] += 1
                elif result['status'] == 'failed':
                    calculation_progress['failed_funds'] += 1
                # 'skipped' doesn't count as failed
                
                # Brief pause to prevent overwhelming the database
                time.sleep(0.1)
                
            except Exception as e:
                logger.error(f"Error processing fund {fund.name}: {str(e)}")
                calculation_progress['failed_funds'] += 1
        
        # Update final log
        calc_log.status = 'completed'
        calc_log.completed_at = timezone.now()
        calc_log.funds_processed_successfully = calculation_progress['successful_funds']
        calc_log.funds_with_partial_data = calculation_progress['partial_funds']
        calc_log.funds_failed = calculation_progress['failed_funds']
        calc_log.save()
        
        calculation_progress['status'] = 'completed'
        logger.info(f"Metrics calculation completed: {calculation_progress}")
        
    except Exception as e:
        logger.error(f"Error in metrics calculation thread: {str(e)}")
        calculation_progress['status'] = 'failed'
        calculation_progress['error_message'] = str(e)
        
        # Update log if it exists
        if calculation_progress['log_id']:
            try:
                calc_log = MetricsCalculationLog.objects.get(log_id=calculation_progress['log_id'])
                calc_log.status = 'failed'
                calc_log.completed_at = timezone.now()
                calc_log.error_summary = str(e)
                calc_log.save()
            except:
                pass


@login_required
@require_http_methods(["GET"])
def mf_metrics_progress(request):
    """
    Get current progress of metrics calculation
    """
    global calculation_progress
    
    return JsonResponse({
        'status': calculation_progress['status'],
        'total_funds': calculation_progress['total_funds'],
        'processed_funds': calculation_progress['processed_funds'],
        'successful_funds': calculation_progress['successful_funds'],
        'partial_funds': calculation_progress['partial_funds'],
        'failed_funds': calculation_progress['failed_funds'],
        'current_fund': calculation_progress['current_fund'],
        'error_message': calculation_progress['error_message'],
        'log_id': calculation_progress['log_id']
    })




@login_required
@require_http_methods(["POST"])
def download_portfolio_analysis(request):
    """
    Generate and download Portfolio Analysis Excel for a specific fund
    """
    try:
        # Parse request data
        data = json.loads(request.body)
        fund_id = data.get('fund_id')
        
        if not fund_id:
            return JsonResponse({
                'success': False,
                'message': 'Fund ID is required'
            })
        
        # Get the fund
        try:
            fund = AMCFundScheme.objects.get(amcfundscheme_id=fund_id, is_active=True)
        except AMCFundScheme.DoesNotExist:
            return JsonResponse({
                'success': False,
                'message': 'Fund not found or inactive'
            })
        
        # Get fund metrics (calculate if needed)
        calculator = MFMetricsCalculator()
        metrics_result = calculator.calculate_single_fund_metrics(fund, force_recalculate=False)
        
        if metrics_result['status'] == 'failed':
            return JsonResponse({
                'success': False,
                'message': f'Cannot generate analysis: {metrics_result.get("message", "No data available")}'
            })
        
        # Generate Excel file
        excel_generator = PortfolioAnalysisExcelGenerator()
        excel_file_path = excel_generator.generate_portfolio_analysis_excel(fund, metrics_result)
        
        if not excel_file_path or not os.path.exists(excel_file_path):
            return JsonResponse({
                'success': False,
                'message': 'Failed to generate Excel file'
            })
        
        # Prepare file response
        try:
            with open(excel_file_path, 'rb') as excel_file:
                response = HttpResponse(
                    excel_file.read(),
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
                # Generate filename
                safe_fund_name = re.sub(r'[^\w\s-]', '', fund.name).strip()
                safe_fund_name = re.sub(r'[-\s]+', '_', safe_fund_name)
                current_date = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"{safe_fund_name}_Portfolio_Analysis_{current_date}.xlsx"
                
                response['Content-Disposition'] = f'attachment; filename="{filename}"'
                
                return response
                
        except Exception as e:
            logger.error(f"Error reading generated Excel file: {str(e)}")
            return JsonResponse({
                'success': False,
                'message': f'Error reading generated file: {str(e)}'
            })
        
        finally:
            # Clean up temporary file
            try:
                if os.path.exists(excel_file_path):
                    os.remove(excel_file_path)
            except Exception as e:
                logger.warning(f"Could not delete temporary file {excel_file_path}: {e}")
                
    except json.JSONDecodeError:
        return JsonResponse({
            'success': False,
            'message': 'Invalid JSON data received'
        })
    except Exception as e:
        logger.error(f"Error in download_portfolio_analysis: {str(e)}")
        return JsonResponse({
            'success': False,
            'message': f'An unexpected error occurred: {str(e)}'
        })


# Portfolio Analysis Excel Generator
class PortfolioAnalysisExcelGenerator:
    """
    Generate Portfolio Analysis Excel file matching the Old Bridge format
    """
    
    def __init__(self):
        self.workbook = None
        self.worksheet = None
    
    def generate_portfolio_analysis_excel(self, fund: AMCFundScheme, metrics_result: Dict) -> str:
        """
        Generate Portfolio Analysis Excel file for the given fund
        
        Returns:
            str: Path to generated Excel file
        """
        try:
            # Create temporary file
            temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
            temp_file.close()
            
            # Get holdings data
            calculator = MFMetricsCalculator()
            holdings_data = calculator.get_portfolio_holdings_data(fund)
            
            if not holdings_data:
                raise ValueError("No holdings data available for this fund")
            
            # Create Excel file with pandas and openpyxl
            with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
                # Create Portfolio Analysis sheet
                self.create_portfolio_analysis_sheet(writer, fund, holdings_data, metrics_result)
            
            logger.info(f"Generated Portfolio Analysis Excel for {fund.name} at {temp_file.name}")
            return temp_file.name
            
        except Exception as e:
            logger.error(f"Error generating Portfolio Analysis Excel: {str(e)}")
            raise
    
    def create_portfolio_analysis_sheet(self, writer, fund: AMCFundScheme, holdings_data: List[Dict], metrics_result: Dict):
        """
        Create the Portfolio Analysis sheet matching Old Bridge format
        """
        # Prepare data for Excel
        portfolio_data = []
        
        # Add header rows (matching Old Bridge format)
        header_data = {
            'Company Name': fund.name,
            'Accord Code': '',
            'Sector': 'Portfolio Analysis',
            'Cap': f'As of {datetime.now().strftime("%B %Y")}',
            'Weightage (%)': '',
            'Market Cap (Cr)': '',
            'TTM PAT (Cr)': '',
            'Current PE': '',
            'Quarterly Data': ''
        }
        portfolio_data.append(header_data)
        
        # Add empty rows for formatting
        for i in range(6):
            portfolio_data.append({col: '' for col in header_data.keys()})
        
        # Add holdings data (rows 8-34 equivalent)
        for holding in holdings_data:
            if holding['quarterly_data']:
                latest_quarter = holding['quarterly_data'][0]
                
                holding_row = {
                    'Company Name': holding['stock_name'],
                    'Accord Code': '',  # Could be mapped from stock data if available
                    'Sector': holding['sector'] or '',
                    'Cap': '',  # Market cap category
                    'Weightage (%)': holding['weightage'],
                    'Market Cap (Cr)': latest_quarter.get('mcap', ''),
                    'TTM PAT (Cr)': latest_quarter.get('ttm_pat', ''),
                    'Current PE': latest_quarter.get('pe_ratio', ''),
                    'Quarterly Data': f"Q{latest_quarter.get('quarter_number', '')}-{latest_quarter.get('quarter_year', '')}"
                }
                
                # Add quarterly data columns (dates as column headers)
                for qdata in holding['quarterly_data'][:8]:  # Last 8 quarters
                    quarter_key = f"Q{qdata.get('quarter_number')}-{qdata.get('quarter_year')}"
                    if quarter_key not in holding_row:
                        holding_row[quarter_key] = qdata.get('mcap', '')
                
                portfolio_data.append(holding_row)
        
        # Add metrics calculation rows (equivalent to rows 35-60)
        metrics_rows = self.create_metrics_rows(fund, metrics_result)
        portfolio_data.extend(metrics_rows)
        
        # Convert to DataFrame and write to Excel
        df = pd.DataFrame(portfolio_data)
        df.to_excel(writer, sheet_name='Portfolio Analysis', index=False)
        
        # Get worksheet for formatting
        worksheet = writer.sheets['Portfolio Analysis']
        
        # Apply formatting
        self.apply_excel_formatting(worksheet, len(holdings_data))
    
    def create_metrics_rows(self, fund: AMCFundScheme, metrics_result: Dict) -> List[Dict]:
        """
        Create calculated metrics rows for the Excel
        """
        metrics_rows = []
        
        if 'calculated_metrics' in metrics_result:
            calculated = metrics_result['calculated_metrics']
            
            # Add empty separator row
            metrics_rows.append({col: '' for col in ['Company Name', 'Accord Code', 'Sector', 'Cap', 'Weightage (%)', 'Market Cap (Cr)', 'TTM PAT (Cr)', 'Current PE', 'Quarterly Data']})
            
            # TOTALS row
            metrics_rows.append({
                'Company Name': 'TOTALS',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': '',
                'Market Cap (Cr)': calculated.get('portfolio_market_cap', ''),
                'TTM PAT (Cr)': calculated.get('portfolio_ttm_pat', ''),
                'Current PE': calculated.get('portfolio_current_pe', ''),
                'Quarterly Data': ''
            })
            
            # PATM row
            metrics_rows.append({
                'Company Name': 'PATM',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': '',
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': calculated.get('portfolio_pat', ''),
                'Current PE': '',
                'Quarterly Data': ''
            })
            
            # QoQ Growth row
            metrics_rows.append({
                'Company Name': 'QoQ Growth (%)',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': calculated.get('portfolio_qoq_growth', ''),
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': '',
                'Quarterly Data': ''
            })
            
            # YoY Growth row
            metrics_rows.append({
                'Company Name': 'YoY Growth (%)',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': calculated.get('portfolio_yoy_growth', ''),
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': '',
                'Quarterly Data': ''
            })
            
            # 6 Year CAGR rows
            metrics_rows.append({
                'Company Name': '6 Year Revenue CAGR (%)',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': calculated.get('portfolio_6yr_revenue_cagr', ''),
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': '',
                'Quarterly Data': ''
            })
            
            metrics_rows.append({
                'Company Name': '6 Year PAT CAGR (%)',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': calculated.get('portfolio_6yr_pat_cagr', ''),
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': '',
                'Quarterly Data': ''
            })
            
            # PE Average rows
            metrics_rows.append({
                'Company Name': '2 Yr Avg PE',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': '',
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': calculated.get('portfolio_2yr_avg_pe', ''),
                'Quarterly Data': ''
            })
            
            metrics_rows.append({
                'Company Name': '5 Yr Avg PE',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': '',
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': calculated.get('portfolio_5yr_avg_pe', ''),
                'Quarterly Data': ''
            })
            
            # Reval/Deval row
            metrics_rows.append({
                'Company Name': 'Reval/Deval (%)',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': calculated.get('portfolio_reval_deval', ''),
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': '',
                'Quarterly Data': ''
            })
            
            # P/R ratios
            metrics_rows.append({
                'Company Name': 'Current P/R',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': '',
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': calculated.get('portfolio_current_pr', ''),
                'Quarterly Data': ''
            })
            
            metrics_rows.append({
                'Company Name': '2 Yr Avg P/R',
                'Accord Code': '',
                'Sector': '',
                'Cap': '',
                'Weightage (%)': '',
                'Market Cap (Cr)': '',
                'TTM PAT (Cr)': '',
                'Current PE': calculated.get('portfolio_2yr_avg_pr', ''),
                'Quarterly Data': ''
            })
            
            # Performance metrics
            if calculated.get('portfolio_alpha'):
                metrics_rows.append({
                    'Company Name': 'Alpha (%)',
                    'Accord Code': '',
                    'Sector': '',
                    'Cap': '',
                    'Weightage (%)': calculated.get('portfolio_alpha', ''),
                    'Market Cap (Cr)': '',
                    'TTM PAT (Cr)': '',
                    'Current PE': '',
                    'Quarterly Data': ''
                })
        
        return metrics_rows
    
    def apply_excel_formatting(self, worksheet, holdings_count: int):
        """
        Apply formatting to make the Excel look professional
        """
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        
        # Define styles
        header_font = Font(bold=True, size=12)
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        metrics_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Format header rows
        for row in range(1, 8):
            for col in range(1, 10):
                cell = worksheet.cell(row=row, column=col)
                if row == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                cell.border = thin_border
        
        # Format holdings data
        for row in range(8, 8 + holdings_count):
            for col in range(1, 10):
                cell = worksheet.cell(row=row, column=col)
                cell.border = thin_border
                if col == 5:  # Weightage column
                    cell.number_format = '0.00%'
                elif col in [6, 7]:  # Market cap and PAT columns
                    cell.number_format = '#,##0.00'
        
        # Format metrics rows
        metrics_start_row = 8 + holdings_count + 1
        for row in range(metrics_start_row, worksheet.max_row + 1):
            for col in range(1, 10):
                cell = worksheet.cell(row=row, column=col)
                cell.fill = metrics_fill
                cell.border = thin_border
                if col == 1:  # Metric names
                    cell.font = Font(bold=True)
        
        # Adjust column widths
        column_widths = [25, 12, 15, 10, 12, 15, 15, 12, 12]
        for i, width in enumerate(column_widths, 1):
            worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = width