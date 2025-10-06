# This is the add_portfolio_metric_rows function to be added to enhanced_excel_export.py
# Add this at the end of the file (after line 587)

def add_portfolio_metric_rows(ws, scheme, section_start_columns, periods, total_columns):
    """
    Add 27 portfolio metric rows at bottom with values mapped to correct data sections

    Args:
        ws: Worksheet object
        scheme: AMCFundScheme instance
        section_start_columns: Dict mapping data_type to column index
        periods: Dict with period lists (ttm_periods, quarterly_periods, etc.)
        total_columns: Total number of columns in sheet
    """
    from gcia_app.models import PortfolioMetricsLog

    # Get all portfolio metrics for this fund
    portfolio_metrics_query = PortfolioMetricsLog.objects.filter(
        scheme=scheme
    ).order_by('-period_date')

    # Create period lookup
    metrics_by_period = {pm.period_date: pm for pm in portfolio_metrics_query}

    # Get all periods (TTM + Quarterly combined, sorted descending)
    all_ttm_periods = sorted(periods.get('ttm_periods', []), reverse=True)
    all_quarterly_periods = sorted(periods.get('quarterly_periods', []), reverse=True)

    # Define 27 metric rows (22 with data + 5 blank)
    # Format: (field_name, label, section_type, periods_list)
    metric_definitions = [
        ('patm', 'PATM', 'ttm_pat', all_ttm_periods),
        ('qoq_growth', 'QoQ', 'quarterly_revenue', all_quarterly_periods),
        ('yoy_growth', 'YoY', 'ttm_revenue', all_ttm_periods),
        ('revenue_6yr_cagr', '6 year CAGR', 'ttm_revenue', all_ttm_periods),
        (None, None, None, None),  # Blank row
        ('current_pe', 'Current PE', 'market_cap_free_float', all_ttm_periods),
        ('pe_2yr_avg', '2 year average', 'market_cap_free_float', all_ttm_periods),
        ('pe_5yr_avg', '5 year average', 'market_cap_free_float', all_ttm_periods),
        ('pe_2yr_reval_deval', '2 years - Reval / Deval', 'market_cap_free_float', all_ttm_periods),
        ('pe_5yr_reval_deval', '5 years - Reval / Deval', 'market_cap_free_float', all_ttm_periods),
        (None, None, None, None),  # Blank row
        ('current_pr', 'Current PR', 'market_cap_free_float', all_ttm_periods),
        ('pr_2yr_avg', '2 year average', 'market_cap_free_float', all_ttm_periods),
        ('pr_5yr_avg', '5 year average', 'market_cap_free_float', all_ttm_periods),
        ('pr_2yr_reval_deval', '2 years - Reval / Deval', 'market_cap_free_float', all_ttm_periods),
        ('pr_5yr_reval_deval', '5 years - Reval / Deval', 'market_cap_free_float', all_ttm_periods),
        (None, None, None, None),  # Blank row
        ('pr_10q_low', '10 quarter- PR- low', 'market_cap_free_float', all_ttm_periods),
        ('pr_10q_high', '10 quarter- PR- high', 'market_cap_free_float', all_ttm_periods),
        (None, None, None, None),  # Blank row
        ('alpha_bond_cagr', 'Alpha over the bond- CAGR', 'market_cap_free_float', all_ttm_periods),
        ('alpha_absolute', 'Alpha- Absolute', 'market_cap_free_float', all_ttm_periods),
        ('pe_yield', 'PE Yield', 'market_cap_free_float', all_ttm_periods),
        ('growth_rate', 'Growth', 'market_cap_free_float', all_ttm_periods),
        ('bond_rate', 'Bond Rate', 'market_cap_free_float', all_ttm_periods),
        (None, None, None, None),  # Blank row
        (None, None, None, None),  # Blank row
    ]

    # For each metric, create a row
    for metric_field, metric_label, section_type, periods_list in metric_definitions:
        metric_row = [''] * total_columns
        metric_row[0] = metric_label if metric_label else ''  # Label in Column A

        if metric_field and section_type:  # Skip blank rows
            # Get section start column
            section_start_col = section_start_columns.get(section_type)

            if section_start_col is not None and periods_list:
                # Populate values for each period in this section
                for period_idx, period in enumerate(periods_list):
                    col_idx = section_start_col + period_idx

                    if col_idx < total_columns:
                        # Get portfolio metric for this period
                        pm = metrics_by_period.get(period)
                        if pm:
                            value = getattr(pm, metric_field, None)
                            metric_row[col_idx] = value

        ws.append(metric_row)
