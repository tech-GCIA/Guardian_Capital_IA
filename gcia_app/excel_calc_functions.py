"""
Portfolio Analysis Calculation Functions for Excel Export
Calculates metrics from TOTALS row values in Excel sheet
"""

import logging
logger = logging.getLogger(__name__)


def calculate_patm_from_totals(totals_data, section_cols):
    """
    PATM = (Total PAT / Total Revenue) × 100
    Calculate for: TTM PAT, TTM PAT FF, Q PAT (each quarter), Q PAT FF (each quarter)
    """
    results = {}

    # TTM PAT
    if 'ttm_pat' in section_cols and 'ttm_revenue' in section_cols:
        ttm_pat_cols = section_cols['ttm_pat']
        ttm_rev_cols = section_cols['ttm_revenue']
        if len(ttm_pat_cols) > 0 and len(ttm_rev_cols) > 0:
            pat_col = ttm_pat_cols[0]
            rev_col = ttm_rev_cols[0]
            if totals_data[rev_col] and totals_data[rev_col] != 0:
                results['ttm_pat'] = (totals_data[pat_col] / totals_data[rev_col]) * 100

    # TTM PAT Free Float
    if 'ttm_pat_free_float' in section_cols and 'ttm_revenue_free_float' in section_cols:
        ttm_pat_ff_cols = section_cols['ttm_pat_free_float']
        ttm_rev_ff_cols = section_cols['ttm_revenue_free_float']
        if len(ttm_pat_ff_cols) > 0 and len(ttm_rev_ff_cols) > 0:
            pat_col = ttm_pat_ff_cols[0]
            rev_col = ttm_rev_ff_cols[0]
            if totals_data[rev_col] and totals_data[rev_col] != 0:
                results['ttm_pat_ff'] = (totals_data[pat_col] / totals_data[rev_col]) * 100

    # Quarterly PAT (for each quarter)
    if 'quarterly_pat' in section_cols and 'quarterly_revenue' in section_cols:
        q_pat_cols = section_cols['quarterly_pat']
        q_rev_cols = section_cols['quarterly_revenue']
        for idx in range(min(len(q_pat_cols), len(q_rev_cols))):
            if totals_data[q_rev_cols[idx]] and totals_data[q_rev_cols[idx]] != 0:
                results[f'q_pat_{idx}'] = (totals_data[q_pat_cols[idx]] / totals_data[q_rev_cols[idx]]) * 100

    # Quarterly PAT Free Float
    if 'quarterly_pat_free_float' in section_cols and 'quarterly_revenue_free_float' in section_cols:
        q_pat_ff_cols = section_cols['quarterly_pat_free_float']
        q_rev_ff_cols = section_cols['quarterly_revenue_free_float']
        for idx in range(min(len(q_pat_ff_cols), len(q_rev_ff_cols))):
            if totals_data[q_rev_ff_cols[idx]] and totals_data[q_rev_ff_cols[idx]] != 0:
                results[f'q_pat_ff_{idx}'] = (totals_data[q_pat_ff_cols[idx]] / totals_data[q_rev_ff_cols[idx]]) * 100

    return results


def calculate_qoq_from_totals(totals_data, section_cols):
    """
    QoQ = (Current Quarter / Previous Quarter) - 1
    Calculate for: Q Rev, Q Rev FF, Q PAT, Q PAT FF
    """
    results = {}

    # Quarterly Revenue: Each column compared to next
    if 'quarterly_revenue' in section_cols:
        q_rev_cols = section_cols['quarterly_revenue']
        for idx in range(len(q_rev_cols) - 1):
            current = totals_data[q_rev_cols[idx]]
            previous = totals_data[q_rev_cols[idx + 1]]
            if previous and previous != 0 and current is not None:
                results[f'q_rev_{idx}'] = (current / previous) - 1

    # Quarterly Revenue Free Float
    if 'quarterly_revenue_free_float' in section_cols:
        q_rev_ff_cols = section_cols['quarterly_revenue_free_float']
        for idx in range(len(q_rev_ff_cols) - 1):
            current = totals_data[q_rev_ff_cols[idx]]
            previous = totals_data[q_rev_ff_cols[idx + 1]]
            if previous and previous != 0 and current is not None:
                results[f'q_rev_ff_{idx}'] = (current / previous) - 1

    # Quarterly PAT
    if 'quarterly_pat' in section_cols:
        q_pat_cols = section_cols['quarterly_pat']
        for idx in range(len(q_pat_cols) - 1):
            current = totals_data[q_pat_cols[idx]]
            previous = totals_data[q_pat_cols[idx + 1]]
            if previous and previous != 0 and current is not None:
                results[f'q_pat_{idx}'] = (current / previous) - 1

    # Quarterly PAT Free Float
    if 'quarterly_pat_free_float' in section_cols:
        q_pat_ff_cols = section_cols['quarterly_pat_free_float']
        for idx in range(len(q_pat_ff_cols) - 1):
            current = totals_data[q_pat_ff_cols[idx]]
            previous = totals_data[q_pat_ff_cols[idx + 1]]
            if previous and previous != 0 and current is not None:
                results[f'q_pat_ff_{idx}'] = (current / previous) - 1

    return results


def calculate_yoy_from_totals(totals_data, section_cols):
    """
    YoY = (Current Period / Same Period Last Year) - 1
    Compare period to 4 quarters earlier
    """
    results = {}

    # TTM Revenue
    if 'ttm_revenue' in section_cols:
        ttm_rev_cols = section_cols['ttm_revenue']
        for idx in range(len(ttm_rev_cols) - 4):
            current = totals_data[ttm_rev_cols[idx]]
            year_ago = totals_data[ttm_rev_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'ttm_rev_{idx}'] = (current / year_ago) - 1

    # TTM Revenue Free Float
    if 'ttm_revenue_free_float' in section_cols:
        ttm_rev_ff_cols = section_cols['ttm_revenue_free_float']
        for idx in range(len(ttm_rev_ff_cols) - 4):
            current = totals_data[ttm_rev_ff_cols[idx]]
            year_ago = totals_data[ttm_rev_ff_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'ttm_rev_ff_{idx}'] = (current / year_ago) - 1

    # TTM PAT
    if 'ttm_pat' in section_cols:
        ttm_pat_cols = section_cols['ttm_pat']
        for idx in range(len(ttm_pat_cols) - 4):
            current = totals_data[ttm_pat_cols[idx]]
            year_ago = totals_data[ttm_pat_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'ttm_pat_{idx}'] = (current / year_ago) - 1

    # TTM PAT Free Float
    if 'ttm_pat_free_float' in section_cols:
        ttm_pat_ff_cols = section_cols['ttm_pat_free_float']
        for idx in range(len(ttm_pat_ff_cols) - 4):
            current = totals_data[ttm_pat_ff_cols[idx]]
            year_ago = totals_data[ttm_pat_ff_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'ttm_pat_ff_{idx}'] = (current / year_ago) - 1

    # Quarterly Revenue
    if 'quarterly_revenue' in section_cols:
        q_rev_cols = section_cols['quarterly_revenue']
        for idx in range(len(q_rev_cols) - 4):
            current = totals_data[q_rev_cols[idx]]
            year_ago = totals_data[q_rev_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'q_rev_{idx}'] = (current / year_ago) - 1

    # Quarterly Revenue Free Float
    if 'quarterly_revenue_free_float' in section_cols:
        q_rev_ff_cols = section_cols['quarterly_revenue_free_float']
        for idx in range(len(q_rev_ff_cols) - 4):
            current = totals_data[q_rev_ff_cols[idx]]
            year_ago = totals_data[q_rev_ff_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'q_rev_ff_{idx}'] = (current / year_ago) - 1

    # Quarterly PAT
    if 'quarterly_pat' in section_cols:
        q_pat_cols = section_cols['quarterly_pat']
        for idx in range(len(q_pat_cols) - 4):
            current = totals_data[q_pat_cols[idx]]
            year_ago = totals_data[q_pat_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'q_pat_{idx}'] = (current / year_ago) - 1

    # Quarterly PAT Free Float
    if 'quarterly_pat_free_float' in section_cols:
        q_pat_ff_cols = section_cols['quarterly_pat_free_float']
        for idx in range(len(q_pat_ff_cols) - 4):
            current = totals_data[q_pat_ff_cols[idx]]
            year_ago = totals_data[q_pat_ff_cols[idx + 4]]
            if year_ago and year_ago != 0 and current is not None:
                results[f'q_pat_ff_{idx}'] = (current / year_ago) - 1

    return results


def calculate_6yr_cagr_from_totals(totals_data, section_cols):
    """
    6 Year CAGR = ((Current TTM / 6 Years Ago TTM) ^ (1/6)) - 1
    Calculate for: TTM Revenue, TTM Revenue FF, TTM PAT, TTM PAT FF
    """
    results = {}

    # TTM Revenue (need 24 quarters = 6 years)
    if 'ttm_revenue' in section_cols:
        ttm_rev_cols = section_cols['ttm_revenue']
        if len(ttm_rev_cols) >= 24:
            current = totals_data[ttm_rev_cols[0]]
            six_years_ago = totals_data[ttm_rev_cols[23]]
            if (six_years_ago and six_years_ago > 0 and
                current and current > 0):
                results['ttm_rev'] = (current / six_years_ago) ** (1/6) - 1

    # TTM Revenue Free Float
    if 'ttm_revenue_free_float' in section_cols:
        ttm_rev_ff_cols = section_cols['ttm_revenue_free_float']
        if len(ttm_rev_ff_cols) >= 24:
            current = totals_data[ttm_rev_ff_cols[0]]
            six_years_ago = totals_data[ttm_rev_ff_cols[23]]
            if (six_years_ago and six_years_ago > 0 and
                current and current > 0):
                results['ttm_rev_ff'] = (current / six_years_ago) ** (1/6) - 1

    # TTM PAT
    if 'ttm_pat' in section_cols:
        ttm_pat_cols = section_cols['ttm_pat']
        if len(ttm_pat_cols) >= 24:
            current = totals_data[ttm_pat_cols[0]]
            six_years_ago = totals_data[ttm_pat_cols[23]]
            if (six_years_ago and six_years_ago > 0 and
                current and current > 0):
                results['ttm_pat'] = (current / six_years_ago) ** (1/6) - 1

    # TTM PAT Free Float
    if 'ttm_pat_free_float' in section_cols:
        ttm_pat_ff_cols = section_cols['ttm_pat_free_float']
        if len(ttm_pat_ff_cols) >= 24:
            current = totals_data[ttm_pat_ff_cols[0]]
            six_years_ago = totals_data[ttm_pat_ff_cols[23]]
            if (six_years_ago and six_years_ago > 0 and
                current and current > 0):
                results['ttm_pat_ff'] = (current / six_years_ago) ** (1/6) - 1

    return results


def calculate_pe_pr_from_totals(totals_data, section_cols):
    """
    Current PE = Total Market Cap FF / Total TTM PAT
    Current PR = Total Market Cap FF / Total TTM Revenue
    """
    results = {}

    if 'market_cap_free_float' not in section_cols:
        return results

    market_cap_ff_col = section_cols['market_cap_free_float'][0]
    market_cap_ff = totals_data[market_cap_ff_col]

    # Current PE
    if 'ttm_pat' in section_cols:
        ttm_pat_col = section_cols['ttm_pat'][0]
        ttm_pat = totals_data[ttm_pat_col]
        if ttm_pat and ttm_pat != 0:
            results['current_pe'] = market_cap_ff / ttm_pat

    # Current PR
    if 'ttm_revenue' in section_cols:
        ttm_rev_col = section_cols['ttm_revenue'][0]
        ttm_rev = totals_data[ttm_rev_col]
        if ttm_rev and ttm_rev != 0:
            results['current_pr'] = market_cap_ff / ttm_rev

    return results


def calculate_pe_pr_averages_from_totals(totals_data, section_cols):
    """
    2-Year Avg = Average of ratios over last 8 quarters
    5-Year Avg = Average of ratios over last 20 quarters
    """
    results = {}

    if 'market_cap_free_float' not in section_cols:
        return results

    market_cap_ff_cols = section_cols['market_cap_free_float']

    # Calculate PE for each period
    pe_ratios = []
    if 'ttm_pat' in section_cols:
        ttm_pat_cols = section_cols['ttm_pat']
        for idx in range(min(len(market_cap_ff_cols), len(ttm_pat_cols))):
            mc = totals_data[market_cap_ff_cols[idx]]
            pat = totals_data[ttm_pat_cols[idx]]
            if pat and pat != 0 and mc:
                pe_ratios.append(mc / pat)

    # PE 2-Year Average (8 quarters)
    if len(pe_ratios) >= 8:
        results['pe_2yr_avg'] = sum(pe_ratios[:8]) / 8

    # PE 5-Year Average (20 quarters)
    if len(pe_ratios) >= 20:
        results['pe_5yr_avg'] = sum(pe_ratios[:20]) / 20

    # Calculate PR for each period
    pr_ratios = []
    if 'ttm_revenue' in section_cols:
        ttm_rev_cols = section_cols['ttm_revenue']
        for idx in range(min(len(market_cap_ff_cols), len(ttm_rev_cols))):
            mc = totals_data[market_cap_ff_cols[idx]]
            rev = totals_data[ttm_rev_cols[idx]]
            if rev and rev != 0 and mc:
                pr_ratios.append(mc / rev)

    # PR 2-Year Average (8 quarters)
    if len(pr_ratios) >= 8:
        results['pr_2yr_avg'] = sum(pr_ratios[:8]) / 8

    # PR 5-Year Average (20 quarters)
    if len(pr_ratios) >= 20:
        results['pr_5yr_avg'] = sum(pr_ratios[:20]) / 20

    return results


def calculate_reval_deval_from_totals(pe_pr_metrics):
    """
    Reval/Deval = (Average - Current) / Current
    """
    results = {}

    # PE 2-Year Reval/Deval
    if pe_pr_metrics.get('current_pe') and pe_pr_metrics.get('pe_2yr_avg'):
        results['pe_2yr_reval_deval'] = (
            (pe_pr_metrics['pe_2yr_avg'] - pe_pr_metrics['current_pe']) /
            pe_pr_metrics['current_pe']
        )

    # PE 5-Year Reval/Deval
    if pe_pr_metrics.get('current_pe') and pe_pr_metrics.get('pe_5yr_avg'):
        results['pe_5yr_reval_deval'] = (
            (pe_pr_metrics['pe_5yr_avg'] - pe_pr_metrics['current_pe']) /
            pe_pr_metrics['current_pe']
        )

    # PR 2-Year Reval/Deval
    if pe_pr_metrics.get('current_pr') and pe_pr_metrics.get('pr_2yr_avg'):
        results['pr_2yr_reval_deval'] = (
            (pe_pr_metrics['pr_2yr_avg'] - pe_pr_metrics['current_pr']) /
            pe_pr_metrics['current_pr']
        )

    # PR 5-Year Reval/Deval
    if pe_pr_metrics.get('current_pr') and pe_pr_metrics.get('pr_5yr_avg'):
        results['pr_5yr_reval_deval'] = (
            (pe_pr_metrics['pr_5yr_avg'] - pe_pr_metrics['current_pr']) /
            pe_pr_metrics['current_pr']
        )

    return results


def calculate_pr_10q_extremes_from_totals(totals_data, section_cols):
    """
    10Q PR Low/High from quarterly PR ratios
    """
    results = {}

    if 'market_cap_free_float' not in section_cols or 'quarterly_revenue' not in section_cols:
        return results

    market_cap_ff_cols = section_cols['market_cap_free_float'][:10]
    q_rev_cols = section_cols['quarterly_revenue'][:10]

    pr_ratios = []
    for idx in range(min(len(market_cap_ff_cols), len(q_rev_cols))):
        mc = totals_data[market_cap_ff_cols[idx]]
        rev = totals_data[q_rev_cols[idx]]
        if rev and rev != 0 and mc:
            pr_ratios.append(mc / rev)

    if pr_ratios:
        results['pr_10q_low'] = min(pr_ratios)
        results['pr_10q_high'] = max(pr_ratios)

    return results


def calculate_pe_yield_from_totals(pe_metrics):
    """PE Yield = 1 / PE × 100"""
    if pe_metrics.get('current_pe') and pe_metrics['current_pe'] != 0:
        return (1 / pe_metrics['current_pe']) * 100
    return None


def calculate_growth_from_totals(cagr_metrics):
    """Growth = Revenue 6-Year CAGR (reference)"""
    return cagr_metrics.get('ttm_rev', None)


def get_bond_rate():
    """Fixed bond rate"""
    return 5.117  # 7.31% × 0.7


def build_section_column_mapping(section_start_columns, periods):
    """
    Build detailed column mapping for each section.

    Returns: {
        'ttm_revenue': [col_1, col_2, ...],
        'quarterly_pat': [col_a, col_b, ...],
        ...
    }
    """
    section_cols = {}

    for section_type, start_col in section_start_columns.items():
        # Determine period count based on section type
        if 'ttm' in section_type:
            period_count = len(periods.get('ttm_periods', []))
        elif 'quarterly' in section_type:
            period_count = len(periods.get('quarterly_periods', []))
        elif 'market_cap' in section_type:
            period_count = len(periods.get('market_cap_dates', []))
        else:
            period_count = 0

        # Build column list
        section_cols[section_type] = [start_col + i for i in range(period_count)]

    return section_cols


def create_metric_row(row_def, section_cols, total_columns):
    """
    Create a single metric row from definition.

    row_def: {
        'label': 'PATM',
        'data': {'ttm_pat': 12.5, 'q_pat_0': 11.2, ...},
        'sections': ['ttm_pat', 'quarterly_pat', ...],
        'single_value': False  # True for metrics with one value
    }
    """
    if row_def is None:  # Blank row
        return [''] * total_columns

    metric_row = [''] * total_columns
    metric_row[0] = row_def['label']

    if row_def.get('single_value'):
        # Single value (Current PE, Bond Rate, etc.)
        value = row_def['data'].get('value')
        if value is not None:
            first_section = row_def['sections'][0]
            if first_section in section_cols:
                col_idx = section_cols[first_section][0]
                metric_row[col_idx] = value
    else:
        # Multi-column (PATM, QoQ, etc.)
        data = row_def['data']

        # Map section names to abbreviated key prefixes used by calculation functions
        section_key_map = {
            'quarterly_revenue': 'q_rev',
            'quarterly_revenue_free_float': 'q_rev_ff',
            'quarterly_pat': 'q_pat',
            'quarterly_pat_free_float': 'q_pat_ff',
            'ttm_revenue': 'ttm_rev',
            'ttm_revenue_free_float': 'ttm_rev_ff',
            'ttm_pat': 'ttm_pat',
            'ttm_pat_free_float': 'ttm_pat_ff',
        }

        for section in row_def['sections']:
            if section in section_cols:
                cols = section_cols[section]
                key_prefix = section_key_map.get(section, section)

                for idx, col_idx in enumerate(cols):
                    # Try different key formats (both full section names and abbreviated)
                    key_formats = [
                        section,              # 'ttm_pat' (exact section match)
                        f'{section}_{idx}',   # 'quarterly_pat_0' (full name with index)
                        key_prefix,           # 'q_pat' (abbreviated)
                        f'{key_prefix}_{idx}' # 'q_pat_0' (abbreviated with index)
                    ]
                    for key in key_formats:
                        if key in data:
                            metric_row[col_idx] = data[key]
                            break

    return metric_row
