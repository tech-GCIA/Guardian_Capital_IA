# Portfolio Analysis Metrics - Complete Fix Plan

**Created**: 2026-01-21
**Project**: Guardian Capital Investment Advisor (GCIA)
**Issue**: Portfolio metrics calculations incorrect - using weighted averages instead of TOTALS

---

## Executive Summary

### The Problem
Current system calculates portfolio metrics using **weighted averages** of individual stock metrics. Excel formulas reveal all calculations must use **simple sums (TOTALS)** of underlying stock financials.

**Example - PATM**:
```python
# CURRENT (WRONG):
portfolio_patm = Σ(stock_patm × weight)

# REQUIRED (CORRECT):
total_ttm_pat = Σ(stock_ttm_pat)          # Simple sum
total_ttm_revenue = Σ(stock_ttm_revenue)  # Simple sum
portfolio_patm = (total_ttm_pat / total_ttm_revenue) × 100
```

This affects ~15 of 22 metrics and produces **fundamentally incorrect results**.

### The Solution
**Dual-Fix Approach**:
1. Fix database calculations (metrics_calculator.py) - Remove weighted averaging
2. Fix Excel export (enhanced_excel_export.py) - Calculate from TOTALS row directly

---

## Root Cause Analysis

### Issue #1: Wrong Aggregation Method
**File**: `gcia_app/metrics_calculator.py` (lines 1286, 1292-1294, 1308-1311)

```python
# WRONG (current):
portfolio_financials['ttm_pat'] += ttm_data.ttm_pat * weight

# CORRECT (required):
portfolio_financials['ttm_pat'] += ttm_data.ttm_pat  # No weight!
```

### Issue #2: Multi-Column Calculations Missing
PATM has **4 different values** in Excel:
- TTM PAT (one value)
- TTM PAT Free Float (one value)
- Each Quarterly PAT column (multiple values)
- Each Quarterly PAT Free Float column (multiple values)

Database stores only ONE PATM per period - can't handle multi-column metrics.

### Issue #3: Sequential Calculations Not Supported
QoQ compares **consecutive quarters**:
- Column N: `TOTAL(Q1) / TOTAL(Q2) - 1`
- Column O: `TOTAL(Q2) / TOTAL(Q3) - 1`
- Column P: `TOTAL(Q3) / TOTAL(Q4) - 1`

Database stores ONE QoQ - can't handle sequential comparisons.

### Issue #4: Alpha Calculations Missing
Lines 675-683, 1196-1201: Return 0.0 (placeholders)

### Issue #5: Wrong Bond Rate
Line 717: Returns 6.0, should be **5.117** (7.31% × 0.7)

---

## Implementation Plan

### Phase 0: Fix Database Calculations (metrics_calculator.py)

#### 0.1 Fix update_fund_latest_metrics_weighted()
**Lines**: 1228-1342

**Changes**:
```python
# Line 1286 - Remove weight multiplication
# BEFORE:
if market_cap_data and market_cap_data.market_cap:
    portfolio_financials['market_cap'] += market_cap_data.market_cap * weight

# AFTER:
if market_cap_data and market_cap_data.market_cap:
    portfolio_financials['market_cap'] += market_cap_data.market_cap

# Line 1292-1294 - Remove weight multiplication
# BEFORE:
if ttm_data:
    if ttm_data.ttm_revenue:
        portfolio_financials['ttm_revenue'] += ttm_data.ttm_revenue * weight
    if ttm_data.ttm_pat:
        portfolio_financials['ttm_pat'] += ttm_data.ttm_pat * weight

# AFTER:
if ttm_data:
    if ttm_data.ttm_revenue:
        portfolio_financials['ttm_revenue'] += ttm_data.ttm_revenue
    if ttm_data.ttm_pat:
        portfolio_financials['ttm_pat'] += ttm_data.ttm_pat
```

#### 0.2 Fix calculate_portfolio_metrics_all_periods()
**Lines**: 1343-1484

**Changes**: Remove weight multiplication at lines:
- 1406: `market_cap_data.market_cap * weight` → `market_cap_data.market_cap`
- 1418-1420: Remove `* weight` from ttm_revenue and ttm_pat
- 1425-1428: Remove `* weight` from quarterly_revenue and quarterly_pat

#### 0.3 Rename Functions
**Line 1228**:
- `update_fund_latest_metrics_weighted()` → `update_fund_latest_metrics_from_totals()`
- Update all calls to this function (line 316, etc.)
- Update docstrings to say "TOTALS-based" instead of "weighted"

#### 0.4 Implement Alpha Calculations
**Lines 675-683**:

```python
def calculate_alpha_bond_cagr_cached(self, price_data, pe_yield_pct, growth_pct):
    """
    Alpha over Bond CAGR formula from Excel:
    (((100+((100*PE_YIELD)+(100*PE_YIELD)*(1+GROWTH)+...)))/100)^(1/10)-1)-BOND_RATE

    This calculates 10-year compounded return using PE Yield and Growth,
    then subtracts the bond rate to get alpha.
    """
    if not price_data or pe_yield_pct is None or growth_pct is None:
        return 0.0

    # Convert percentages to decimals
    pe_yield = pe_yield_pct / 100
    growth = growth_pct / 100
    bond_rate = 0.05117  # 5.117%

    # Calculate 10-year compounded value
    # Starting with 100, add PE_YIELD each year compounded by growth
    total = 100
    for year in range(10):
        total += 100 * pe_yield * ((1 + growth) ** year)

    # Calculate CAGR
    cagr = (total / 100) ** (1/10) - 1

    # Subtract bond rate to get alpha
    alpha = cagr - bond_rate

    # Return as percentage
    return alpha * 100

def calculate_alpha_absolute_cached(self, price_data, alpha_bond_cagr_pct):
    """
    Alpha Absolute formula: (1 + Alpha_Bond_CAGR)^10 - 1
    This is the total 10-year return.
    """
    if alpha_bond_cagr_pct is None:
        return 0.0

    alpha_bond = alpha_bond_cagr_pct / 100  # Convert to decimal
    return ((1 + alpha_bond) ** 10 - 1) * 100  # Return as percentage
```

**Note**: Call these with PE Yield and Growth from metrics dict:
```python
metrics['alpha_bond_cagr'] = self.calculate_alpha_bond_cagr_cached(
    filtered_price,
    metrics.get('pe_yield'),
    metrics.get('growth_rate')
)
metrics['alpha_absolute'] = self.calculate_alpha_absolute_cached(
    filtered_price,
    metrics.get('alpha_bond_cagr')
)
```

#### 0.5 Fix Bond Rate
**Line 717**:
```python
def calculate_bond_rate_cached(self):
    """Bond rate - 7.31% × 0.7"""
    return 5.117  # Changed from 6.0
```

---

### Phase 1: Create Calculation Functions (enhanced_excel_export.py)

**Location**: After line 869

#### 1. calculate_patm_from_totals()
```python
def calculate_patm_from_totals(totals_data, section_cols):
    """
    PATM = (Total PAT / Total Revenue) × 100
    Calculate for: TTM PAT, TTM PAT FF, Q PAT (each quarter), Q PAT FF (each quarter)

    Args:
        totals_data: List of values from TOTALS row
        section_cols: Dict mapping section types to column indices

    Returns:
        Dict with PATM values: {
            'ttm_pat': value,
            'ttm_pat_ff': value,
            'q_pat_0': value,  # For first quarterly column
            ...
        }
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
```

#### 2. calculate_qoq_from_totals()
```python
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
```

#### 3. calculate_yoy_from_totals()
```python
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
```

#### 4. calculate_6yr_cagr_from_totals()
```python
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
```

#### 5-11. Remaining Functions
(Continue with calculate_pe_pr_from_totals, calculate_pe_pr_averages_from_totals, calculate_reval_deval_from_totals, calculate_pr_10q_extremes_from_totals, calculate_pe_yield_from_totals, calculate_growth_from_totals, get_bond_rate)

See detailed implementations in plan agent output above.

---

### Phase 2: Helper Functions (enhanced_excel_export.py)

#### build_section_column_mapping()
Maps section types to column index lists.

#### create_metric_row()
Builds a single metric row array from definition.

(Full implementations in plan agent output)

---

### Phase 3: Rewrite add_portfolio_metric_rows()

**Location**: Lines 870-977

**New signature**:
```python
def add_portfolio_metric_rows(ws, scheme, section_start_columns, periods,
                              total_columns, totals_row_index):
```

**Steps**:
1. Read TOTALS row from worksheet
2. Build section column mapping
3. Call all calculation functions
4. Define 27 metric row definitions
5. Create and append rows

---

### Phase 4: Update Function Calls

**generate_enhanced_portfolio_analysis_excel()** (~line 648):
```python
totals_row_index = ws.max_row
add_portfolio_metric_rows(ws, scheme, generator.section_start_columns,
                         periods, total_columns, totals_row_index)
```

**generate_recalculated_analysis_excel()** (~line 776):
```python
totals_row_index = ws.max_row
add_portfolio_metric_rows(ws, scheme, generator.section_start_columns,
                         periods, total_columns, totals_row_index)
```

---

## Testing Strategy

### Test 1: PATM Verification
1. Export Excel for test fund
2. Manually calculate: TOTAL TTM PAT / TOTAL TTM Revenue × 100
3. Compare with PATM row value
4. **Expected**: Exact match

### Test 2: QoQ Multi-Column
1. Export with 5+ quarterly periods
2. Verify QoQ row has values in ALL quarterly revenue columns
3. Manually verify: TOTAL(Q1)/TOTAL(Q2) - 1
4. **Expected**: Each column correct

### Test 3: 6-Year CAGR
1. Export with 24+ TTM periods
2. Calculate: (TOTAL_latest / TOTAL_6yrsago)^(1/6) - 1
3. **Expected**: Match within 0.01%

### Test 4: PE Calculation Chain
1. Current PE = TOTAL Market Cap FF / TOTAL TTM PAT
2. 2Yr Avg PE = Average of 8 quarterly PE values
3. Reval/Deval = (Avg - Current) / Current
4. **Expected**: All correct

### Test 5: Bond Rate
- Check "Bond Rate" row
- **Expected**: 5.117

### Test 6: Sample Excel Match
- Use same fund as Old Bridge Focused Equity Fund.xlsm
- Compare all metric rows cell-by-cell
- **Expected**: Match within 0.01%

### Test 7: Edge Cases
- Zero revenue (PATM)
- < 24 periods (CAGR)
- Negative PAT
- **Expected**: No errors, blank cells

### Test 8: Recalculated Sheet
- Exclude 2 stocks
- Verify TOTALS differ
- Verify metrics reflect filtered TOTALS
- **Expected**: Correct recalculation

---

## Files to Modify

### Primary Files

**1. gcia_app/metrics_calculator.py**
- Lines 1286, 1292-1294: Remove `* weight`
- Lines 1406, 1418-1420, 1425-1428: Remove `* weight`
- Lines 675-683: Implement Alpha calculations
- Line 717: Change 6.0 → 5.117
- Line 1228: Rename function
- Update docstrings
- **Estimate**: ~90 lines

**2. gcia_app/enhanced_excel_export.py**
- Add 11 calculation functions
- Add 2 helper functions
- Rewrite add_portfolio_metric_rows()
- Update 2 function calls
- **Estimate**: ~550 lines

### Secondary Files

**3. gcia_app/models.py**
- Update PortfolioMetricsLog docstring
- **Estimate**: ~3 lines

**4. gcia_app/add_metric_rows_function.py**
- Add deprecation comment
- **Estimate**: ~5 lines

**5. CLAUDE.md**
- Document methodology change
- **Estimate**: ~15 lines

---

## Key Formula Reference

All formulas use TOTALS row values (row 36 in sample file):

```
TOTALS: =SUBTOTAL(9, U$9:U$34)
PATM: =EP36/BX36
QoQ: =HH36/HI36-1
YoY: =BX36/CB36-1
6yr CAGR: =(DG36/DV36)^(1/6)-1
Current PE: =AV36/$FY$36
PE 2yr Avg: =AVERAGE($AW42:$BF42)
PE Reval/Deval: =(AV43-$AV$42)/$AV$42
PR 10Q Low: =MIN(AX48:BG48)
PE Yield: =1/AV42
Growth: =DG40
Bond Rate: 7.31%*0.7 = 5.117
Alpha Bond CAGR: (((100+((100*PE_YIELD)+(100*PE_YIELD)*(1+GROWTH)+...)))/100)^(1/10)-1)-BOND_RATE
Alpha Absolute: =(1+Alpha_Bond_CAGR)^10-1
```

---

## Implementation Checklist

- [ ] Phase 0: Fix database calculations
  - [ ] 0.1: Fix update_fund_latest_metrics_weighted()
  - [ ] 0.2: Fix calculate_portfolio_metrics_all_periods()
  - [ ] 0.3: Rename functions
  - [ ] 0.4: Implement Alpha calculations
  - [ ] 0.5: Fix bond rate
- [ ] Phase 1: Create calculation functions in Excel export
  - [ ] 1: calculate_patm_from_totals()
  - [ ] 2: calculate_qoq_from_totals()
  - [ ] 3: calculate_yoy_from_totals()
  - [ ] 4: calculate_6yr_cagr_from_totals()
  - [ ] 5: calculate_pe_pr_from_totals()
  - [ ] 6: calculate_pe_pr_averages_from_totals()
  - [ ] 7: calculate_reval_deval_from_totals()
  - [ ] 8: calculate_pr_10q_extremes_from_totals()
  - [ ] 9: calculate_pe_yield_from_totals()
  - [ ] 10: calculate_growth_from_totals()
  - [ ] 11: get_bond_rate()
- [ ] Phase 2: Create helper functions
  - [ ] build_section_column_mapping()
  - [ ] create_metric_row()
- [ ] Phase 3: Rewrite add_portfolio_metric_rows()
- [ ] Phase 4: Update function calls
- [ ] Phase 5: Update add_metric_rows_function.py
- [ ] Phase 6: Documentation updates
- [ ] Testing: Run all 8 tests

---

## Notes

- Alpha formula is complex and may need adjustment after testing
- First few exports should be compared cell-by-cell with sample file
- Database changes affect all future metric calculations
- Excel changes only affect Excel exports (not API if any)

---

**End of Plan**
