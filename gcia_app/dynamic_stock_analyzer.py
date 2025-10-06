"""
Dynamic Stock Sheet Analyzer
============================

This module provides intelligent analysis of Excel stock data sheets with variable column structures.
It replaces the static column mapping approach with dynamic detection of data categories and periods.

Key Features:
- Intelligent header parsing (8-row header structure)
- Dynamic period detection for Market Cap, TTM, and Quarterly data
- Flexible column mapping based on actual file structure
- Support for variable numbers of time periods (3 quarters to 30+ quarters)
"""

import pandas as pd
import re
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any
import logging

logger = logging.getLogger(__name__)

class DynamicStockSheetAnalyzer:
    """
    Analyzes Excel stock data sheets to dynamically detect structure and create flexible column mappings.
    """

    def __init__(self):
        self.column_mapping = {}
        self.detected_periods = {}
        self.data_categories = {}
        self.header_structure = {}

    def analyze_excel_structure_header_driven(self, excel_file) -> Dict[str, Any]:
        """
        NEW: Header-driven analysis that reads Rows 6, 7, and 8 together
        to create a complete column-by-column mapping.

        This method replaces the complex detection logic with simple header reading.

        Args:
            excel_file: Excel file object or path

        Returns:
            dict: Complete column mapping with data types identified from headers
        """
        try:
            # Read all 8 header rows
            headers_df = pd.read_excel(excel_file, nrows=8, header=None)

            logger.info(f"Analyzing Excel structure (header-driven): {headers_df.shape[1]} columns detected")

            # Extract the 3 important header rows
            row_6 = headers_df.iloc[5]  # Category labels (e.g., "Market Cap (in crores)")
            row_7 = headers_df.iloc[6]  # Subcategory labels
            row_8 = headers_df.iloc[7]  # Period values or column names

            # Build complete column-by-column mapping
            complete_column_mapping = {}

            for col_idx in range(len(row_8)):
                category = str(row_6[col_idx]).strip() if pd.notna(row_6[col_idx]) else ''
                subcategory = str(row_7[col_idx]).strip() if pd.notna(row_7[col_idx]) else ''
                period = row_8[col_idx]

                # Identify data type from category label
                data_type = self._identify_data_type_from_headers(category, subcategory, period, col_idx)

                complete_column_mapping[col_idx] = {
                    'column_index': col_idx,
                    'category': category,
                    'subcategory': subcategory,
                    'period': period,
                    'data_type': data_type,
                    'is_separator': self._is_separator_column(category, subcategory, period)
                }

            # Store header rows for later use
            header_structure = {
                'row_1': headers_df.iloc[0].tolist(),
                'row_2': headers_df.iloc[1].tolist(),
                'row_3': headers_df.iloc[2].tolist(),
                'row_4': headers_df.iloc[3].tolist(),
                'row_5': headers_df.iloc[4].tolist(),
                'row_6': row_6.tolist(),
                'row_7': row_7.tolist(),
                'row_8': row_8.tolist(),
            }

            analysis_results = {
                'total_columns': headers_df.shape[1],
                'complete_column_mapping': complete_column_mapping,
                'header_structure': header_structure,
                'method': 'header_driven'
            }

            logger.info(f"Header-driven analysis completed: {len(complete_column_mapping)} columns mapped")
            return analysis_results

        except Exception as e:
            logger.error(f"Error in header-driven analysis: {str(e)}")
            raise

    def _identify_data_type_from_headers(self, category: str, subcategory: str, period, col_idx: int) -> str:
        """
        Identify the data type of a column based on its header information.

        Args:
            category: Row 6 category label
            subcategory: Row 7 subcategory label
            period: Row 8 period value
            col_idx: Column index

        Returns:
            str: Data type identifier
        """
        # Convert to lowercase for comparison
        cat_lower = category.lower()
        subcat_lower = subcategory.lower()

        # Handle basic info columns (first ~13 columns, no category label)
        if col_idx < 14 and (not category or category == 'nan'):
            return 'basic_info'

        # Market Cap
        if 'market cap' in cat_lower:
            if 'free float' in cat_lower:
                return 'market_cap_free_float'
            else:
                return 'market_cap'

        # TTM Data
        if 'ttm' in cat_lower:
            if 'revenue' in cat_lower or 'net sales' in cat_lower:
                if 'free float' in cat_lower:
                    return 'ttm_revenue_free_float'
                else:
                    return 'ttm_revenue'
            elif 'pat' in cat_lower or 'profit after tax' in cat_lower:
                if 'free float' in cat_lower:
                    return 'ttm_pat_free_float'
                else:
                    return 'ttm_pat'

        # Quarterly Data
        if 'quarterly' in cat_lower:
            if 'revenue' in cat_lower or 'net sales' in cat_lower:
                if 'free float' in cat_lower:
                    return 'quarterly_revenue_free_float'
                else:
                    return 'quarterly_revenue'
            elif 'pat' in cat_lower or 'profit after tax' in cat_lower:
                if 'free float' in cat_lower:
                    return 'quarterly_pat_free_float'
                else:
                    return 'quarterly_pat'

        # Annual Ratios
        if 'roce' in cat_lower:
            return 'roce'
        if 'roe' in cat_lower:
            return 'roe'
        if 'retention' in cat_lower or 'dividend' in cat_lower:
            return 'retention'

        # Price Data
        if 'share price' in cat_lower:
            return 'share_price'

        # PR ratio - flexible matching (standalone "PR" or "P/R")
        if 'price to revenue' in cat_lower:
            return 'pr_ratio'
        elif cat_lower.strip() == 'pr' or cat_lower.strip() == 'p/r':
            return 'pr_ratio'
        elif 'pr' in cat_lower and 'ratio' in cat_lower:
            return 'pr_ratio'

        # PE ratio - flexible matching (standalone "PE" or "P/E")
        if 'price to earnings' in cat_lower:
            return 'pe_ratio'
        elif cat_lower.strip() == 'pe' or cat_lower.strip() == 'p/e':
            return 'pe_ratio'
        elif 'pe' in cat_lower and 'ratio' in cat_lower:
            return 'pe_ratio'

        # Identifiers (last columns)
        if 'bse code' in cat_lower or 'bse' in str(period).lower():
            return 'bse_code'
        if 'nse' in cat_lower or 'nse' in str(period).lower():
            return 'nse_symbol'
        if 'isin' in cat_lower or 'isin' in str(period).lower():
            return 'isin'

        # Unknown or separator
        return 'unknown'

    def _is_separator_column(self, category: str, subcategory: str, period) -> bool:
        """
        Determine if a column is a separator (empty column).

        Args:
            category: Row 6 value
            subcategory: Row 7 value
            period: Row 8 value

        Returns:
            bool: True if this is a separator column
        """
        # Check if all header values are empty/null
        is_cat_empty = pd.isna(category) or category == '' or category == 'nan'
        is_subcat_empty = pd.isna(subcategory) or subcategory == '' or subcategory == 'nan'
        is_period_empty = pd.isna(period) or period == '' or str(period) == 'nan'

        return is_cat_empty and is_subcat_empty and is_period_empty

    def analyze_excel_structure(self, excel_file) -> Dict[str, Any]:
        """
        Main analysis method that parses the Excel file structure and returns comprehensive mapping.

        Args:
            excel_file: Excel file object or path

        Returns:
            dict: Complete analysis results including column mappings and detected periods
        """
        try:
            # Read header rows (first 8 rows)
            headers_df = pd.read_excel(excel_file, nrows=8, header=None)

            # Read data portion to understand total columns
            data_df = pd.read_excel(excel_file, skiprows=8, nrows=5)  # Sample data rows

            logger.info(f"Analyzing Excel structure: {headers_df.shape[1]} columns detected")

            # Store header structure
            self.header_structure = self._parse_header_rows(headers_df)

            # Detect basic column structure
            basic_columns = self._detect_basic_columns(headers_df)

            # Detect data category regions
            category_regions = self._detect_data_categories(headers_df)

            # Detect periods within each category
            periods_mapping = self._detect_periods_in_categories(headers_df, category_regions)

            # Build comprehensive column mapping
            column_mapping = self._build_dynamic_column_mapping(
                basic_columns, category_regions, periods_mapping
            )

            # Validate structure
            validation_results = self._validate_detected_structure(headers_df, column_mapping)

            analysis_results = {
                'total_columns': headers_df.shape[1],
                'header_structure': self.header_structure,
                'basic_columns': basic_columns,
                'category_regions': category_regions,
                'periods_mapping': periods_mapping,
                'column_mapping': column_mapping,
                'validation': validation_results,
                'analyzer_instance': self
            }

            logger.info(f"Analysis completed successfully: {len(category_regions)} data categories detected")
            return analysis_results

        except Exception as e:
            logger.error(f"Error analyzing Excel structure: {str(e)}")
            raise

    def _parse_header_rows(self, headers_df: pd.DataFrame) -> Dict[str, List]:
        """
        Parse the 8-row header structure and extract information from each row.

        Args:
            headers_df: DataFrame containing the header rows

        Returns:
            dict: Parsed header information for each row
        """
        header_structure = {}

        for row_idx in range(min(8, len(headers_df))):
            row_data = headers_df.iloc[row_idx].tolist()
            header_structure[f'row_{row_idx + 1}'] = row_data

        return header_structure

    def _detect_basic_columns(self, headers_df: pd.DataFrame) -> Dict[str, int]:
        """
        Detect basic information columns (Company Name, Accord Code, Sector, etc.).

        Args:
            headers_df: DataFrame containing the header rows

        Returns:
            dict: Mapping of basic column names to their positions
        """
        basic_columns = {}

        # Check multiple header rows for basic column identifiers
        for row_idx in range(len(headers_df)):
            row_data = headers_df.iloc[row_idx].astype(str).str.lower()

            for col_idx, cell_value in enumerate(row_data):
                if pd.isna(cell_value) or cell_value == 'nan':
                    continue

                # Map common basic column patterns (order matters - more specific first)
                if any(pattern in cell_value for pattern in ['company name']):
                    basic_columns['company_name'] = col_idx
                elif 'accord code' in cell_value:  # More specific for Accord Code
                    basic_columns['accord_code'] = col_idx
                elif cell_value == 'sector':
                    basic_columns['sector'] = col_idx
                elif cell_value in ['cap', 'large cap', 'mid cap', 'small cap']:
                    basic_columns['cap'] = col_idx
                elif 'free float' in cell_value and 'market cap' not in cell_value:
                    basic_columns['free_float'] = col_idx
                elif '6 year cagr' in cell_value and 'revenue' in cell_value:
                    basic_columns['revenue_6yr_cagr'] = col_idx
                elif 'ttm revenue' in cell_value and 'free float' not in cell_value:
                    basic_columns['revenue_ttm'] = col_idx
                elif '6 year cagr' in cell_value and 'pat' in cell_value:
                    basic_columns['pat_6yr_cagr'] = col_idx
                elif 'ttm pat' in cell_value and 'free float' not in cell_value:
                    basic_columns['pat_ttm'] = col_idx
                elif 'current' in cell_value and len(cell_value) < 15:  # Avoid long descriptions
                    basic_columns['current'] = col_idx
                elif '2 yr avg' in cell_value or '2 year avg' in cell_value:
                    basic_columns['two_yr_avg'] = col_idx
                elif 'reval' in cell_value or 'deval' in cell_value:
                    basic_columns['reval_deval'] = col_idx
                # Stock identifiers - explicit detection
                elif cell_value == 'bse code':
                    basic_columns['bse_code'] = col_idx
                elif cell_value in ['nse code', 'nse symbol']:
                    basic_columns['nse_symbol'] = col_idx
                elif cell_value == 'isin':
                    basic_columns['isin'] = col_idx

        logger.info(f"Detected basic columns: {list(basic_columns.keys())}")
        return basic_columns

    def _detect_data_categories(self, headers_df: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
        """
        Detect major data category regions (Market Cap, TTM, Quarterly, etc.).

        Args:
            headers_df: DataFrame containing the header rows

        Returns:
            dict: Information about each detected data category region
        """
        category_regions = {}

        # Analyze row 6 (index 5) which typically contains category descriptions
        if len(headers_df) > 5:
            row_6 = headers_df.iloc[5].astype(str).str.lower()

            current_category = None
            category_start = None

            for col_idx, cell_value in enumerate(row_6):
                if pd.isna(cell_value) or cell_value == 'nan':
                    continue

                # Detect category starts
                new_category = None
                if 'market cap' in cell_value and 'free float' not in cell_value:
                    new_category = 'market_cap'
                elif 'market cap' in cell_value and 'free float' in cell_value:
                    new_category = 'market_cap_free_float'
                elif 'ttm revenue' in cell_value and 'free float' not in cell_value:
                    new_category = 'ttm_revenue'
                elif 'ttm revenue' in cell_value and 'free float' in cell_value:
                    new_category = 'ttm_revenue_free_float'
                elif 'ttm pat' in cell_value and 'free float' not in cell_value:
                    new_category = 'ttm_pat'
                elif 'ttm pat' in cell_value and 'free float' in cell_value:
                    new_category = 'ttm_pat_free_float'
                elif 'quarterly' in cell_value and 'revenue' in cell_value and 'free float' not in cell_value:
                    new_category = 'quarterly_revenue'
                elif 'quarterly' in cell_value and 'revenue' in cell_value and 'free float' in cell_value:
                    new_category = 'quarterly_revenue_free_float'
                elif 'quarterly' in cell_value and 'pat' in cell_value and 'free float' not in cell_value:
                    new_category = 'quarterly_pat'
                elif 'quarterly' in cell_value and 'pat' in cell_value and 'free float' in cell_value:
                    new_category = 'quarterly_pat_free_float'
                elif 'roce' in cell_value:
                    new_category = 'annual_roce'
                elif 'roe' in cell_value:
                    new_category = 'annual_roe'
                elif 'retention' in cell_value or 'dividend' in cell_value:
                    new_category = 'annual_retention'
                elif 'share price' in cell_value:
                    new_category = 'share_price'
                elif ('pr' in cell_value and 'ratio' in cell_value) or ('price' in cell_value and 'revenue' in cell_value):
                    new_category = 'pr_ratio'
                elif 'pe' in cell_value and 'ratio' in cell_value:
                    new_category = 'pe_ratio'

                # Handle category transitions
                if new_category and new_category != current_category:
                    # Close previous category
                    if current_category and category_start is not None:
                        category_regions[current_category] = {
                            'start_col': category_start,
                            'end_col': col_idx - 1,
                            'name': current_category
                        }

                    # Start new category
                    current_category = new_category
                    category_start = col_idx

            # Close final category
            if current_category and category_start is not None:
                category_regions[current_category] = {
                    'start_col': category_start,
                    'end_col': len(row_6) - 1,
                    'name': current_category
                }

        logger.info(f"Detected data categories: {list(category_regions.keys())}")
        return category_regions

    def _detect_periods_in_categories(self, headers_df: pd.DataFrame, category_regions: Dict) -> Dict[str, List]:
        """
        Detect specific periods/dates within each data category and map them to exact column positions.

        Args:
            headers_df: DataFrame containing the header rows
            category_regions: Information about detected category regions

        Returns:
            dict: Mapping of categories to their detected periods with column positions
        """
        periods_mapping = {}

        # Check multiple header rows for period information
        for category_name, region_info in category_regions.items():
            start_col = region_info['start_col']
            end_col = region_info['end_col']
            detected_periods = []
            period_columns = {}  # Map periods to their actual column positions

            # Check rows 3, 4, 7, 8 (indices 2, 3, 6, 7) for period information
            for row_idx in [2, 3, 6, 7]:
                if row_idx < len(headers_df):
                    row_data = headers_df.iloc[row_idx]

                    for col_idx in range(start_col, min(end_col + 1, len(row_data))):
                        cell_value = str(row_data.iloc[col_idx])

                        if pd.isna(cell_value) or cell_value == 'nan':
                            continue

                        # Detect different period formats
                        period = self._parse_period_format(cell_value)
                        if period and period not in detected_periods:
                            detected_periods.append(period)
                            period_columns[period] = col_idx

            periods_mapping[category_name] = detected_periods

            # Store column mapping for later use
            if not hasattr(self, 'period_column_mapping'):
                self.period_column_mapping = {}
            self.period_column_mapping[category_name] = period_columns

            logger.info(f"Category '{category_name}': {len(detected_periods)} periods detected at columns {list(period_columns.values())}")

        return periods_mapping

    def _parse_period_format(self, cell_value) -> Optional[str]:
        """
        Parse various period formats (dates, YYYYMM, financial years, etc.).
        Handles both datetime objects and string values from pandas Excel parsing.

        Args:
            cell_value: String or datetime object from header cell

        Returns:
            str or None: Standardized period string or None if not a valid period
        """
        # Handle datetime objects from pandas
        if hasattr(cell_value, 'strftime'):  # datetime object
            return cell_value.strftime('%Y-%m-%d')

        # Convert to string and validate
        cell_value = str(cell_value)
        if not cell_value or cell_value == 'nan' or len(cell_value) < 4:
            return None

        # Date formats: YYYY-MM-DD, DD-MM-YYYY, etc.
        date_patterns = [
            r'\d{4}-\d{2}-\d{2}',  # 2024-12-31
            r'\d{2}-\d{2}-\d{4}',  # 31-12-2024
            r'\d{4}/\d{2}/\d{2}',  # 2024/12/31
            r'\d{2}/\d{2}/\d{4}',  # 31/12/2024
        ]

        for pattern in date_patterns:
            if re.match(pattern, cell_value):
                try:
                    # Try to parse as date
                    for fmt in ['%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y']:
                        try:
                            parsed_date = datetime.strptime(cell_value, fmt)
                            return parsed_date.strftime('%Y-%m-%d')
                        except ValueError:
                            continue
                except:
                    pass

        # YYYYMM format
        if re.match(r'\d{6}', cell_value):
            try:
                year = int(cell_value[:4])
                month = int(cell_value[4:6])
                if 1900 <= year <= 2100 and 1 <= month <= 12:
                    return cell_value
            except ValueError:
                pass

        # Financial year format: 2023-24, FY2024, etc.
        fy_patterns = [
            r'(\d{4})-(\d{2})',
            r'FY\s*(\d{4})',
            r'(\d{4})\s*-\s*(\d{2})',
        ]

        for pattern in fy_patterns:
            match = re.match(pattern, cell_value, re.IGNORECASE)
            if match:
                return cell_value

        # Quarter references: L, L-1, L-2, etc.
        if re.match(r'L(-\d+)?$', cell_value, re.IGNORECASE):
            return cell_value

        return None

    def _build_dynamic_column_mapping(self, basic_columns: Dict, category_regions: Dict, periods_mapping: Dict) -> Dict:
        """
        Build comprehensive column mapping combining basic columns and detected categories.
        Uses actual column positions for each period instead of assuming consecutive layout.

        Args:
            basic_columns: Mapping of basic column positions
            category_regions: Information about category regions
            periods_mapping: Detected periods for each category

        Returns:
            dict: Complete column mapping for dynamic processing
        """
        column_mapping = {
            'basic_info': basic_columns,
            'data_categories': category_regions,
            'periods': periods_mapping,
            'time_series_ranges': {}
        }

        # Build time series ranges for each category using actual column positions
        for category_name, region_info in category_regions.items():
            periods = periods_mapping.get(category_name, [])
            period_columns = getattr(self, 'period_column_mapping', {}).get(category_name, {})

            if periods and period_columns:
                # Use actual column positions for each period
                actual_columns = [period_columns.get(period) for period in periods if period in period_columns]
                actual_columns = [col for col in actual_columns if col is not None]

                column_mapping['time_series_ranges'][category_name] = {
                    'start_col': region_info['start_col'],
                    'end_col': region_info['end_col'],
                    'periods': periods,
                    'period_count': len(periods),
                    'actual_columns': actual_columns,  # Add actual column positions
                    'period_column_map': period_columns  # Map each period to its column
                }
            else:
                # Fallback to original logic for categories without explicit column mapping
                start_col = region_info['start_col']
                column_mapping['time_series_ranges'][category_name] = {
                    'start_col': start_col,
                    'end_col': start_col + len(periods) - 1 if periods else start_col,
                    'periods': periods,
                    'period_count': len(periods)
                }

        return column_mapping

    def _validate_detected_structure(self, headers_df: pd.DataFrame, column_mapping: Dict) -> Dict[str, Any]:
        """
        Validate the detected structure for consistency and completeness.

        Args:
            headers_df: DataFrame containing the header rows
            column_mapping: The generated column mapping

        Returns:
            dict: Validation results and warnings
        """
        validation = {
            'is_valid': True,
            'warnings': [],
            'errors': [],
            'summary': {}
        }

        # Check basic columns coverage
        basic_required = ['company_name', 'accord_code']
        basic_found = column_mapping['basic_info']

        for required_col in basic_required:
            if required_col not in basic_found:
                validation['errors'].append(f"Required basic column '{required_col}' not found")
                validation['is_valid'] = False

        # Check data categories
        expected_categories = ['market_cap', 'ttm_revenue', 'ttm_pat']
        categories_found = list(column_mapping['data_categories'].keys())

        for expected_cat in expected_categories:
            if expected_cat not in categories_found:
                validation['warnings'].append(f"Expected category '{expected_cat}' not found")

        # Check period counts
        for category, time_range in column_mapping['time_series_ranges'].items():
            period_count = time_range['period_count']
            if period_count == 0:
                validation['warnings'].append(f"No periods detected for category '{category}'")
            elif period_count < 3:
                validation['warnings'].append(f"Only {period_count} periods detected for '{category}' - may be insufficient")

        # Summary statistics
        validation['summary'] = {
            'total_columns': headers_df.shape[1],
            'basic_columns_found': len(basic_found),
            'data_categories_found': len(categories_found),
            'total_periods_detected': sum(tr['period_count'] for tr in column_mapping['time_series_ranges'].values())
        }

        return validation

    def get_column_for_category_period(self, category: str, period_index: int) -> Optional[int]:
        """
        Get the column index for a specific category and period.

        Args:
            category: Category name (e.g., 'market_cap')
            period_index: Index of the period (0-based)

        Returns:
            int or None: Column index or None if not found
        """
        if category not in self.column_mapping.get('time_series_ranges', {}):
            return None

        range_info = self.column_mapping['time_series_ranges'][category]
        if period_index >= range_info['period_count']:
            return None

        return range_info['start_col'] + period_index

    def get_periods_for_category(self, category: str) -> List[str]:
        """
        Get all detected periods for a specific category.

        Args:
            category: Category name

        Returns:
            list: List of period strings
        """
        return self.column_mapping.get('periods', {}).get(category, [])