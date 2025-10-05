import pandas as pd
import pdfplumber
import re
from datetime import datetime
from typing import Optional, List, Dict, Any
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ColpensionesPDFExtractor:
    """
    Extract 'Relación de semanas cotizadas' table from Colpensiones pension report PDFs.
    """
    
    def __init__(self):
        # Expected column headers in Spanish
        self.expected_headers = [
            'Identificación aportante',
            'Nombre o razón Social', 
            'Desde',
            'Hasta',
            'Último salario',
            'Semanas',
            'Licencias (Lic.)',
            'Simultáneos (Sim.)',
            'Total'
        ]
        
        # Target column names for the output DataFrame
        self.output_columns = [
            'cont_id', 'cont_name', 'cont_from', 'cont_to', 
            'cont_last_salary', 'cont_weeks', 'cont_deduction_weeks', 
            'sim_weeks', 'net_weeks'
        ]
        
        # Performance optimizations: Pre-compile regex patterns
        self._compiled_patterns = {
            'date_dd_mm_yyyy': re.compile(r'(\d{1,2})/(\d{1,2})/(\d{4})'),
            'date_dd_mm_yyyy_dash': re.compile(r'(\d{1,2})-(\d{1,2})-(\d{4})'),
            'date_yyyy_mm_dd': re.compile(r'(\d{4})-(\d{1,2})-(\d{1,2})'),
            'summary_numeric_comma': re.compile(r'(\d+,\d{2})'),
            'summary_numeric_digits': re.compile(r'(\d+)'),
            'header_flag': re.compile(r'^\s*\[\d+\]\s*$'),
            'colombian_number_full': re.compile(r'(\d{1,3}(?:\.\d{3})*,\d{2})'),
            'colombian_number_4plus': re.compile(r'(\d{4,},\d{2})'),
            'colombian_number_simple': re.compile(r'(\d+,\d{2})'),
            'colombian_number_no_decimal': re.compile(r'(\d{1,3}(?:\.\d{3})*)'),
            'colombian_number_4plus_no_decimal': re.compile(r'(\d{4,})'),
        }
        
        # Note: Header matching optimization maintains original logic for accuracy
    
    def normalize_date(self, date_str: str) -> Optional[str]:
        """
        Normalize date string to YYYY-MM-DD format.
        Handles various Colombian date formats. Optimized with pre-compiled patterns.
        """
        if not date_str or date_str.strip() in ['', '--', 'N/A']:
            return None
            
        date_str = date_str.strip()
        
        # Use pre-compiled patterns for better performance
        patterns = [
            (self._compiled_patterns['date_dd_mm_yyyy'], False),  # DD/MM/YYYY
            (self._compiled_patterns['date_dd_mm_yyyy_dash'], False),  # DD-MM-YYYY
            (self._compiled_patterns['date_yyyy_mm_dd'], True),  # YYYY-MM-DD
        ]
        
        for pattern, is_yyyy_first in patterns:
            match = pattern.match(date_str)
            if match:
                groups = match.groups()
                if len(groups) == 3:
                    if is_yyyy_first:  # YYYY-MM-DD
                        year, month, day = groups
                    else:  # DD/MM/YYYY or DD-MM-YYYY
                        day, month, year = groups
                    
                    try:
                        # Validate and format the date
                        date_obj = datetime(int(year), int(month), int(day))
                        return date_obj.strftime('%Y-%m-%d')
                    except ValueError:
                        continue
        
        logger.warning(f"Could not parse date: {date_str}")
        return None
    
    def clean_salary(self, salary_str: str) -> Optional[float]:
        """
        Clean salary string and convert to float using Colombian number formatting.
        In Colombian format: dots are thousand separators, commas are decimal separators.
        Examples: $1.235.000 -> 1235000.00, $1.235.000,50 -> 1235000.50
        """
        if not salary_str or salary_str.strip() in ['', '--', 'N/A', '0']:
            return None
            
        # Remove currency symbols and spaces
        cleaned = re.sub(r'[$,\s]', '', str(salary_str).strip())
        
        # Handle Colombian number format for salaries
        # Dots are thousand separators, commas are decimal separators
        if ',' in cleaned and '.' in cleaned:
            # Both comma and dot present: dot=thousands, comma=decimal
            # Example: 1.235.000,50 -> 1235000.50
            parts = cleaned.split(',')
            if len(parts) == 2:
                integer_part = parts[0].replace('.', '')  # Remove thousand separators
                decimal_part = parts[1]
                cleaned = integer_part + '.' + decimal_part
        elif ',' in cleaned and '.' not in cleaned:
            # Only comma present: could be decimal separator
            # Example: 1235000,50 -> 1235000.50
            parts = cleaned.split(',')
            if len(parts) == 2 and len(parts[1]) <= 2:
                # Likely decimal separator: 1-2 digits after comma
                cleaned = parts[0] + '.' + parts[1]
            else:
                # Likely thousand separator: remove comma
                cleaned = cleaned.replace(',', '')
        elif '.' in cleaned and ',' not in cleaned:
            # Only dot present: could be thousand separator or decimal
            # For salaries, if it's a large number, dots are likely thousand separators
            parts = cleaned.split('.')
            if len(parts) == 2 and len(parts[1]) <= 2:
                # Likely decimal separator: 1-2 digits after dot
                # Keep as is: 1235000.50
                pass
            else:
                # Likely thousand separators: remove dots
                # Example: 1.235.000 -> 1235000
                cleaned = cleaned.replace('.', '')
        
        try:
            return round(float(cleaned), 2)
        except ValueError:
            logger.warning(f"Could not parse salary: {salary_str} -> {cleaned}")
            return None
    
    def clean_numeric_colombian(self, value_str: str) -> Optional[float]:
        """
        Clean numeric values using Colombian number formatting.
        In Colombia: commas are decimal separators, dots are thousand separators.
        Examples: 1.234,56 = 1234.56, 1,5 = 1.5
        """
        if not value_str or str(value_str).strip() in ['', '--', 'N/A', '0']:
            return None
            
        value_str = str(value_str).strip()
        
        # Remove any spaces
        cleaned = value_str.replace(' ', '')
        
        # Handle Colombian number format: dots as thousand separators, commas as decimal
        if ',' in cleaned and '.' in cleaned:
            # Both comma and dot present: dot=thousands, comma=decimal
            # Example: 1.234,56 -> 1234.56
            parts = cleaned.split(',')
            if len(parts) == 2:
                integer_part = parts[0].replace('.', '')  # Remove thousand separators
                decimal_part = parts[1]
                cleaned = integer_part + '.' + decimal_part
        elif ',' in cleaned and '.' not in cleaned:
            # Only comma present: could be decimal separator or thousand separator
            # Check if it looks like a decimal (1-3 digits after comma)
            parts = cleaned.split(',')
            if len(parts) == 2 and len(parts[1]) <= 3:
                # Likely decimal separator: 1,5 -> 1.5
                cleaned = parts[0] + '.' + parts[1]
            else:
                # Likely thousand separator: 1,234 -> 1234
                cleaned = cleaned.replace(',', '')
        elif '.' in cleaned and ',' not in cleaned:
            # Only dot present: could be decimal or thousand separator
            # If more than 3 digits after dot, it's likely a thousand separator
            parts = cleaned.split('.')
            if len(parts) == 2 and len(parts[1]) > 3:
                # Thousand separator: 1.2345 -> 12345
                cleaned = parts[0] + parts[1]
            # Otherwise keep as decimal separator
        
        try:
            return float(cleaned)
        except ValueError:
            logger.warning(f"Could not parse numeric value: {value_str} -> {cleaned}")
            return None
    
    def clean_numeric(self, value_str: str) -> Optional[float]:
        """
        Clean numeric values (weeks, licenses, etc.).
        Convert to float, handle empty/null values.
        """
        return self.clean_numeric_colombian(value_str)
    
    def find_table_with_headers(self, page) -> Optional[List[List[str]]]:
        """
        Find table on page that matches expected headers.
        Optimized with pre-computed keyword matching.
        """
        tables = page.extract_tables()
        
        for table in tables:
            if not table or len(table) < 2:
                continue
                
            # Check if first row contains our expected headers
            headers = [cell.strip() if cell else '' for cell in table[0]]
            
            # Check if we have 9 columns and matching headers
            if len(headers) == 9:
                # Optimized header matching using pre-computed keywords
                matches = self._count_header_matches(headers)
                
                if matches >= 6:  # At least 6 out of 9 headers should match
                    logger.info(f"Found matching table with {matches}/9 header matches")
                    return table
        
        return None
    
    def _count_header_matches(self, headers: List[str]) -> int:
        """
        Count header matches using optimized keyword lookup.
        Maintains the same logic as the original but with better performance.
        """
        matches = 0
        
        # For each expected header, check if any of the found headers match
        for expected in self.expected_headers:
            expected_lower = expected.lower()
            for header in headers:
                header_lower = header.lower()
                # Check for partial matches (same logic as original)
                if expected_lower in header_lower or header_lower in expected_lower:
                    matches += 1
                    break  # Found a match for this expected header, move to next
        
        return matches
    
    def _has_table_headers_cached(self, tables: List[List[List[str]]]) -> bool:
        """
        Check if tables contain expected headers using cached table data.
        Optimized version of has_table_headers that reuses already extracted tables.
        """
        try:
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                headers = [cell.strip() if cell else '' for cell in table[0]]
                
                if len(headers) == 9:
                    # Use optimized header matching
                    matches = self._count_header_matches(headers)
                    
                    if matches >= 6:  # At least 6 out of 9 headers should match
                        return True
            
            return False
        except Exception as e:
            logger.warning(f"Error checking headers in cached tables: {str(e)}")
            return False
    
    def clean_row_data(self, row: List[str]) -> Dict[str, Any]:
        """
        Clean and convert row data to proper types.
        """
        if len(row) != 9:
            logger.warning(f"Row has {len(row)} columns, expected 9")
            return None
        
        # Clean each field according to its type
        cleaned_data = {
            'cont_id': row[0].strip() if row[0] else None,
            'cont_name': row[1].strip() if row[1] else None,
            'cont_from': self.normalize_date(row[2]),
            'cont_to': self.normalize_date(row[3]),
            'cont_last_salary': self.clean_salary(row[4]),
            'cont_weeks': self.clean_numeric(row[5]),
            'cont_deduction_weeks': self.clean_numeric(row[6]),
            'sim_weeks': self.clean_numeric(row[7]),
            'net_weeks': self.clean_numeric(row[8])
        }
        
        return cleaned_data
    
    def check_for_table_end(self, page, tables_cache=None) -> bool:
        """
        Check if we've reached the end of the table by looking for 'total semanas cotizadas' text.
        Only stop if we can't find any valid data tables on the page.
        Optimized to reuse tables if already extracted.
        """
        try:
            # Use cached tables if provided, otherwise extract them
            if tables_cache is None:
                tables = page.extract_tables()
            else:
                tables = tables_cache
            
            has_valid_table = False
            
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                headers = [cell.strip() if cell else '' for cell in table[0]]
                if len(headers) == 9:
                    # Use optimized header matching
                    matches = self._count_header_matches(headers)
                    
                    if matches >= 6:  # Valid table found
                        has_valid_table = True
                        break
            
            # If we have a valid table, don't stop even if we see the summary text
            if has_valid_table:
                return False
            
            # Only check for end indicators if there's no valid table
            text = page.extract_text()
            if text:
                # Look for variations of the end text
                end_indicators = [
                    'total semanas cotizadas',
                    'total de semanas cotizadas',
                    'total semanas',
                    'resumen de semanas',
                    'total general'
                ]
                
                text_lower = text.lower()
                for indicator in end_indicators:
                    if indicator in text_lower:
                        logger.info(f"Found table end indicator: '{indicator}' and no valid tables on page")
                        return True
        except Exception as e:
            logger.warning(f"Error checking for table end: {str(e)}")
        
        return False
    
    def has_table_headers(self, page) -> bool:
        """
        Check if a page contains the expected table headers.
        Returns True if headers are found, False otherwise.
        Optimized with pre-computed keyword matching.
        """
        try:
            tables = page.extract_tables()
            
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                headers = [cell.strip() if cell else '' for cell in table[0]]
                
                if len(headers) == 9:
                    # Use optimized header matching
                    matches = self._count_header_matches(headers)
                    
                    if matches >= 6:  # At least 6 out of 9 headers should match
                        return True
            
            return False
        except Exception as e:
            logger.warning(f"Error checking headers on page: {str(e)}")
            return False
    
    def extract_summary_values(self, page) -> Dict[str, Optional[float]]:
        """
        Extract summary values from the page text.
        Looks for [26] TOTAL SEMANAS which includes quoted weeks + reported public times - simultaneous weeks.
        
        Returns:
            Dictionary with weeks_total_report and weeks_high_risk values
        """
        summary_values = {
            "weeks_total_report": None,
            "weeks_high_risk": None
        }
        
        try:
            text = page.extract_text()
            if not text:
                return summary_values
            
            # Look for the summary indicators
            lines = text.split('\n')
            
            for i, line in enumerate(lines):
                line = line.strip()
                
                # Look for [26] TOTAL SEMANAS (the comprehensive total) - be more flexible with matching
                if '[26]' in line and 'total semanas' in line.lower():
                    logger.info(f"Found [26] TOTAL SEMANAS line: '{line}'")
                    # The value should be on the next line or within the same line
                    # First try to find it in the same line using simple parsing
                    value = self.extract_summary_numeric(line)
                    if value is not None and value != 26.0:  # Make sure it's not just the [26] flag
                        summary_values["weeks_total_report"] = value
                        logger.info(f"Found total weeks [26] in same line: {value}")
                    else:
                        # Look in the next few lines for the numeric value
                        for j in range(i + 1, min(i + 5, len(lines))):  # Increased range to 5 lines
                            next_line = lines[j].strip()
                            if next_line:
                                logger.info(f"Checking line {j+1}: '{next_line}'")
                                value = self.extract_summary_numeric(next_line)
                                if value is not None and value != 26.0 and value > 100:  # Should be a reasonable total
                                    summary_values["weeks_total_report"] = value
                                    logger.info(f"Found total weeks [26] in line {j+1}: {value}")
                                    break
                
                # Also look for just "TOTAL SEMANAS" without [26] in case the format is different
                elif 'total semanas' in line.lower() and not summary_values["weeks_total_report"]:
                    logger.info(f"Found TOTAL SEMANAS line: '{line}'")
                    value = self.extract_summary_numeric(line)
                    if value is not None and value > 100:
                        summary_values["weeks_total_report"] = value
                        logger.info(f"Found total weeks from TOTAL SEMANAS line: {value}")
                
                # Also look for [11] SEMANAS COTIZADAS CON TARIFA DE ALTO RIESGO (keep this for high risk)
                elif '[11]' in line and 'semanas cotizadas con tarifa de alto riesgo' in line.lower():
                    # The value should be on the next line or within the same line
                    # First try to find it in the same line
                    value = self.extract_numeric_from_line(line)
                    if value is not None and value != 11.0:  # Make sure it's not just the [11] flag
                        summary_values["weeks_high_risk"] = value
                        logger.info(f"Found high risk weeks in same line: {value}")
                    else:
                        # Look in the next few lines for the numeric value
                        for j in range(i + 1, min(i + 3, len(lines))):
                            next_line = lines[j].strip()
                            if next_line:
                                value = self.extract_numeric_from_line(next_line)
                                if value is not None and value != 11.0:
                                    summary_values["weeks_high_risk"] = value
                                    logger.info(f"Found high risk weeks in next line: {value}")
                                    break
        
        except Exception as e:
            logger.warning(f"Error extracting summary values: {str(e)}")
        
        return summary_values
    
    def extract_summary_numeric(self, line: str) -> Optional[float]:
        """
        Extract numeric value from summary lines using simple comma-to-decimal conversion.
        For summary values like "1193,00", just replace comma with decimal point.
        Optimized with pre-compiled patterns.
        """
        try:
            # Skip lines that are just header flags
            if self._compiled_patterns['header_flag'].match(line.strip()):
                return None
            
            # Look for patterns with comma as decimal separator (no thousand separators in summary)
            patterns = [
                self._compiled_patterns['summary_numeric_comma'],  # 1193,00 or 0,00
                self._compiled_patterns['summary_numeric_digits']  # 1193 (no decimal)
            ]
            
            for pattern in patterns:
                matches = pattern.findall(line)
                if matches:
                    # Take the last match (usually the value on the right)
                    value_str = matches[-1]
                    
                    # Skip if it's just a single digit or two digits (likely part of header flag)
                    if len(value_str) <= 2 and value_str.isdigit():
                        continue
                    
                    # Simple conversion: replace comma with decimal point
                    if ',' in value_str:
                        # 1193,00 -> 1193.00
                        standard_value = value_str.replace(',', '.')
                    else:
                        # 1193 -> 1193
                        standard_value = value_str
                    
                    # Convert to float and validate it's a reasonable value
                    result = float(standard_value)
                    
                    # Skip very small values that are likely header flags
                    if result < 0.01:
                        continue
                    
                    return result
            
            return None
            
        except Exception as e:
            logger.warning(f"Error extracting summary numeric from line '{line}': {str(e)}")
            return None
    
    def extract_numeric_from_line(self, line: str) -> Optional[float]:
        """
        Extract numeric value from a line containing summary information.
        Handles Colombian number format (comma as decimal separator).
        Avoids extracting header flag numbers like [10] or [11].
        Optimized with pre-compiled patterns.
        """
        try:
            # Skip lines that are just header flags
            if self._compiled_patterns['header_flag'].match(line.strip()):
                return None
            
            # Look for patterns like "1.184,29" or "0,00"
            # Find the last numeric pattern in the line
            patterns = [
                self._compiled_patterns['colombian_number_full'],      # 1.184,29 or 1.193,00
                self._compiled_patterns['colombian_number_4plus'],     # 1193,00 (4+ digits with comma decimal)
                self._compiled_patterns['colombian_number_simple'],    # 1184,29 or 193,00
                self._compiled_patterns['colombian_number_no_decimal'], # 1.184 (no decimal)
                self._compiled_patterns['colombian_number_4plus_no_decimal'], # 1193 (4+ digits, likely a total)
                self._compiled_patterns['summary_numeric_digits']      # 1184
            ]
            
            for pattern in patterns:
                matches = pattern.findall(line)
                if matches:
                    # Take the last match (usually the value on the right)
                    value_str = matches[-1]
                    
                    # Skip if it's just a single digit or two digits (likely part of header flag like [10], [11], [26])
                    if len(value_str) <= 2 and value_str.isdigit():
                        continue
                    
                    # Convert Colombian format to standard format
                    if ',' in value_str and '.' in value_str:
                        # Both comma and dot: dot=thousands, comma=decimal
                        # Example: 1.336,43 -> 1336.43
                        parts = value_str.split(',')
                        integer_part = parts[0].replace('.', '')  # Remove thousand separators
                        decimal_part = parts[1]
                        standard_value = integer_part + '.' + decimal_part
                    elif ',' in value_str:
                        # Only comma: could be decimal separator or thousand separator
                        # For summary values like "1193,00", comma is decimal separator
                        # For large numbers with comma, it's likely decimal
                        parts = value_str.split(',')
                        if len(parts) == 2 and len(parts[1]) <= 2:
                            # Likely decimal separator: 1193,00 -> 1193.00
                            standard_value = value_str.replace(',', '.')
                        else:
                            # Likely thousand separator: 1,193 -> 1193
                            standard_value = value_str.replace(',', '')
                    elif '.' in value_str:
                        # Only dot: could be thousand separator or decimal
                        # For large numbers, dots are likely thousand separators
                        parts = value_str.split('.')
                        if len(parts) == 2 and len(parts[1]) <= 2:
                            # Likely decimal: 336.43
                            standard_value = value_str
                        else:
                            # Likely thousand separator: 1.336 -> 1336
                            standard_value = value_str.replace('.', '')
                    else:
                        # No separators: just digits
                        standard_value = value_str
                    
                    # Convert to float and validate it's a reasonable value
                    result = float(standard_value)
                    
                    # Skip very small values that are likely header flags
                    if result < 0.01:
                        continue
                    
                    return result
            
            return None
            
        except Exception as e:
            logger.warning(f"Error extracting numeric from line '{line}': {str(e)}")
            return None

    def extract_table_from_pdf(self, pdf_path: str) -> pd.DataFrame:
        """
        Extract the 'Relación de semanas cotizadas' table from PDF.
        Only processes pages that contain table headers to avoid overhead.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            pandas DataFrame with cleaned data
        """
        logger.info(f"Extracting table from PDF: {pdf_path}")
        
        all_rows = []
        table_found = False
        headers_processed = False
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                logger.info(f"PDF has {len(pdf.pages)} pages")
                
                for page_num, page in enumerate(pdf.pages):
                    # Extract tables once per page to avoid redundant calls
                    tables = page.extract_tables()
                    
                    # First, quickly check if this page has table headers using cached tables
                    if not self._has_table_headers_cached(tables):
                        logger.info(f"Page {page_num + 1}: No table headers found, skipping")
                        continue
                    
                    logger.info(f"Page {page_num + 1}: Table headers found, processing...")
                    
                    # Check if we've reached the end of the table using cached tables
                    if self.check_for_table_end(page, tables):
                        logger.info(f"Reached end of table on page {page_num + 1}")
                        break
                    
                    # Process tables using the already extracted tables
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        
                        # Check if this table matches our expected format
                        headers = [cell.strip() if cell else '' for cell in table[0]]
                        
                        if len(headers) == 9:
                            # Use optimized header matching
                            matches = self._count_header_matches(headers)
                            
                            if matches >= 6:  # At least 6 out of 9 headers should match
                                table_found = True
                                logger.info(f"Found matching table on page {page_num + 1} with {matches}/9 header matches")
                                
                                # Process rows
                                start_row = 1 if not headers_processed else 0  # Skip headers on subsequent pages
                                headers_processed = True
                                
                                for row in table[start_row:]:
                                    if row and any(cell for cell in row):  # Skip empty rows
                                        # Check if this row looks like a data row (not a summary/total row)
                                        if self.is_data_row(row):
                                            cleaned_row = self.clean_row_data(row)
                                            if cleaned_row:
                                                all_rows.append(cleaned_row)
                                        else:
                                            logger.info(f"Skipping non-data row: {row[:3]}...")
                                
                                break  # Found our table, move to next page
        
        except Exception as e:
            logger.error(f"Error processing PDF: {str(e)}")
            raise
        
        if not all_rows:
            logger.warning("No data rows found in PDF")
            return pd.DataFrame(columns=self.output_columns)
        
        # Create DataFrame
        df = pd.DataFrame(all_rows)
        
        # Ensure all expected columns are present
        for col in self.output_columns:
            if col not in df.columns:
                df[col] = None
        
        # Reorder columns
        df = df[self.output_columns]
        
        logger.info(f"Successfully extracted {len(df)} rows from multi-page table")
        return df
    
    def is_data_row(self, row: List[str]) -> bool:
        """
        Check if a row is a data row (not a summary/total row or repeated header).
        """
        if not row or len(row) < 3:
            return False
        
        # Check if first column looks like an ID (numeric or alphanumeric)
        first_cell = str(row[0]).strip() if row[0] else ''
        
        # Skip rows that look like totals or summaries
        total_indicators = ['total', 'suma', 'resumen', 'subtotal', 'gran total']
        for indicator in total_indicators:
            if indicator in first_cell.lower():
                return False
        
        # Skip repeated header rows that start with [1] Identificación
        header_indicators = [
            '[1] identificación',
            'identificación aportante',
            'identificacion aportante',  # without accent
            'identificación',
            'identificacion'
        ]
        
        first_cell_lower = first_cell.lower()
        for indicator in header_indicators:
            if indicator in first_cell_lower:
                logger.info(f"Skipping header row: {first_cell}")
                return False
        
        # Skip completely empty rows
        if not first_cell:
            return False
        
        # Be more lenient with ID formats - accept various formats
        # Skip only if it's clearly not a valid ID
        if len(first_cell) < 2:
            return False
        
        # Accept numeric IDs (like 890326878 from the image)
        if first_cell.isdigit() and len(first_cell) >= 6:
            return True
        
        # Accept alphanumeric IDs
        if first_cell.replace(' ', '').replace('-', '').replace('.', '').isalnum():
            return True
        
        # If we can't determine, err on the side of including it
        # (we can filter out bad data later)
        return True
    
    def extract_table_and_summary_from_pdf(self, pdf_path: str) -> tuple[pd.DataFrame, Dict[str, Optional[float]]]:
        """
        Extract both the contribution table and summary values from PDF.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            Tuple of (DataFrame with contribution data, Dictionary with summary values)
        """
        logger.info(f"Extracting table and summary from PDF: {pdf_path}")
        
        all_rows = []
        summary_values = {
            "weeks_total_report": None,
            "weeks_high_risk": None
        }
        table_found = False
        headers_processed = False
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                logger.info(f"PDF has {len(pdf.pages)} pages")
                
                for page_num, page in enumerate(pdf.pages):
                    # Extract tables once per page to avoid redundant calls
                    tables = page.extract_tables()
                    
                    # First, quickly check if this page has table headers using cached tables
                    if not self._has_table_headers_cached(tables):
                        logger.info(f"Page {page_num + 1}: No table headers found, checking for summary values...")
                        
                        # Even if no table headers, check for summary values
                        page_summary = self.extract_summary_values(page)
                        if page_summary["weeks_total_report"] is not None:
                            summary_values["weeks_total_report"] = page_summary["weeks_total_report"]
                            logger.info(f"Found summary values on page {page_num + 1} (no table headers)")
                        if page_summary["weeks_high_risk"] is not None:
                            summary_values["weeks_high_risk"] = page_summary["weeks_high_risk"]
                            logger.info(f"Found high risk weeks on page {page_num + 1} (no table headers)")
                        
                        # Continue to next page if we haven't found table data yet
                        if not table_found:
                            continue
                        else:
                            # If we already found table data, we can stop here
                            break
                    
                    logger.info(f"Page {page_num + 1}: Table headers found, processing...")
                    
                    # Check if we've reached the end of the table using cached tables
                    if self.check_for_table_end(page, tables):
                        logger.info(f"Reached end of table on page {page_num + 1}")
                        
                        # Try to extract summary values from this page (where table ends)
                        page_summary = self.extract_summary_values(page)
                        if page_summary["weeks_total_report"] is not None:
                            summary_values["weeks_total_report"] = page_summary["weeks_total_report"]
                        if page_summary["weeks_high_risk"] is not None:
                            summary_values["weeks_high_risk"] = page_summary["weeks_high_risk"]
                        break
                    
                    # Process tables using the already extracted tables
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        
                        # Check if this table matches our expected format
                        headers = [cell.strip() if cell else '' for cell in table[0]]
                        
                        if len(headers) == 9:
                            # Use optimized header matching
                            matches = self._count_header_matches(headers)
                            
                            if matches >= 6:  # At least 6 out of 9 headers should match
                                table_found = True
                                logger.info(f"Found matching table on page {page_num + 1} with {matches}/9 header matches")
                                
                                # Process rows
                                start_row = 1 if not headers_processed else 0  # Skip headers on subsequent pages
                                headers_processed = True
                                
                                for row in table[start_row:]:
                                    if row and any(cell for cell in row):  # Skip empty rows
                                        # Check if this row looks like a data row (not a summary/total row)
                                        if self.is_data_row(row):
                                            cleaned_row = self.clean_row_data(row)
                                            if cleaned_row:
                                                all_rows.append(cleaned_row)
                                        else:
                                            logger.info(f"Skipping non-data row: {row[:3]}...")
                                
                                # Try to extract summary values from this page
                                page_summary = self.extract_summary_values(page)
                                if page_summary["weeks_total_report"] is not None:
                                    summary_values["weeks_total_report"] = page_summary["weeks_total_report"]
                                if page_summary["weeks_high_risk"] is not None:
                                    summary_values["weeks_high_risk"] = page_summary["weeks_high_risk"]
                                
                                break  # Found our table, move to next page
        
        except Exception as e:
            logger.error(f"Error processing PDF: {str(e)}")
            raise
        
        if not all_rows:
            logger.warning("No data rows found in PDF")
            df = pd.DataFrame(columns=self.output_columns)
        else:
            # Create DataFrame
            df = pd.DataFrame(all_rows)
            
            # Ensure all expected columns are present
            for col in self.output_columns:
                if col not in df.columns:
                    df[col] = None
            
            # Reorder columns
            df = df[self.output_columns]
        
        logger.info(f"Successfully extracted {len(df)} rows and summary values: {summary_values}")
        return df, summary_values

    def save_to_excel(self, df: pd.DataFrame, output_path: str):
        """
        Save DataFrame to Excel file.
        """
        try:
            df.to_excel(output_path, index=False, engine='openpyxl')
            logger.info(f"Data saved to Excel: {output_path}")
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")
            raise

def main():
    """
    Example usage of the ColpensionesPDFExtractor.
    """
    extractor = ColpensionesPDFExtractor()
    
    # Example usage
    pdf_path = "path/to/your/colpensiones_report.pdf"
    output_path = "extracted_pension_data.xlsx"
    
    try:
        # Extract table from PDF
        df = extractor.extract_table_from_pdf(pdf_path)
        
        # Display basic info
        print(f"Extracted {len(df)} rows")
        print("\nFirst few rows:")
        print(df.head())
        
        print("\nData types:")
        print(df.dtypes)
        
        print("\nSummary statistics:")
        print(df.describe())
        
        # Save to Excel
        extractor.save_to_excel(df, output_path)
        
    except Exception as e:
        print(f"Error: {str(e)}")

class ColpensionesPost1995PaymentsExtractor:
    """
    Extract 'DETALLE DE PAGOS EFECTUADOS A PARTIR DE 1995' table from Colpensiones pension report PDFs.
    """
    
    def __init__(self):
        # Expected column headers for post-1995 payments table (exact format from image)
        self.expected_headers = [
            'Identificación Aportante',
            'Nombre o Razón Social',
            'RA',
            'Período',
            'Fecha De Pago',
            'Referencia de Pago',
            'IBC Reportado',
            'Cotización Pagada',
            'Cotización Mora Sin Intereses',
            'Nov.',
            'Días Rep.',
            'Días Cot.',
            'Observación'
        ]
        
        # Alternative header patterns that might appear in the PDF
        self.header_patterns = [
            '[34] Identificación Aportante',
            '[35] Nombre o Razón Social',
            '[36] RA',
            '[37] Período',
            '[38] Fecha De Pago',
            '[39] Referencia de Pago',
            '[40] IBC Reportado',
            '[41] Cotización Pagada',
            '[42] Cotización Mora Sin Intereses',
            '[43] Nov.',
            '[44] Días Rep.',
            '[45] Días Cot.',
            '[46] Observación'
        ]
        
        # Target column names for the output DataFrame
        self.output_columns = [
            'cont_id', 'cont_name', 'cont_period', 'cont_reported_ibc', 
            'cont_reported_days', 'cont_contributed_days'
        ]
        
        # Performance optimizations: Pre-compile regex patterns
        self._compiled_patterns = {
            'period_yyyymm': re.compile(r'(\d{4})(\d{2})'),  # 199712 -> 1997-12
            'currency_value': re.compile(r'\$?\s*([\d.,]+)'),  # $ 471.800 -> 471.800
            'header_flag': re.compile(r'^\s*\[\d+\]\s*$'),
        }
    
    def normalize_period(self, period_str: str) -> Optional[str]:
        """
        Normalize period string to YYYY-MM format.
        Handles YYYYMM format (e.g., "199712" -> "1997-12").
        """
        if not period_str or period_str.strip() in ['', '--', 'N/A']:
            return None
            
        period_str = period_str.strip()
        
        # Handle YYYYMM format
        match = self._compiled_patterns['period_yyyymm'].match(period_str)
        if match:
            year, month = match.groups()
            try:
                # Validate month
                month_int = int(month)
                if 1 <= month_int <= 12:
                    return f"{year}-{month}"
            except ValueError:
                pass
        
        logger.warning(f"Could not parse period: {period_str}")
        return None
    
    def clean_ibc_value(self, ibc_str: str) -> Optional[float]:
        """
        Clean IBC (Ingreso Base de Cotización) value and convert to float.
        Handles Colombian currency format: $ 471.800 -> 471800.00
        """
        if not ibc_str or ibc_str.strip() in ['', '--', 'N/A', '0']:
            return None
            
        # Remove currency symbols and spaces
        cleaned = re.sub(r'[$,\s]', '', str(ibc_str).strip())
        
        # Handle Colombian number format: dots are thousand separators
        if '.' in cleaned:
            # Remove thousand separators (dots)
            cleaned = cleaned.replace('.', '')
        
        try:
            return round(float(cleaned), 2)
        except ValueError:
            logger.warning(f"Could not parse IBC value: {ibc_str} -> {cleaned}")
            return None
    
    def clean_days_value(self, days_str: str) -> Optional[int]:
        """
        Clean days value and convert to integer.
        """
        if not days_str or str(days_str).strip() in ['', '--', 'N/A', '0']:
            return None
            
        try:
            return int(str(days_str).strip())
        except ValueError:
            logger.warning(f"Could not parse days value: {days_str}")
            return None
    
    def find_table_with_headers(self, page) -> Optional[List[List[str]]]:
        """
        Find table on page that matches expected headers for post-1995 payments.
        """
        tables = page.extract_tables()
        
        for table in tables:
            if not table or len(table) < 2:
                continue
                
            # Check if first row contains our expected headers
            headers = [cell.strip() if cell else '' for cell in table[0]]
            
            # Check if we have 13 columns and matching headers
            if len(headers) == 13:
                # Check for partial matches (headers might be slightly different)
                matches = 0
                for expected in self.expected_headers:
                    for header in headers:
                        if expected.lower() in header.lower() or header.lower() in expected.lower():
                            matches += 1
                            break
                
                if matches >= 8:  # At least 8 out of 13 headers should match
                    logger.info(f"Found matching post-1995 payments table with {matches}/13 header matches")
                    return table
        
        return None
    
    def clean_row_data(self, row: List[str]) -> Dict[str, Any]:
        """
        Clean and convert row data to proper types for post-1995 payments.
        Handles both 13-column format and flexible column detection.
        """
        if len(row) < 3:  # Need at least 3 columns for basic data
            return None
        
        cleaned_data = {
            'cont_id': row[0].strip() if row[0] else None,
            'cont_name': row[1].strip() if row[1] else None,
            'cont_period': None,
            'cont_reported_ibc': None,
            'cont_reported_days': None,
            'cont_contributed_days': None
        }
        
        # If we have 13 columns, use the exact mapping from the image
        if len(row) >= 13:
            try:
                # Column mapping based on the image:
                # [34] Identificación Aportante -> row[0]
                # [35] Nombre o Razón Social -> row[1] 
                # [36] RA -> row[2]
                # [37] Período -> row[3]
                # [38] Fecha De Pago -> row[4]
                # [39] Referencia de Pago -> row[5]
                # [40] IBC Reportado -> row[6]
                # [41] Cotización Pagada -> row[7]
                # [42] Cotización Mora Sin Intereses -> row[8]
                # [43] Nov. -> row[9]
                # [44] Días Rep. -> row[10]
                # [45] Días Cot. -> row[11]
                # [46] Observación -> row[12]
                
                cleaned_data['cont_period'] = self.normalize_period(row[3])  # [37] Período
                cleaned_data['cont_reported_ibc'] = self.clean_ibc_value(row[6])  # [40] IBC Reportado
                cleaned_data['cont_reported_days'] = self.clean_days_value(row[10])  # [44] Días Rep.
                cleaned_data['cont_contributed_days'] = self.clean_days_value(row[11])  # [45] Días Cot.
            except (IndexError, TypeError):
                # Fall back to flexible detection if standard mapping fails
                pass
        
        # If standard mapping didn't work or we don't have 13 columns, use flexible detection
        if not cleaned_data['cont_period']:
            # Try to find period (look for YYYYMM format)
            for cell in row:
                if cell and str(cell).strip():
                    cell_str = str(cell).strip()
                    if len(cell_str) == 6 and cell_str.isdigit():
                        cleaned_data['cont_period'] = self.normalize_period(cell_str)
                        break
        
        if not cleaned_data['cont_reported_ibc']:
            # Try to find IBC value (look for currency format)
            for cell in row:
                if cell and str(cell).strip():
                    cell_str = str(cell).strip()
                    if '$' in cell_str or ('.' in cell_str and ',' in cell_str):
                        cleaned_data['cont_reported_ibc'] = self.clean_ibc_value(cell_str)
                        break
        
        if not cleaned_data['cont_reported_days'] or not cleaned_data['cont_contributed_days']:
            # Try to find days values (look for small integers)
            days_found = 0
            for cell in row:
                if cell and str(cell).strip():
                    cell_str = str(cell).strip()
                    if cell_str.isdigit() and 1 <= int(cell_str) <= 31:  # Likely days
                        if days_found == 0:
                            cleaned_data['cont_reported_days'] = self.clean_days_value(cell_str)
                            days_found += 1
                        elif days_found == 1:
                            cleaned_data['cont_contributed_days'] = self.clean_days_value(cell_str)
                            break
        
        # Only return if we have at least ID and name
        if cleaned_data['cont_id'] and cleaned_data['cont_name']:
            return cleaned_data
        else:
            return None
    
    def is_data_row(self, row: List[str]) -> bool:
        """
        Check if a row is a data row (not a summary/total row or repeated header).
        More flexible for payments data.
        """
        if not row or len(row) < 3:
            return False
        
        # Check if first column looks like an ID (numeric)
        first_cell = str(row[0]).strip() if row[0] else ''
        
        # Skip rows that look like totals or summaries
        total_indicators = ['total', 'suma', 'resumen', 'subtotal', 'gran total']
        for indicator in total_indicators:
            if indicator in first_cell.lower():
                return False
        
        # Skip repeated header rows (more comprehensive list)
        header_indicators = [
            '[34] identificación',
            '[35] nombre',
            '[36] ra',
            '[37] período',
            '[38] fecha',
            '[39] referencia',
            '[40] ibc',
            '[41] cotización',
            '[42] mora',
            '[43] nov',
            '[44] días rep',
            '[45] días cot',
            '[46] observación',
            'identificación aportante',
            'identificacion aportante',
            'identificación empleador',
            'identificacion empleador',
            'nombre o razón',
            'nombre o razon',
            'período',
            'periodo',
            'fecha de pago',
            'ibc reportado',
            'días rep',
            'dias rep',
            'días cot',
            'dias cot'
        ]
        
        # Check for header indicators in the first cell
        first_cell_lower = first_cell.lower()
        for indicator in header_indicators:
            if indicator in first_cell_lower:
                logger.info(f"Skipping header row: {first_cell}")
                return False
        
        # Check for header patterns with newlines (like "[34]\nIdentificación\nAportante")
        if '\n' in first_cell:
            # Split by newlines and check if any part contains header indicators
            cell_parts = [part.strip().lower() for part in first_cell.split('\n')]
            for part in cell_parts:
                for indicator in header_indicators:
                    if indicator in part:
                        logger.info(f"Skipping header row with newlines: {first_cell}")
                        return False
        
        # Check if the row contains header patterns in any cell (not just first cell)
        # But only if the cell contains the specific numbered header pattern
        for i, cell in enumerate(row):
            if cell and '\n' in str(cell):
                cell_str = str(cell).strip().lower()
                cell_parts = [part.strip().lower() for part in cell_str.split('\n')]
                for part in cell_parts:
                    # Only skip if it contains the specific numbered header pattern like [34], [35], etc.
                    if any(f'[{num}]' in part for num in range(34, 47)):  # [34] to [46]
                        for indicator in header_indicators:
                            if indicator in part:
                                logger.info(f"Skipping header row (cell {i}): {cell}")
                                return False
        
        # Check for rows that contain multiple header-like patterns (likely header rows)
        header_count = 0
        for cell in row[:3]:  # Check first 3 cells
            if cell:
                cell_str = str(cell).strip().lower()
                for indicator in header_indicators:
                    if indicator in cell_str:
                        header_count += 1
                        break
        
        # If multiple cells contain header indicators, it's likely a header row
        if header_count >= 2:
            logger.info(f"Skipping multi-header row: {row[:3]}")
            return False
        
        # Special case: Check for rows where first cell is empty but second cell contains headers
        # This handles cases like row 69 where the header is in the second column
        if len(row) > 1 and not first_cell and row[1]:
            second_cell = str(row[1]).strip().lower()
            for indicator in header_indicators:
                if indicator in second_cell:
                    logger.info(f"Skipping header row (second cell): {row[1]}")
                    return False
        
        # Check for rows with None/NaN values in multiple columns (typical of header rows)
        # But be more careful - only skip if it's clearly a header pattern
        none_count = sum(1 for cell in row if cell is None or str(cell).strip() in ['', 'None', 'NaN', 'nan'])
        if none_count >= 3 and header_count >= 1:  # Only skip if it has both None/NaN AND header indicators
            logger.info(f"Skipping row with multiple None/NaN values and header indicators: {row}")
            return False
        
        # Skip completely empty rows
        if not first_cell:
            return False
        
        # Accept numeric IDs (like 16610898 from the image)
        if first_cell.isdigit() and len(first_cell) >= 6:
            return True
        
        # Accept alphanumeric IDs for payments
        if first_cell.replace(' ', '').replace('-', '').replace('.', '').isalnum() and len(first_cell) >= 6:
            return True
        
        # Special case: If the row has a valid contributor ID pattern and contributor name, it's likely valid data
        if len(row) >= 2 and first_cell.isdigit() and len(first_cell) >= 6:
            second_cell = str(row[1]).strip() if len(row) > 1 and row[1] else ''
            # If second cell looks like a name (contains letters and spaces), it's likely valid data
            if second_cell and any(c.isalpha() for c in second_cell) and ' ' in second_cell:
                return True
        
        # If we can't determine, err on the side of including it
        return True
    
    def extract_post1995_payments_from_pdf(self, pdf_path: str) -> pd.DataFrame:
        """
        Extract the post-1995 payments table from PDF.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            pandas DataFrame with cleaned data
        """
        logger.info(f"Extracting post-1995 payments from PDF: {pdf_path}")
        
        all_rows = []
        table_found = False
        headers_processed = False
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                logger.info(f"PDF has {len(pdf.pages)} pages")
                
                for page_num, page in enumerate(pdf.pages):
                    # Look for tables on this page
                    tables = page.extract_tables()
                    
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        
                        # Check if this table matches our expected format
                        headers = [cell.strip() if cell else '' for cell in table[0]]
                        
                        # Check if this looks like a payments table
                        is_payments_table = False
                        
                        # Method 1: Check for numbered header patterns (most reliable)
                        if len(headers) == 13:
                            numbered_matches = 0
                            for i, header in enumerate(headers):
                                if i < len(self.header_patterns):
                                    expected_pattern = self.header_patterns[i].lower()
                                    header_lower = header.lower()
                                    # Check if header contains the numbered pattern
                                    if any(part in header_lower for part in expected_pattern.split() if part not in ['[34]', '[35]', '[36]', '[37]', '[38]', '[39]', '[40]', '[41]', '[42]', '[43]', '[44]', '[45]', '[46]']):
                                        numbered_matches += 1
                            
                            if numbered_matches >= 8:  # Most headers should match
                                is_payments_table = True
                                logger.info(f"Found payments table (numbered headers, {numbered_matches} matches) on page {page_num + 1}")
                        
                        # Method 2: Check for exact column count and header matches
                        if not is_payments_table and len(headers) == 13:
                            matches = 0
                            for expected in self.expected_headers:
                                for header in headers:
                                    if expected.lower() in header.lower() or header.lower() in expected.lower():
                                        matches += 1
                                        break
                            if matches >= 8:  # Higher threshold for exact matches
                                is_payments_table = True
                                logger.info(f"Found payments table (13 cols, {matches} matches) on page {page_num + 1}")
                        
                        # Method 3: Check for payments keywords regardless of column count
                        if not is_payments_table:
                            header_text = ' '.join(headers).lower()
                            payments_keywords = ['pago', 'ibc', 'días', 'período', 'fecha', 'detalle', 'posterior', 'efectuados']
                            found_keywords = [kw for kw in payments_keywords if kw in header_text]
                            
                            if found_keywords and len(found_keywords) >= 4:  # Higher threshold
                                is_payments_table = True
                                logger.info(f"Found payments table (keywords: {found_keywords}) on page {page_num + 1}")
                        
                        # Method 4: Check for specific payments table title
                        if not is_payments_table:
                            header_text = ' '.join(headers).lower()
                            if any(title in header_text for title in ['detalle de pagos', 'pagos efectuados', 'posterior a 1995']):
                                is_payments_table = True
                                logger.info(f"Found payments table (title match) on page {page_num + 1}")
                        
                        if is_payments_table:
                            table_found = True
                            
                            # Process rows
                            start_row = 1 if not headers_processed else 0  # Skip headers on subsequent pages
                            headers_processed = True
                            
                            for row in table[start_row:]:
                                if row and any(cell for cell in row):  # Skip empty rows
                                    # Check if this row looks like a data row
                                    if self.is_data_row(row):
                                        cleaned_row = self.clean_row_data(row)
                                        if cleaned_row:
                                            all_rows.append(cleaned_row)
                                    else:
                                        logger.info(f"Skipping non-data row: {row[:3]}...")
                            
                            break  # Found our table, move to next page
        
        except Exception as e:
            logger.error(f"Error processing PDF: {str(e)}")
            raise
        
        if not all_rows:
            logger.warning("No data rows found in PDF")
            return pd.DataFrame(columns=self.output_columns)
        
        # Create DataFrame
        df = pd.DataFrame(all_rows)
        
        # Ensure all expected columns are present
        for col in self.output_columns:
            if col not in df.columns:
                df[col] = None
        
        # Reorder columns
        df = df[self.output_columns]
        
        logger.info(f"Successfully extracted {len(df)} rows from post-1995 payments table")
        return df
    
    def save_to_excel(self, df: pd.DataFrame, output_path: str):
        """
        Save DataFrame to Excel file.
        """
        try:
            df.to_excel(output_path, index=False, engine='openpyxl')
            logger.info(f"Post-1995 payments data saved to Excel: {output_path}")
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")
            raise
    
    def get_missing_periods_json(self, payments: pd.DataFrame) -> dict:
        """
        Given a payments DataFrame with a 'cont_period' column (YYYY-MM or YYYYMM),
        returns a JSON-serializable dict with missing periods between the first and last period.
        """
        import pandas as pd

        # Defensive copy to avoid modifying original DataFrame
        payments = payments.copy()

        # Convert 'cont_period' to datetime
        if payments['cont_period'].astype(str).str.contains('-').any():
            payments['cont_period_dt'] = pd.to_datetime(payments['cont_period'], format='%Y-%m')
        else:
            payments['cont_period_dt'] = pd.to_datetime(payments['cont_period'], format='%Y%m')

        # Sort by period
        df_sorted = payments.sort_values('cont_period_dt')

        # Get the first and last period
        start_period = df_sorted['cont_period_dt'].iloc[0]
        end_period = df_sorted['cont_period_dt'].iloc[-1]

        # Generate all months between start and end
        all_periods = pd.date_range(start=start_period, end=end_period, freq='MS')

        # Find which periods are missing from the DataFrame
        existing_periods = set(df_sorted['cont_period_dt'])
        missing_periods = [p for p in all_periods if p not in existing_periods]

        # Prepare output as JSON-serializable dict
        result = {
            "missing_periods": [p.strftime('%Y-%m') for p in missing_periods],
            "start_period": start_period.strftime('%Y-%m'),
            "end_period": end_period.strftime('%Y-%m'),
            "n_missing": len(missing_periods)
        }
        return result

class ColpensionesUnifiedExtractor:
    """
    Unified extractor that combines both table extractors to get all data in one call.
    Returns: (weeks_df, summary_values, payments_df)
    """
    
    def __init__(self):
        self.weeks_extractor = ColpensionesPDFExtractor()
        self.payments_extractor = ColpensionesPost1995PaymentsExtractor()
    
    def extract_all_from_pdf(self, pdf_path: str) -> tuple[pd.DataFrame, Dict[str, Optional[float]], pd.DataFrame]:
        """
        Extract all data from a Colpensiones PDF using separate extractors.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            Tuple of (weeks_df, summary_values, payments_df):
            - weeks_df: DataFrame with contribution weeks data
            - summary_values: Dictionary with total weeks summary
            - payments_df: DataFrame with post-1995 payments data
        """
        logger.info(f"Extracting all data from PDF: {pdf_path}")
        
        # Extract weeks data and summary values using the weeks extractor
        logger.info("Extracting weeks data and summary values...")
        weeks_df, summary_values = self.weeks_extractor.extract_table_and_summary_from_pdf(pdf_path)
        
        # Extract post-1995 payments data using the payments extractor
        logger.info("Extracting post-1995 payments data...")
        payments_df = self.payments_extractor.extract_post1995_payments_from_pdf(pdf_path)
        
        logger.info(f"Unified extraction complete:")
        logger.info(f"  - Weeks data: {len(weeks_df)} rows")
        logger.info(f"  - Summary values: {summary_values}")
        logger.info(f"  - Payments data: {len(payments_df)} rows")
        
        return weeks_df, summary_values, payments_df
    
    def save_all_to_excel(self, weeks_df: pd.DataFrame, summary_values: Dict[str, Optional[float]], 
                         payments_df: pd.DataFrame, output_path: str):
        """
        Save all extracted data to an Excel file with multiple sheets.
        
        Args:
            weeks_df: DataFrame with contribution weeks data
            summary_values: Dictionary with total weeks summary
            payments_df: DataFrame with post-1995 payments data
            output_path: Path for the output Excel file
        """
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Save weeks data
                weeks_df.to_excel(writer, sheet_name='Weeks_Data', index=False)
                
                # Save summary values
                summary_df = pd.DataFrame([summary_values])
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Save payments data
                payments_df.to_excel(writer, sheet_name='Payments_Data', index=False)
            
            logger.info(f"All data saved to Excel: {output_path}")
            logger.info(f"  - Weeks_Data sheet: {len(weeks_df)} rows")
            logger.info(f"  - Summary sheet: 1 row")
            logger.info(f"  - Payments_Data sheet: {len(payments_df)} rows")
            
        except Exception as e:
            logger.error(f"Error saving to Excel: {str(e)}")
            raise

if __name__ == "__main__":
    main()