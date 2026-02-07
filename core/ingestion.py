"""
Data Ingestion Module for JV Cost Reconciliation

Handles loading and normalizing data from:
- CSV (HR master files)
- Excel (Site attendance logs)
- PDF (TFR timesheets with table extraction)

Provides automatic column normalization and schema drift detection.
"""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from datetime import datetime

import pandas as pd
import yaml
import pdfplumber

# Configure logging
logger = logging.getLogger(__name__)


@dataclass
class IngestionResult:
    """Result container for data ingestion operations."""
    success: bool
    data: Optional[pd.DataFrame]
    row_count: int
    column_count: int
    source_file: str
    partner: str
    source_type: str
    warnings: List[str]
    errors: List[str]
    schema_drift: List[str]


class DataIngestion:
    """
    Handles data ingestion from multiple file formats.
    
    Supports:
    - CSV files (HR master lists)
    - Excel files (Site attendance logs)
    - PDF files (TFR timesheets with table extraction)
    """
    
    def __init__(self, config_path: Path):
        """
        Initialize the data ingestion engine.
        
        Args:
            config_path: Path to the partners.yaml configuration file
        """
        self.config_path = Path(config_path)
        self._load_config()
        self._baseline_schemas: Dict[str, List[str]] = {}
        
    def _load_config(self) -> None:
        """Load partner configuration from YAML."""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f)
            logger.info(f"Loaded partner config from {self.config_path}")
        except Exception as e:
            logger.error(f"Failed to load config: {e}")
            raise
            
    def set_baseline_schema(self, source_type: str, columns: List[str]) -> None:
        """
        Set baseline schema for drift detection.
        
        Args:
            source_type: Type of source (hr_master, tfr, attendance)
            columns: List of expected normalized column names
        """
        self._baseline_schemas[source_type] = columns
        logger.info(f"Set baseline schema for {source_type}: {columns}")
    
    def _detect_schema_drift(
        self, 
        df: pd.DataFrame, 
        source_type: str
    ) -> List[str]:
        """
        Detect schema drift compared to baseline.
        
        Args:
            df: Loaded DataFrame
            source_type: Type of source for baseline comparison
            
        Returns:
            List of drift warnings
        """
        drift_warnings = []
        
        if source_type not in self._baseline_schemas:
            return drift_warnings
            
        baseline = set(self._baseline_schemas[source_type])
        current = set(df.columns.tolist())
        
        # New columns not in baseline
        new_cols = current - baseline
        if new_cols:
            drift_warnings.append(
                f"New columns detected: {', '.join(new_cols)}"
            )
            
        # Missing columns from baseline
        missing_cols = baseline - current
        if missing_cols:
            drift_warnings.append(
                f"Missing columns from baseline: {', '.join(missing_cols)}"
            )
            
        return drift_warnings
    
    def _normalize_columns(
        self, 
        df: pd.DataFrame, 
        partner: str, 
        source_type: str
    ) -> pd.DataFrame:
        """
        Normalize column names according to partner mapping.
        
        Args:
            df: Input DataFrame
            partner: Partner identifier (dragages, gammon)
            source_type: Type of source (hr_master, tfr, attendance)
            
        Returns:
            DataFrame with normalized column names
        """
        try:
            mappings = self.config['partners'][partner]['column_mappings'][source_type]
            
            # Create rename dict for columns that exist
            rename_dict = {}
            for orig, norm in mappings.items():
                if orig in df.columns:
                    rename_dict[orig] = norm
                    
            df = df.rename(columns=rename_dict)
            logger.debug(f"Normalized columns: {rename_dict}")
            
        except KeyError as e:
            logger.warning(f"No column mapping found for {partner}/{source_type}: {e}")
            
        return df
    
    def _normalize_role(self, role: str) -> str:
        """
        Normalize role names across partners.
        
        Args:
            role: Original role name
            
        Returns:
            Normalized role name
        """
        role_mappings = self.config.get('role_mappings', {})
        return role_mappings.get(role, role.lower().replace(' ', '_'))
    
    def load_csv(
        self, 
        file_path: Path, 
        partner: str,
        source_type: str = 'hr_master'
    ) -> IngestionResult:
        """
        Load and normalize a CSV file.
        
        Args:
            file_path: Path to the CSV file
            partner: Partner identifier
            source_type: Type of source
            
        Returns:
            IngestionResult with loaded data and metadata
        """
        file_path = Path(file_path)
        warnings = []
        errors = []
        
        try:
            # Try different encodings
            encodings = ['utf-8', 'utf-8-sig', 'cp1252', 'latin1']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    logger.info(f"Loaded CSV with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
                    
            if df is None:
                raise ValueError("Could not decode CSV with any known encoding")
                
            # Normalize columns
            df = self._normalize_columns(df, partner, source_type)
            
            # Normalize role names if role column exists
            if 'role' in df.columns:
                df['role_normalized'] = df['role'].apply(self._normalize_role)
                
            # Detect schema drift
            schema_drift = self._detect_schema_drift(df, source_type)
            if schema_drift:
                warnings.extend(schema_drift)
                
            # Add metadata columns
            df['_source_file'] = file_path.name
            df['_partner'] = partner
            df['_ingestion_time'] = datetime.now().isoformat()
            
            return IngestionResult(
                success=True,
                data=df,
                row_count=len(df),
                column_count=len(df.columns),
                source_file=str(file_path),
                partner=partner,
                source_type=source_type,
                warnings=warnings,
                errors=errors,
                schema_drift=schema_drift
            )
            
        except Exception as e:
            logger.error(f"Failed to load CSV {file_path}: {e}")
            errors.append(str(e))
            return IngestionResult(
                success=False,
                data=None,
                row_count=0,
                column_count=0,
                source_file=str(file_path),
                partner=partner,
                source_type=source_type,
                warnings=warnings,
                errors=errors,
                schema_drift=[]
            )
    
    def load_excel(
        self, 
        file_path: Path, 
        partner: str,
        source_type: str = 'attendance',
        sheet_name: Optional[str] = None
    ) -> IngestionResult:
        """
        Load and normalize an Excel file.
        
        Args:
            file_path: Path to the Excel file
            partner: Partner identifier
            source_type: Type of source
            sheet_name: Specific sheet to load (None = first sheet)
            
        Returns:
            IngestionResult with loaded data and metadata
        """
        file_path = Path(file_path)
        warnings = []
        errors = []
        
        try:
            # Load Excel file
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
                
            logger.info(f"Loaded Excel file with {len(df)} rows")
            
            # Handle merged cells (forward fill)
            df = df.ffill()
            
            # Clean column names (remove whitespace, newlines)
            df.columns = df.columns.str.strip().str.replace('\n', ' ')
            
            # Normalize columns
            df = self._normalize_columns(df, partner, source_type)
            
            # Detect schema drift
            schema_drift = self._detect_schema_drift(df, source_type)
            if schema_drift:
                warnings.extend(schema_drift)
                
            # Parse date columns
            date_columns = ['date', 'attendance_date', 'Date', 'Attendance_Date']
            for col in date_columns:
                if col in df.columns:
                    try:
                        df[col] = pd.to_datetime(df[col], errors='coerce')
                    except Exception as e:
                        warnings.append(f"Could not parse dates in {col}: {e}")
                        
            # Add metadata columns
            df['_source_file'] = file_path.name
            df['_partner'] = partner
            df['_ingestion_time'] = datetime.now().isoformat()
            
            return IngestionResult(
                success=True,
                data=df,
                row_count=len(df),
                column_count=len(df.columns),
                source_file=str(file_path),
                partner=partner,
                source_type=source_type,
                warnings=warnings,
                errors=errors,
                schema_drift=schema_drift
            )
            
        except Exception as e:
            logger.error(f"Failed to load Excel {file_path}: {e}")
            errors.append(str(e))
            return IngestionResult(
                success=False,
                data=None,
                row_count=0,
                column_count=0,
                source_file=str(file_path),
                partner=partner,
                source_type=source_type,
                warnings=warnings,
                errors=errors,
                schema_drift=[]
            )
    
    def load_pdf(
        self, 
        file_path: Path, 
        partner: str,
        source_type: str = 'tfr'
    ) -> IngestionResult:
        """
        Extract tables from a PDF file.
        
        Uses pdfplumber for table detection with fallback heuristics.
        
        Args:
            file_path: Path to the PDF file
            partner: Partner identifier
            source_type: Type of source
            
        Returns:
            IngestionResult with extracted table data
        """
        file_path = Path(file_path)
        warnings = []
        errors = []
        all_rows = []
        headers = None
        
        try:
            with pdfplumber.open(file_path) as pdf:
                logger.info(f"Processing PDF with {len(pdf.pages)} pages")
                
                for page_num, page in enumerate(pdf.pages, 1):
                    # Try to extract tables
                    tables = page.extract_tables()
                    
                    if not tables:
                        # Fallback: try with different settings
                        tables = page.extract_tables(
                            table_settings={
                                "vertical_strategy": "lines",
                                "horizontal_strategy": "lines",
                                "snap_tolerance": 5,
                            }
                        )
                        
                    if not tables:
                        # Last resort: text extraction
                        text = page.extract_text()
                        if text:
                            warnings.append(
                                f"Page {page_num}: No table detected, using text extraction"
                            )
                            # Try to parse structured text
                            lines = text.split('\n')
                            for line in lines:
                                parts = line.split()
                                if len(parts) >= 3:  # Minimum columns expected
                                    all_rows.append(parts)
                        continue
                        
                    for table in tables:
                        if not table:
                            continue
                            
                        for row_idx, row in enumerate(table):
                            # Clean row data
                            cleaned_row = [
                                str(cell).strip() if cell else '' 
                                for cell in row
                            ]
                            
                            # Skip empty rows
                            if not any(cleaned_row):
                                continue
                                
                            # First non-empty row with text is likely header
                            if headers is None and any(cleaned_row):
                                # Check if it looks like a header row
                                if any(
                                    any(kw in str(c).lower() for kw in 
                                        ['id', 'name', 'date', 'hours', 'rate'])
                                    for c in cleaned_row
                                ):
                                    headers = cleaned_row
                                    continue
                                    
                            # Skip repeated header rows
                            if headers and cleaned_row == headers:
                                continue
                                
                            all_rows.append(cleaned_row)
            
            if not all_rows:
                raise ValueError("No table data extracted from PDF")
                
            # Create DataFrame
            if headers:
                # Ensure headers and data have same length
                max_cols = max(len(headers), max(len(r) for r in all_rows))
                headers = headers + [''] * (max_cols - len(headers))
                all_rows = [r + [''] * (max_cols - len(r)) for r in all_rows]
                df = pd.DataFrame(all_rows, columns=headers)
            else:
                df = pd.DataFrame(all_rows)
                # Try to infer headers from first row
                if len(df) > 0:
                    first_row = df.iloc[0]
                    if any(
                        any(kw in str(c).lower() for kw in 
                            ['id', 'name', 'date', 'hours', 'rate'])
                        for c in first_row
                    ):
                        df.columns = df.iloc[0]
                        df = df.iloc[1:].reset_index(drop=True)
                warnings.append("Headers inferred from data - please verify")
                
            # Clean up empty columns
            df = df.loc[:, ~(df.columns == '')]
            df = df.dropna(axis=1, how='all')
            
            # Normalize columns
            df = self._normalize_columns(df, partner, source_type)
            
            # Convert numeric columns
            numeric_cols = ['hours', 'ot_hours', 'daily_rate']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    
            # Detect schema drift
            schema_drift = self._detect_schema_drift(df, source_type)
            if schema_drift:
                warnings.extend(schema_drift)
                
            # Add metadata columns
            df['_source_file'] = file_path.name
            df['_partner'] = partner
            df['_ingestion_time'] = datetime.now().isoformat()
            
            return IngestionResult(
                success=True,
                data=df,
                row_count=len(df),
                column_count=len(df.columns),
                source_file=str(file_path),
                partner=partner,
                source_type=source_type,
                warnings=warnings,
                errors=errors,
                schema_drift=schema_drift
            )
            
        except Exception as e:
            logger.error(f"Failed to load PDF {file_path}: {e}")
            errors.append(str(e))
            return IngestionResult(
                success=False,
                data=None,
                row_count=0,
                column_count=0,
                source_file=str(file_path),
                partner=partner,
                source_type=source_type,
                warnings=warnings,
                errors=errors,
                schema_drift=[]
            )
    
    def load_file(
        self, 
        file_path: Path, 
        partner: str,
        source_type: Optional[str] = None
    ) -> IngestionResult:
        """
        Auto-detect file type and load accordingly.
        
        Args:
            file_path: Path to the file
            partner: Partner identifier
            source_type: Type of source (auto-detected if None)
            
        Returns:
            IngestionResult with loaded data
        """
        file_path = Path(file_path)
        suffix = file_path.suffix.lower()
        
        # Auto-detect source type from filename if not provided
        if source_type is None:
            if 'hr' in file_path.stem.lower():
                source_type = 'hr_master'
            elif 'tfr' in file_path.stem.lower():
                source_type = 'tfr'
            elif 'attendance' in file_path.stem.lower():
                source_type = 'attendance'
            else:
                source_type = 'unknown'
                
        if suffix == '.csv':
            return self.load_csv(file_path, partner, source_type)
        elif suffix in ['.xlsx', '.xls']:
            return self.load_excel(file_path, partner, source_type)
        elif suffix == '.pdf':
            return self.load_pdf(file_path, partner, source_type)
        else:
            return IngestionResult(
                success=False,
                data=None,
                row_count=0,
                column_count=0,
                source_file=str(file_path),
                partner=partner,
                source_type=source_type,
                warnings=[],
                errors=[f"Unsupported file type: {suffix}"],
                schema_drift=[]
            )
