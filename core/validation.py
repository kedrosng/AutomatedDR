"""
Data Validation Module for JV Cost Reconciliation (Australia)

Implements schema validation and business rule checks:
- Mandatory field validation per source type
- Australian compliance checks (Visa/Work Rights, Superannuation)
- Rate band validation per role
"""

import logging
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field
from enum import Enum

import pandas as pd
import pandera as pa
from pandera import Column, Check, DataFrameSchema
import yaml

# Configure logging
logger = logging.getLogger(__name__)


class ValidationSeverity(Enum):
    """Severity levels for validation issues."""
    CRITICAL = "critical"
    MAJOR = "major"
    MINOR = "minor"
    INFO = "info"


@dataclass
class ValidationIssue:
    """Container for a single validation issue."""
    severity: ValidationSeverity
    category: str
    message: str
    row_index: Optional[int] = None
    column: Optional[str] = None
    staff_id: Optional[str] = None
    partner: Optional[str] = None
    value: Optional[Any] = None


@dataclass
class ValidationResult:
    """Result container for validation operations."""
    is_valid: bool
    issues: List[ValidationIssue] = field(default_factory=list)
    row_count: int = 0
    valid_row_count: int = 0
    issue_summary: Dict[str, int] = field(default_factory=dict)
    
    def add_issue(self, issue: ValidationIssue) -> None:
        """Add a validation issue and update summary."""
        self.issues.append(issue)
        key = f"{issue.severity.value}_{issue.category}"
        self.issue_summary[key] = self.issue_summary.get(key, 0) + 1
        if issue.severity in [ValidationSeverity.CRITICAL, ValidationSeverity.MAJOR]:
            self.is_valid = False


class DataValidator:
    """
    Validates data against schemas and business rules.
    
    Supports:
    - Pandera schema validation
    - Mandatory field checks
    - Australian compliance (Visa, Super)
    - Rate band validation
    """
    
    def __init__(self, rules_path: Path):
        """
        Initialize the validator.
        
        Args:
            rules_path: Path to rules.yaml configuration
        """
        self.rules_path = Path(rules_path)
        self._load_rules()
        self._build_schemas()
        
    def _load_rules(self) -> None:
        """Load validation rules from YAML."""
        try:
            with open(self.rules_path, 'r', encoding='utf-8') as f:
                self.rules = yaml.safe_load(f)
            logger.info(f"Loaded validation rules from {self.rules_path}")
        except Exception as e:
            logger.error(f"Failed to load rules: {e}")
            raise
            
    def _build_schemas(self) -> None:
        """Build Pandera schemas for each source type."""
        
        # HR Master schema
        self.hr_schema = DataFrameSchema({
            "staff_id": Column(
                str, 
                Check(lambda x: x.str.len() > 0, error="Staff ID cannot be empty"),
                nullable=False
            ),
            "name": Column(
                str,
                Check(lambda x: x.str.len() >= 2, error="Name must be at least 2 characters"),
                nullable=False
            ),
            "role": Column(str, nullable=True),
            "daily_rate": Column(
                float,
                Check(lambda x: x >= 0, error="Daily rate must be non-negative"),
                nullable=True,
                coerce=True
            ),
            "nationality": Column(str, nullable=True),
            "visa_status": Column(str, nullable=True),
            "super_fund": Column(str, nullable=True),
        }, strict=False, coerce=True)
        
        # TFR (Timesheet) schema
        self.tfr_schema = DataFrameSchema({
            "staff_id": Column(
                str,
                Check(lambda x: x.str.len() > 0, error="Staff ID cannot be empty"),
                nullable=False
            ),
            "name": Column(str, nullable=True),
            "date": Column(nullable=True),
            "hours": Column(
                float,
                Check(lambda x: (x >= 0) & (x <= 24), error="Hours must be 0-24"),
                nullable=True,
                coerce=True
            ),
            "daily_rate": Column(
                float,
                Check(lambda x: x >= 0, error="Rate must be non-negative"),
                nullable=True,
                coerce=True
            ),
        }, strict=False, coerce=True)
        
        # Attendance schema
        self.attendance_schema = DataFrameSchema({
            "staff_id": Column(
                str,
                Check(lambda x: x.str.len() > 0, error="Staff ID cannot be empty"),
                nullable=False
            ),
            "name": Column(str, nullable=True),
            "date": Column(nullable=True),
            "clock_in": Column(nullable=True),
            "clock_out": Column(nullable=True),
        }, strict=False, coerce=True)
        
    def _get_schema(self, source_type: str) -> Optional[DataFrameSchema]:
        """Get the appropriate schema for a source type."""
        schemas = {
            'hr_master': self.hr_schema,
            'tfr': self.tfr_schema,
            'attendance': self.attendance_schema
        }
        return schemas.get(source_type)
        
    def validate_schema(
        self, 
        df: pd.DataFrame, 
        source_type: str
    ) -> ValidationResult:
        """Validate DataFrame against Pandera schema."""
        result = ValidationResult(is_valid=True, row_count=len(df))
        schema = self._get_schema(source_type)
        
        if schema is None:
            result.add_issue(ValidationIssue(
                severity=ValidationSeverity.INFO,
                category="schema",
                message=f"No schema defined for source type: {source_type}"
            ))
            result.valid_row_count = len(df)
            return result
            
        try:
            # Validate with Pandera
            schema.validate(df, lazy=True)
            result.valid_row_count = len(df)
            
        except pa.errors.SchemaErrors as e:
            # Collect all schema failures
            for _, row in e.failure_cases.iterrows():
                result.add_issue(ValidationIssue(
                    severity=ValidationSeverity.MAJOR,
                    category="schema",
                    message=str(row.get('check', 'Schema validation failed')),
                    row_index=row.get('index'),
                    column=row.get('column'),
                    value=row.get('failure_case')
                ))
            result.valid_row_count = len(df) - len(e.failure_cases)
            
        return result
        
    def validate_mandatory_fields(
        self, 
        df: pd.DataFrame, 
        source_type: str
    ) -> ValidationResult:
        """Check for mandatory fields based on source type."""
        result = ValidationResult(is_valid=True, row_count=len(df))
        
        mandatory_fields = {
            'hr_master': ['staff_id', 'name', 'role', 'daily_rate'],
            'tfr': ['staff_id', 'date', 'hours'],
            'attendance': ['staff_id', 'date']
        }
        
        required = mandatory_fields.get(source_type, [])
        
        for field in required:
            if field not in df.columns:
                result.add_issue(ValidationIssue(
                    severity=ValidationSeverity.CRITICAL,
                    category="mandatory_field",
                    message=f"Missing mandatory column: {field}"
                ))
            else:
                # Check for null/empty values
                null_count = df[field].isna().sum()
                if null_count > 0:
                    result.add_issue(ValidationIssue(
                        severity=ValidationSeverity.MAJOR,
                        category="mandatory_field",
                        message=f"Column '{field}' has {null_count} missing values"
                    ))
                    
        result.valid_row_count = len(df)
        return result
        
    def check_visa_status(
        self, 
        df: pd.DataFrame,
        partner: Optional[str] = None
    ) -> ValidationResult:
        """
        Check Visa / Work Rights compliance for foreign workers.
        
        Australian employers must verify work rights (VEVO).
        """
        result = ValidationResult(is_valid=True, row_count=len(df))
        
        # Get nationalities requiring permits
        nationalities_requiring = self.rules.get('visa_compliance', {}).get(
            'nationalities_requiring_permit', []
        )
        permit_pattern = self.rules.get('visa_compliance', {}).get(
            'permit_format', r'^[0-9]{9,13}$'
        )
        
        if 'nationality' not in df.columns:
            # Try to infer or fallback
            return result
            
        if 'visa_status' not in df.columns:
            # Might be named differently in source, but mapped in ingestion
            return result
            
        for idx, row in df.iterrows():
            nationality = str(row.get('nationality', '')).strip()
            visa = str(row.get('visa_status', '')).strip()
            staff_id = str(row.get('staff_id', ''))
            name = str(row.get('name', ''))
            
            # Check if this nationality requires a permit
            if any(nat.lower() in nationality.lower() for nat in nationalities_requiring):
                # Permit is required
                if not visa or visa.lower() in ['nan', 'none', '']:
                    result.add_issue(ValidationIssue(
                        severity=ValidationSeverity.CRITICAL,
                        category="missing_visa",
                        message=f"Foreign worker missing Visa details: {name} ({nationality})",
                        row_index=idx,
                        column="visa_status",
                        staff_id=staff_id,
                        partner=partner
                    ))
                    
        result.valid_row_count = len(df) - len([
            i for i in result.issues 
            if i.category == "missing_visa"
        ])
        return result
        
    def validate_rate_bands(
        self, 
        df: pd.DataFrame,
        partner: Optional[str] = None
    ) -> ValidationResult:
        """Validate daily rates against role-based bands."""
        result = ValidationResult(is_valid=True, row_count=len(df))
        
        rate_bands = self.rules.get('rate_bands', {})
        
        if 'daily_rate' not in df.columns:
            return result
            
        role_col = 'role_normalized' if 'role_normalized' in df.columns else 'role'
        
        for idx, row in df.iterrows():
            role = str(row.get(role_col, '')).lower().replace(' ', '_')
            rate = row.get('daily_rate')
            staff_id = str(row.get('staff_id', ''))
            name = str(row.get('name', ''))
            
            if pd.isna(rate) or role not in rate_bands:
                continue
                
            band = rate_bands[role]
            min_rate = band.get('min', 0)
            max_rate = band.get('max', float('inf'))
            
            if rate < min_rate:
                result.add_issue(ValidationIssue(
                    severity=ValidationSeverity.MAJOR,
                    category="rate_out_of_band",
                    message=f"Rate ${rate:.0f} below minimum ${min_rate} for {role}: {name}",
                    row_index=idx,
                    column="daily_rate",
                    staff_id=staff_id,
                    partner=partner,
                    value=rate
                ))
            elif rate > max_rate:
                result.add_issue(ValidationIssue(
                    severity=ValidationSeverity.MAJOR,
                    category="rate_out_of_band",
                    message=f"Rate ${rate:.0f} above maximum ${max_rate} for {role}: {name}",
                    row_index=idx,
                    column="daily_rate",
                    staff_id=staff_id,
                    partner=partner,
                    value=rate
                ))
                
        result.valid_row_count = len(df) - len(result.issues)
        return result
        
    def validate_super(
        self, 
        df: pd.DataFrame,
        partner: Optional[str] = None
    ) -> ValidationResult:
        """Validate Superannuation details."""
        result = ValidationResult(is_valid=True, row_count=len(df))
        
        if 'super_fund' not in df.columns:
            return result
            
        for idx, row in df.iterrows():
            super_fund = str(row.get('super_fund', '')).strip()
            staff_id = str(row.get('staff_id', ''))
            name = str(row.get('name', ''))
            
            if not super_fund or super_fund.lower() in ['nan', 'none', '']:
                result.add_issue(ValidationIssue(
                    severity=ValidationSeverity.MAJOR,
                    category="super_error",
                    message=f"Missing Superannuation fund details: {name}",
                    row_index=idx,
                    column="super_fund",
                    staff_id=staff_id,
                    partner=partner
                ))
                
        result.valid_row_count = len(df) - len(result.issues)
        return result
        
    def validate_all(
        self, 
        df: pd.DataFrame, 
        source_type: str,
        partner: Optional[str] = None
    ) -> ValidationResult:
        """Run all applicable validations."""
        combined = ValidationResult(is_valid=True, row_count=len(df))
        
        # Schema validation
        schema_result = self.validate_schema(df, source_type)
        combined.issues.extend(schema_result.issues)
        combined.issue_summary.update(schema_result.issue_summary)
        if not schema_result.is_valid:
            combined.is_valid = False
            
        # Mandatory fields
        mandatory_result = self.validate_mandatory_fields(df, source_type)
        combined.issues.extend(mandatory_result.issues)
        combined.issue_summary.update(mandatory_result.issue_summary)
        if not mandatory_result.is_valid:
            combined.is_valid = False
            
        # Source-specific validations
        if source_type == 'hr_master':
            # Visa compliance
            visa_result = self.check_visa_status(df, partner)
            combined.issues.extend(visa_result.issues)
            combined.issue_summary.update(visa_result.issue_summary)
            if not visa_result.is_valid:
                combined.is_valid = False
                
            # Rate bands
            rate_result = self.validate_rate_bands(df, partner)
            combined.issues.extend(rate_result.issues)
            combined.issue_summary.update(rate_result.issue_summary)
            if not rate_result.is_valid:
                combined.is_valid = False
                
            # Super
            super_result = self.validate_super(df, partner)
            combined.issues.extend(super_result.issues)
            combined.issue_summary.update(super_result.issue_summary)
            if not super_result.is_valid:
                combined.is_valid = False
                
        # Count valid rows
        critical_staff_ids = set(
            i.staff_id for i in combined.issues 
            if i.severity == ValidationSeverity.CRITICAL and i.staff_id
        )
        combined.valid_row_count = len(df) - len(critical_staff_ids)
        
        return combined
