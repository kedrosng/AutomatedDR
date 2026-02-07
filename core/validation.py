"""
Validation module for JV Cost Control System
Enforces data quality rules and Australian compliance checks
"""

import pandas as pd
import pandera as pa
from pandera import Column, DataFrameSchema, Check
from typing import List, Dict, Any
import logging
import yaml
import os

logger = logging.getLogger(__name__)

class DataValidator:
    def __init__(self, config_path: str = "config/rules.yaml"):
        self.config_path = config_path
        self.rules = self.load_validation_rules()
    
    def load_validation_rules(self) -> Dict[str, Any]:
        """Load validation rules from YAML configuration"""
        if os.path.exists(self.config_path):
            with open(self.config_path, 'r') as f:
                return yaml.safe_load(f)
        else:
            # Default rules if config not found
            return {
                "hr_master": {
                    "mandatory_fields": ["staff_id", "name", "role", "rate"],
                    "rate_bands": {
                        "General Builder": {"min": 800, "max": 1500},
                        "Steel Fixer": {"min": 1000, "max": 1800},
                        "Concreter": {"min": 900, "max": 1700},
                        "Safety Officer": {"min": 1200, "max": 2200},
                        "Crane Operator": {"min": 1500, "max": 2500}
                    }
                },
                "tfr_timesheet": {
                    "mandatory_fields": ["staff_id", "date", "hours"],
                    "max_hours_per_day": 12
                },
                "site_attendance": {
                    "mandatory_fields": ["staff_id", "date", "status"],
                    "valid_statuses": ["present", "absent", "leave", "public_holiday"]
                },
                "au_compliance": {
                    "foreign_worker_requires_visa": True,
                    "visa_pattern": r"^\d{3}[A-Z]$",
                    "superannuation_contribution_rate": 0.11,
                    "superannuation_max_salary_base": 275000
                }
            }
    
    def validate_hr_master(self, df: pd.DataFrame) -> pd.Series:
        """Validate HR master data against business rules"""
        errors = []
        
        # Check mandatory fields
        mandatory_fields = self.rules.get("hr_master", {}).get("mandatory_fields", [])
        for field in mandatory_fields:
            if field not in df.columns:
                errors.append(f"Mandatory field '{field}' missing from HR master data")
                continue
            
            missing_values = df[field].isna().sum()
            if missing_values > 0:
                errors.append(f"HR master has {missing_values} missing values in '{field}' field")
        
        # Validate rate bands
        if "role" in df.columns and "rate" in df.columns:
            rate_bands = self.rules.get("hr_master", {}).get("rate_bands", {})
            invalid_rates = []
            
            for _, row in df.iterrows():
                role = row.get("role", "")
                rate = row.get("rate", 0)
                
                if pd.notna(role) and role in rate_bands:
                    band = rate_bands[role]
                    if pd.isna(rate) or rate < band["min"] or rate > band["max"]:
                        invalid_rates.append({
                            "staff_id": row.get("staff_id", "Unknown"),
                            "role": role,
                            "rate": rate,
                            "expected_range": f"{band['min']}-{band['max']}"
                        })
            
            if invalid_rates:
                errors.append(f"Found {len(invalid_rates)} rate violations")
                logger.warning(f"Rate violations: {invalid_rates}")
        
        # Check for duplicate staff IDs
        if "staff_id" in df.columns:
            duplicates = df[df.duplicated(subset=["staff_id"], keep=False)]
            if not duplicates.empty:
                errors.append(f"Found {len(duplicates)} duplicate staff IDs")
        
        return pd.Series({"errors": errors, "is_valid": len(errors) == 0})
    
    def validate_tfr_timesheet(self, df: pd.DataFrame) -> pd.Series:
        """Validate TFR timesheet data"""
        errors = []
        
        # Check mandatory fields
        mandatory_fields = self.rules.get("tfr_timesheet", {}).get("mandatory_fields", [])
        for field in mandatory_fields:
            if field not in df.columns:
                errors.append(f"Mandatory field '{field}' missing from TFR timesheet")
                continue
            
            missing_values = df[field].isna().sum()
            if missing_values > 0:
                errors.append(f"TFR timesheet has {missing_values} missing values in '{field}' field")
        
        # Validate hours worked
        if "hours" in df.columns:
            max_hours = self.rules.get("tfr_timesheet", {}).get("max_hours_per_day", 12)
            invalid_hours = df[df["hours"] > max_hours]
            if not invalid_hours.empty:
                errors.append(f"Found {len(invalid_hours)} records with hours > {max_hours} per day")
        
        # Validate date format
        if "date" in df.columns:
            invalid_dates = []
            for idx, date_val in df["date"].items():
                try:
                    pd.to_datetime(date_val)
                except:
                    invalid_dates.append(idx)
            
            if invalid_dates:
                errors.append(f"Found {len(invalid_dates)} invalid dates")
        
        return pd.Series({"errors": errors, "is_valid": len(errors) == 0})
    
    def validate_site_attendance(self, df: pd.DataFrame) -> pd.Series:
        """Validate site attendance data"""
        errors = []
        
        # Check mandatory fields
        mandatory_fields = self.rules.get("site_attendance", {}).get("mandatory_fields", [])
        for field in mandatory_fields:
            if field not in df.columns:
                errors.append(f"Mandatory field '{field}' missing from site attendance")
                continue
            
            missing_values = df[field].isna().sum()
            if missing_values > 0:
                errors.append(f"Site attendance has {missing_values} missing values in '{field}' field")
        
        # Validate status values
        if "status" in df.columns:
            valid_statuses = self.rules.get("site_attendance", {}).get("valid_statuses", [])
            invalid_statuses = df[~df["status"].isin(valid_statuses)]
            if not invalid_statuses.empty:
                errors.append(f"Found {len(invalid_statuses)} records with invalid status values")
        
        return pd.Series({"errors": errors, "is_valid": len(errors) == 0})
    
    def validate_au_compliance(self, df: pd.DataFrame) -> pd.Series:
        """Validate Australian compliance requirements"""
        errors = []
        
        au_rules = self.rules.get("au_compliance", {})
        
        # Check for missing visas for foreign workers
        if (au_rules.get("foreign_worker_requires_visa") and 
            "nationality" in df.columns and 
            "visa_permit" in df.columns):
            
            # Assuming foreign workers are those not from Australia or New Zealand
            foreign_workers = df[
                (~df["nationality"].str.contains("au|aussie|australia|nz|kiwi|new zealand", case=False, na=False))
            ]
            
            missing_visas = foreign_workers[foreign_workers["visa_permit"].isna() | 
                                         (foreign_workers["visa_permit"] == "") |
                                         (foreign_workers["visa_permit"] == "None")]
            
            if not missing_visas.empty:
                errors.append(f"AU Compliance: Found {len(missing_visas)} foreign workers with missing visas")
                logger.warning(f"Missing visas for: {list(missing_visas['staff_id'])}")
        
        # Validate visa format
        if "visa_permit" in df.columns and au_rules.get("visa_pattern"):
            import re
            pattern = au_rules["visa_pattern"]
            invalid_visas = df[
                (df["visa_permit"].notna()) & 
                (~df["visa_permit"].apply(lambda x: bool(re.match(pattern, str(x))) if pd.notna(x) else False))
            ]
            
            if not invalid_visas.empty:
                errors.append(f"Found {len(invalid_visas)} visas with invalid format")
        
        # Check for award rate compliance
        if au_rules.get("award_rate_check") and "role" in df.columns and "rate" in df.columns:
            # Basic check for extremely low rates that might violate award conditions
            low_rate_threshold = 20  # AUD minimum for many construction awards
            low_rates = df[(df["rate"] < low_rate_threshold) & (df["rate"] > 0)]
            if not low_rates.empty:
                errors.append(f"Award violation: Found {len(low_rates)} workers paid below minimum award rate (${low_rate_threshold}/hr)")
        
        return pd.Series({"errors": errors, "is_valid": len(errors) == 0})
    
    def validate_data(self, df: pd.DataFrame, source_type: str) -> Dict[str, Any]:
        """Main validation method that applies all relevant checks"""
        logger.info(f"Validating {source_type} data with {len(df)} rows")
        
        results = {
            "source_type": source_type,
            "total_records": len(df),
            "validation_results": {},
            "overall_status": "pass",
            "summary_errors": []
        }
        
        # Apply source-specific validations
        if source_type == "hr_master":
            validation_result = self.validate_hr_master(df)
            results["validation_results"]["hr_master_validation"] = validation_result["errors"]
            if not validation_result["is_valid"]:
                results["overall_status"] = "fail"
                results["summary_errors"].extend(validation_result["errors"])
        
        elif source_type == "tfr_timesheet":
            validation_result = self.validate_tfr_timesheet(df)
            results["validation_results"]["tfr_timesheet_validation"] = validation_result["errors"]
            if not validation_result["is_valid"]:
                results["overall_status"] = "fail"
                results["summary_errors"].extend(validation_result["errors"])
        
        elif source_type == "site_attendance":
            validation_result = self.validate_site_attendance(df)
            results["validation_results"]["site_attendance_validation"] = validation_result["errors"]
            if not validation_result["is_valid"]:
                results["overall_status"] = "fail"
                results["summary_errors"].extend(validation_result["errors"])
        
        # Apply compliance validation regardless of source type
        compliance_result = self.validate_au_compliance(df)
        results["validation_results"]["compliance_validation"] = compliance_result["errors"]
        if not compliance_result["is_valid"]:
            results["overall_status"] = "fail"
            results["summary_errors"].extend(compliance_result["errors"])
        
        # Log validation summary
        if results["summary_errors"]:
            logger.warning(f"Validation failed for {source_type} with {len(results['summary_errors'])} errors")
        else:
            logger.info(f"Validation passed for {source_type}")
        
        return results

def validate_dataframe_schema(df: pd.DataFrame, expected_columns: List[str]) -> bool:
    """
    Quick schema validation to check if required columns exist
    """
    missing_cols = set(expected_columns) - set(df.columns)
    if missing_cols:
        logger.error(f"Missing required columns: {missing_cols}")
        return False
    return True

def run_comprehensive_validation(df: pd.DataFrame, source_type: str, partner: str) -> Dict[str, Any]:
    """
    Run comprehensive validation including schema, business rules, and compliance checks
    """
    validator = DataValidator()
    validation_results = validator.validate_data(df, source_type)
    
    # Additional cross-reference validations
    if source_type == "hr_master":
        # Check for consistency in rate information
        if "rate" in df.columns:
            rate_anomalies = detect_rate_anomalies(df)
            if rate_anomalies:
                validation_results["validation_results"]["rate_anomalies"] = rate_anomalies
                validation_results["overall_status"] = "warning"
    
    validation_results["partner"] = partner
    return validation_results

def detect_rate_anomalies(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Detect unusual rate patterns that might indicate errors
    """
    anomalies = []
    
    if "role" in df.columns and "rate" in df.columns:
        # Group by role and calculate statistics
        role_stats = df.groupby("role")["rate"].agg(["mean", "std", "count"]).reset_index()
        
        for _, row in df.iterrows():
            role = row["role"]
            rate = row["rate"]
            
            role_stat = role_stats[role_stats["role"] == role]
            if not role_stat.empty:
                mean_rate = role_stat.iloc[0]["mean"]
                std_rate = role_stat.iloc[0]["std"]
                
                # Flag rates that are more than 2 standard deviations from the mean
                if std_rate > 0 and abs(rate - mean_rate) > 2 * std_rate:
                    anomalies.append({
                        "staff_id": row.get("staff_id", "Unknown"),
                        "name": row.get("name", "Unknown"),
                        "role": role,
                        "rate": rate,
                        "mean_for_role": round(mean_rate, 2),
                        "deviation": round(abs(rate - mean_rate), 2)
                    })
    
    return anomalies