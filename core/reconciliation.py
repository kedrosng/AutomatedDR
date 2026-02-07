"""
Reconciliation Engine for JV Cost Reconciliation

Implements 3-way matching between data sources:
- HR Master ↔ TFR Timesheet ↔ Site Attendance

Uses RapidFuzz for fuzzy name matching with confidence scoring.
"""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set, Any
from dataclasses import dataclass, field
from enum import Enum
from datetime import datetime
import hashlib

import pandas as pd
import numpy as np
from rapidfuzz import fuzz, process
import yaml

# Configure logging
logger = logging.getLogger(__name__)


class DiscrepancyType(Enum):
    """Types of discrepancies detected during reconciliation."""
    GHOST_WORKER = "ghost_worker"              # In TFR but not HR
    MISSING_TFR = "missing_tfr"                # In HR but no TFR
    MISSING_ATTENDANCE = "missing_attendance"  # In TFR but no attendance
    FUZZY_MATCH_LOW = "fuzzy_match_low"        # Name match < 85%
    FUZZY_MATCH_REVIEW = "fuzzy_match_review"  # Name match 85-95%
    RATE_DEVIATION = "rate_deviation"          # >15% rate diff
    RATE_DEVIATION_EXTREME = "rate_deviation_extreme"  # >50% rate diff
    DUPLICATE_ID = "duplicate_id"              # Same ID across partners
    HOURS_MISMATCH = "hours_mismatch"          # TFR vs attendance hours
    CROSS_PARTNER_RATE_DIFF = "cross_partner_rate_diff"  # Rate diff between partners


class DiscrepancySeverity(Enum):
    """Severity classification for discrepancies."""
    CRITICAL = "critical"
    MAJOR = "major"
    MINOR = "minor"


@dataclass
class Discrepancy:
    """Container for a single discrepancy."""
    id: str
    type: DiscrepancyType
    severity: DiscrepancySeverity
    partner: str
    staff_id: str
    staff_name: str
    description: str
    confidence: Optional[float] = None
    amount_hkd: Optional[float] = None
    source_a: Optional[str] = None
    source_b: Optional[str] = None
    value_a: Optional[Any] = None
    value_b: Optional[Any] = None
    status: str = "Open"
    created_at: datetime = field(default_factory=datetime.now)
    resolved_at: Optional[datetime] = None
    resolved_by: Optional[str] = None
    resolution_notes: Optional[str] = None


@dataclass
class ReconciliationResult:
    """Result container for reconciliation operations."""
    success: bool
    discrepancies: List[Discrepancy] = field(default_factory=list)
    matched_count: int = 0
    unmatched_count: int = 0
    summary: Dict[str, int] = field(default_factory=dict)
    reconciled_data: Optional[pd.DataFrame] = None
    
    def get_critical_count(self) -> int:
        """Count critical severity discrepancies."""
        return len([d for d in self.discrepancies if d.severity == DiscrepancySeverity.CRITICAL])
    
    def get_major_count(self) -> int:
        """Count major severity discrepancies."""
        return len([d for d in self.discrepancies if d.severity == DiscrepancySeverity.MAJOR])
    
    def get_minor_count(self) -> int:
        """Count minor severity discrepancies."""
        return len([d for d in self.discrepancies if d.severity == DiscrepancySeverity.MINOR])


class ReconciliationEngine:
    """
    3-way reconciliation engine with fuzzy matching.
    
    Matches records across:
    1. HR Master (source of truth for staff)
    2. TFR Timesheet (claimed hours/rates)
    3. Site Attendance (actual presence)
    """
    
    def __init__(self, rules_path: Path):
        """
        Initialize the reconciliation engine.
        
        Args:
            rules_path: Path to rules.yaml configuration
        """
        self.rules_path = Path(rules_path)
        self._load_rules()
        self._discrepancy_counter = 0
        
    def _load_rules(self) -> None:
        """Load reconciliation rules from YAML."""
        try:
            with open(self.rules_path, 'r', encoding='utf-8') as f:
                self.rules = yaml.safe_load(f)
            logger.info(f"Loaded reconciliation rules from {self.rules_path}")
        except Exception as e:
            logger.error(f"Failed to load rules: {e}")
            raise
            
    def _generate_discrepancy_id(self) -> str:
        """Generate unique discrepancy ID."""
        self._discrepancy_counter += 1
        timestamp = datetime.now().strftime("%Y%m%d")
        return f"DISC-{timestamp}-{self._discrepancy_counter:04d}"
    
    def _classify_severity(self, disc_type: DiscrepancyType) -> DiscrepancySeverity:
        """
        Classify discrepancy severity based on type.
        
        Args:
            disc_type: Type of discrepancy
            
        Returns:
            Severity classification
        """
        critical_types = self.rules.get('discrepancy_classification', {}).get('critical', [])
        major_types = self.rules.get('discrepancy_classification', {}).get('major', [])
        
        if disc_type.value in critical_types:
            return DiscrepancySeverity.CRITICAL
        elif disc_type.value in major_types:
            return DiscrepancySeverity.MAJOR
        else:
            return DiscrepancySeverity.MINOR
    
    def fuzzy_match_names(
        self,
        name_a: str,
        name_b: str
    ) -> float:
        """
        Calculate fuzzy match score between two names.
        
        Uses token_sort_ratio for best results with name ordering differences.
        
        Args:
            name_a: First name
            name_b: Second name
            
        Returns:
            Match score 0-100
        """
        if not name_a or not name_b:
            return 0.0
            
        # Normalize names
        name_a = str(name_a).upper().strip()
        name_b = str(name_b).upper().strip()
        
        # Use token sort ratio (handles word reordering)
        score = fuzz.token_sort_ratio(name_a, name_b)
        
        return score
    
    def find_best_match(
        self,
        query_name: str,
        candidates: pd.DataFrame,
        name_column: str = 'name',
        id_column: str = 'staff_id',
        limit: int = 3
    ) -> List[Tuple[str, str, float]]:
        """
        Find best matching candidates for a name.
        
        Args:
            query_name: Name to match
            candidates: DataFrame with candidate records
            name_column: Column containing names
            id_column: Column containing IDs
            limit: Maximum candidates to return
            
        Returns:
            List of (staff_id, name, score) tuples
        """
        if candidates.empty or name_column not in candidates.columns:
            return []
            
        candidate_names = candidates[name_column].fillna('').astype(str).tolist()
        candidate_ids = candidates[id_column].fillna('').astype(str).tolist()
        
        # Get best matches using rapidfuzz
        matches = process.extract(
            str(query_name).upper().strip(),
            candidate_names,
            scorer=fuzz.token_sort_ratio,
            limit=limit
        )
        
        results = []
        for match_name, score, idx in matches:
            if idx < len(candidate_ids):
                results.append((candidate_ids[idx], match_name, score))
                
        return results
    
    def detect_ghost_workers(
        self,
        tfr_df: pd.DataFrame,
        hr_df: pd.DataFrame,
        partner: str
    ) -> List[Discrepancy]:
        """
        Detect ghost workers: in TFR but not in HR master.
        
        Args:
            tfr_df: TFR timesheet data
            hr_df: HR master data
            partner: Partner identifier
            
        Returns:
            List of ghost worker discrepancies
        """
        discrepancies = []
        
        # Get sets of staff IDs
        tfr_ids = set(tfr_df['staff_id'].astype(str).str.strip())
        hr_ids = set(hr_df['staff_id'].astype(str).str.strip())
        
        # Find IDs in TFR but not HR
        ghost_ids = tfr_ids - hr_ids
        
        auto_accept = self.rules.get('fuzzy_matching', {}).get('auto_accept_threshold', 95)
        review_threshold = self.rules.get('fuzzy_matching', {}).get('review_threshold', 85)
        
        for ghost_id in ghost_ids:
            # Get TFR record
            tfr_record = tfr_df[tfr_df['staff_id'].astype(str).str.strip() == ghost_id].iloc[0]
            tfr_name = str(tfr_record.get('name', 'Unknown'))
            tfr_rate = tfr_record.get('daily_rate', 0)
            
            # Try fuzzy match against HR
            matches = self.find_best_match(tfr_name, hr_df)
            
            if matches and matches[0][2] >= auto_accept:
                # High confidence match - likely a typo in ID
                best_id, best_name, score = matches[0]
                discrepancies.append(Discrepancy(
                    id=self._generate_discrepancy_id(),
                    type=DiscrepancyType.FUZZY_MATCH_REVIEW,
                    severity=DiscrepancySeverity.MAJOR,
                    partner=partner,
                    staff_id=ghost_id,
                    staff_name=tfr_name,
                    description=f"TFR ID '{ghost_id}' not in HR, but name matches '{best_name}' (ID: {best_id})",
                    confidence=score,
                    source_a="TFR",
                    source_b="HR Master",
                    value_a=ghost_id,
                    value_b=best_id
                ))
            elif matches and matches[0][2] >= review_threshold:
                # Needs review
                best_id, best_name, score = matches[0]
                discrepancies.append(Discrepancy(
                    id=self._generate_discrepancy_id(),
                    type=DiscrepancyType.FUZZY_MATCH_REVIEW,
                    severity=DiscrepancySeverity.MAJOR,
                    partner=partner,
                    staff_id=ghost_id,
                    staff_name=tfr_name,
                    description=f"TFR has '{tfr_name}' (ID: {ghost_id}), possible match to HR '{best_name}'",
                    confidence=score,
                    source_a="TFR",
                    source_b="HR Master",
                    value_a=tfr_name,
                    value_b=best_name
                ))
            else:
                # No good match - true ghost worker
                discrepancies.append(Discrepancy(
                    id=self._generate_discrepancy_id(),
                    type=DiscrepancyType.GHOST_WORKER,
                    severity=DiscrepancySeverity.CRITICAL,
                    partner=partner,
                    staff_id=ghost_id,
                    staff_name=tfr_name,
                    description=f"Ghost worker in TFR: '{tfr_name}' (ID: {ghost_id}) not found in HR",
                    confidence=matches[0][2] if matches else 0,
                    amount_hkd=float(tfr_rate) if pd.notna(tfr_rate) else None,
                    source_a="TFR",
                    source_b="HR Master"
                ))
                
        return discrepancies
    
    def detect_rate_deviations(
        self,
        df_a: pd.DataFrame,
        df_b: pd.DataFrame,
        partner: str,
        source_a_name: str = "TFR",
        source_b_name: str = "HR Master"
    ) -> List[Discrepancy]:
        """
        Detect rate deviations between two sources.
        
        Args:
            df_a: First DataFrame
            df_b: Second DataFrame
            partner: Partner identifier
            source_a_name: Name of first source
            source_b_name: Name of second source
            
        Returns:
            List of rate deviation discrepancies
        """
        discrepancies = []
        tolerance = self.rules.get('rate_comparison', {}).get('deviation_tolerance_pct', 15)
        
        # Merge on staff_id
        merged = pd.merge(
            df_a[['staff_id', 'name', 'daily_rate']].rename(columns={'daily_rate': 'rate_a', 'name': 'name_a'}),
            df_b[['staff_id', 'name', 'daily_rate']].rename(columns={'daily_rate': 'rate_b', 'name': 'name_b'}),
            on='staff_id',
            how='inner'
        )
        
        for _, row in merged.iterrows():
            rate_a = row.get('rate_a')
            rate_b = row.get('rate_b')
            
            if pd.isna(rate_a) or pd.isna(rate_b) or rate_a == 0 or rate_b == 0:
                continue
                
            # Calculate deviation percentage
            deviation = abs(rate_a - rate_b) / max(rate_a, rate_b) * 100
            
            if deviation > 50:
                # Extreme deviation
                discrepancies.append(Discrepancy(
                    id=self._generate_discrepancy_id(),
                    type=DiscrepancyType.RATE_DEVIATION_EXTREME,
                    severity=DiscrepancySeverity.CRITICAL,
                    partner=partner,
                    staff_id=str(row['staff_id']),
                    staff_name=str(row.get('name_a', row.get('name_b', 'Unknown'))),
                    description=f"Extreme rate deviation {deviation:.1f}%: ${rate_a:.0f} vs ${rate_b:.0f}",
                    amount_hkd=abs(rate_a - rate_b),
                    source_a=source_a_name,
                    source_b=source_b_name,
                    value_a=rate_a,
                    value_b=rate_b
                ))
            elif deviation > tolerance:
                # Deviation above tolerance
                discrepancies.append(Discrepancy(
                    id=self._generate_discrepancy_id(),
                    type=DiscrepancyType.RATE_DEVIATION,
                    severity=DiscrepancySeverity.MAJOR,
                    partner=partner,
                    staff_id=str(row['staff_id']),
                    staff_name=str(row.get('name_a', row.get('name_b', 'Unknown'))),
                    description=f"Rate deviation {deviation:.1f}%: ${rate_a:.0f} ({source_a_name}) vs ${rate_b:.0f} ({source_b_name})",
                    amount_hkd=abs(rate_a - rate_b),
                    source_a=source_a_name,
                    source_b=source_b_name,
                    value_a=rate_a,
                    value_b=rate_b
                ))
                
        return discrepancies
    
    def detect_cross_partner_discrepancies(
        self,
        partner_a_hr: pd.DataFrame,
        partner_b_hr: pd.DataFrame,
        partner_a_name: str = "Dragages",
        partner_b_name: str = "Gammon"
    ) -> List[Discrepancy]:
        """
        Detect discrepancies between partners.
        
        Checks for:
        - Duplicate staff IDs across partners
        - Rate differences for same roles
        
        Args:
            partner_a_hr: Partner A HR data
            partner_b_hr: Partner B HR data
            partner_a_name: Partner A identifier
            partner_b_name: Partner B identifier
            
        Returns:
            List of cross-partner discrepancies
        """
        discrepancies = []
        tolerance = self.rules.get('rate_comparison', {}).get('deviation_tolerance_pct', 15)
        
        # Check for duplicate IDs
        ids_a = set(partner_a_hr['staff_id'].astype(str).str.strip())
        ids_b = set(partner_b_hr['staff_id'].astype(str).str.strip())
        duplicate_ids = ids_a & ids_b
        
        for dup_id in duplicate_ids:
            rec_a = partner_a_hr[partner_a_hr['staff_id'].astype(str).str.strip() == dup_id].iloc[0]
            rec_b = partner_b_hr[partner_b_hr['staff_id'].astype(str).str.strip() == dup_id].iloc[0]
            
            discrepancies.append(Discrepancy(
                id=self._generate_discrepancy_id(),
                type=DiscrepancyType.DUPLICATE_ID,
                severity=DiscrepancySeverity.CRITICAL,
                partner="JV",
                staff_id=dup_id,
                staff_name=f"{rec_a.get('name', 'Unknown')} / {rec_b.get('name', 'Unknown')}",
                description=f"Duplicate staff ID '{dup_id}' found in both partners",
                source_a=partner_a_name,
                source_b=partner_b_name,
                value_a=str(rec_a.get('name', '')),
                value_b=str(rec_b.get('name', ''))
            ))
            
        # Compare rates for same roles
        role_col_a = 'role_normalized' if 'role_normalized' in partner_a_hr.columns else 'role'
        role_col_b = 'role_normalized' if 'role_normalized' in partner_b_hr.columns else 'role'
        
        # Get average rates by role for each partner
        avg_rates_a = partner_a_hr.groupby(role_col_a)['daily_rate'].mean()
        avg_rates_b = partner_b_hr.groupby(role_col_b)['daily_rate'].mean()
        
        common_roles = set(avg_rates_a.index) & set(avg_rates_b.index)
        
        for role in common_roles:
            rate_a = avg_rates_a[role]
            rate_b = avg_rates_b[role]
            
            if pd.isna(rate_a) or pd.isna(rate_b) or rate_a == 0 or rate_b == 0:
                continue
                
            deviation = abs(rate_a - rate_b) / max(rate_a, rate_b) * 100
            
            if deviation > tolerance:
                discrepancies.append(Discrepancy(
                    id=self._generate_discrepancy_id(),
                    type=DiscrepancyType.CROSS_PARTNER_RATE_DIFF,
                    severity=DiscrepancySeverity.MAJOR,
                    partner="JV",
                    staff_id="N/A",
                    staff_name=f"Role: {role}",
                    description=f"Cross-partner rate diff {deviation:.1f}% for {role}: ${rate_a:.0f} ({partner_a_name}) vs ${rate_b:.0f} ({partner_b_name})",
                    amount_hkd=abs(rate_a - rate_b),
                    source_a=partner_a_name,
                    source_b=partner_b_name,
                    value_a=rate_a,
                    value_b=rate_b
                ))
                
        return discrepancies
    
    def three_way_match(
        self,
        hr_df: pd.DataFrame,
        tfr_df: pd.DataFrame,
        attendance_df: pd.DataFrame,
        partner: str
    ) -> ReconciliationResult:
        """
        Perform 3-way matching: HR ↔ TFR ↔ Attendance.
        
        Args:
            hr_df: HR master data
            tfr_df: TFR timesheet data
            attendance_df: Site attendance data
            partner: Partner identifier
            
        Returns:
            ReconciliationResult with all discrepancies
        """
        result = ReconciliationResult(success=True)
        
        try:
            # Detect ghost workers (TFR vs HR)
            ghost_discs = self.detect_ghost_workers(tfr_df, hr_df, partner)
            result.discrepancies.extend(ghost_discs)
            
            # Detect rate deviations (TFR vs HR)
            rate_discs = self.detect_rate_deviations(tfr_df, hr_df, partner, "TFR", "HR Master")
            result.discrepancies.extend(rate_discs)
            
            # Detect attendance mismatches
            tfr_ids = set(tfr_df['staff_id'].astype(str).str.strip())
            att_ids = set(attendance_df['staff_id'].astype(str).str.strip())
            
            missing_attendance = tfr_ids - att_ids
            for missing_id in missing_attendance:
                tfr_rec = tfr_df[tfr_df['staff_id'].astype(str).str.strip() == missing_id].iloc[0]
                result.discrepancies.append(Discrepancy(
                    id=self._generate_discrepancy_id(),
                    type=DiscrepancyType.MISSING_ATTENDANCE,
                    severity=DiscrepancySeverity.MINOR,
                    partner=partner,
                    staff_id=missing_id,
                    staff_name=str(tfr_rec.get('name', 'Unknown')),
                    description=f"Staff in TFR but no attendance record",
                    source_a="TFR",
                    source_b="Site Attendance"
                ))
                
            # Count matches
            matched = tfr_ids & att_ids
            result.matched_count = len(matched)
            result.unmatched_count = len(missing_attendance)
            
            # Create reconciled dataset
            hr_subset = hr_df.copy()
            hr_subset['_match_status'] = 'HR Only'
            
            for idx, row in hr_subset.iterrows():
                staff_id = str(row['staff_id']).strip()
                if staff_id in tfr_ids:
                    if staff_id in att_ids:
                        hr_subset.at[idx, '_match_status'] = 'Full Match'
                    else:
                        hr_subset.at[idx, '_match_status'] = 'HR+TFR Only'
                        
            result.reconciled_data = hr_subset
            
            # Summary stats
            result.summary = {
                'total_hr': len(hr_df),
                'total_tfr': len(tfr_ids),
                'total_attendance': len(att_ids),
                'matched': result.matched_count,
                'ghost_workers': len(ghost_discs),
                'rate_deviations': len(rate_discs),
                'missing_attendance': len(missing_attendance)
            }
            
        except Exception as e:
            logger.error(f"Reconciliation failed: {e}")
            result.success = False
            
        return result
    
    def reconcile_all(
        self,
        data: Dict[str, Dict[str, pd.DataFrame]]
    ) -> ReconciliationResult:
        """
        Run full reconciliation across all partners and sources.
        
        Args:
            data: Nested dict: {partner: {source_type: DataFrame}}
            
        Returns:
            Combined ReconciliationResult
        """
        combined = ReconciliationResult(success=True)
        
        partners = list(data.keys())
        
        for partner in partners:
            partner_data = data.get(partner, {})
            
            hr_df = partner_data.get('hr_master')
            tfr_df = partner_data.get('tfr')
            attendance_df = partner_data.get('attendance')
            
            if hr_df is not None and tfr_df is not None:
                if attendance_df is None:
                    attendance_df = pd.DataFrame(columns=['staff_id', 'date'])
                    
                result = self.three_way_match(hr_df, tfr_df, attendance_df, partner)
                combined.discrepancies.extend(result.discrepancies)
                combined.matched_count += result.matched_count
                combined.unmatched_count += result.unmatched_count
                
                for key, val in result.summary.items():
                    combined.summary[f"{partner}_{key}"] = val
                    
        # Cross-partner checks
        if len(partners) >= 2:
            hr_a = data.get(partners[0], {}).get('hr_master')
            hr_b = data.get(partners[1], {}).get('hr_master')
            
            if hr_a is not None and hr_b is not None:
                cross_discs = self.detect_cross_partner_discrepancies(
                    hr_a, hr_b, partners[0], partners[1]
                )
                combined.discrepancies.extend(cross_discs)
                
        # Update summary
        combined.summary['total_discrepancies'] = len(combined.discrepancies)
        combined.summary['critical'] = combined.get_critical_count()
        combined.summary['major'] = combined.get_major_count()
        combined.summary['minor'] = combined.get_minor_count()
        
        return combined
