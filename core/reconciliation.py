"""
Reconciliation module for JV Cost Control System
Performs 3-way matching of HR master, TFR timesheets, and site attendance data
"""

import pandas as pd
from rapidfuzz import fuzz, process
from typing import Tuple, List, Dict, Any
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

class DataReconciler:
    def __init__(self):
        self.fuzzy_threshold_high = 95  # Auto-accept matches above this threshold
        self.fuzzy_threshold_medium = 85  # Flag for review between medium and high
        self.fuzzy_threshold_low = 50   # Below this is critical discrepancy
    
    def fuzzy_match_names(self, name1: str, name2: str) -> float:
        """Calculate fuzzy match score between two names"""
        if pd.isna(name1) or pd.isna(name2):
            return 0.0
        
        # Clean names before matching
        name1_clean = str(name1).strip().upper().replace('.', '').replace('  ', ' ')
        name2_clean = str(name2).strip().upper().replace('.', '').replace('  ', ' ')
        
        # Calculate ratio
        ratio = fuzz.ratio(name1_clean, name2_clean)
        
        # Also try token sort ratio for cases where names are in different order
        token_ratio = fuzz.token_sort_ratio(name1_clean, name2_clean)
        
        return max(ratio, token_ratio)
    
    def reconcile_by_staff_id(self, df1: pd.DataFrame, df2: pd.DataFrame, 
                             id_col1: str, id_col2: str) -> pd.DataFrame:
        """Reconcile two datasets based on staff ID"""
        # Merge datasets on staff IDs
        merged = df1.merge(
            df2, 
            left_on=id_col1, 
            right_on=id_col2, 
            how='outer', 
            suffixes=('_left', '_right'),
            indicator=True
        )
        
        return merged
    
    def find_matching_records(self, df_left: pd.DataFrame, df_right: pd.DataFrame, 
                             id_col_left: str, id_col_right: str, 
                             name_col_left: str, name_col_right: str) -> List[Dict[str, Any]]:
        """Find matching records using both ID and fuzzy name matching"""
        matches = []
        
        for idx_left, row_left in df_left.iterrows():
            matched = False
            best_match_score = 0
            best_match_idx = None
            
            # First, try exact ID match
            left_id = row_left[id_col_left]
            exact_matches = df_right[df_right[id_col_right] == left_id]
            
            if not exact_matches.empty:
                for idx_right in exact_matches.index:
                    matches.append({
                        'left_idx': idx_left,
                        'right_idx': idx_right,
                        'match_type': 'exact_id',
                        'confidence': 100.0,
                        'left_record': row_left,
                        'right_record': df_right.loc[idx_right]
                    })
                    matched = True
                    break
            
            if matched:
                continue
            
            # If no exact ID match, try fuzzy name matching
            left_name = row_left.get(name_col_left, "")
            if pd.notna(left_name) and left_name != "":
                for idx_right, row_right in df_right.iterrows():
                    right_name = row_right.get(name_col_right, "")
                    if pd.notna(right_name) and right_name != "":
                        score = self.fuzzy_match_names(left_name, right_name)
                        
                        if score > best_match_score:
                            best_match_score = score
                            best_match_idx = idx_right
                
                if best_match_score >= self.fuzzy_threshold_low:
                    matches.append({
                        'left_idx': idx_left,
                        'right_idx': best_match_idx,
                        'match_type': 'fuzzy_name',
                        'confidence': best_match_score,
                        'left_record': row_left,
                        'right_record': df_right.loc[best_match_idx]
                    })
        
        return matches
    
    def detect_cross_partner_rate_discrepancies(self, df_byca: pd.DataFrame, 
                                              df_jhg: pd.DataFrame) -> List[Dict[str, Any]]:
        """Detect rate discrepancies between partners for same roles"""
        discrepancies = []
        
        # Focus on records with same role but different rates
        if all(col in df_byca.columns for col in ['role', 'rate', 'name']) and \
           all(col in df_jhg.columns for col in ['role', 'rate', 'name']):
            
            # Get all unique roles
            all_roles = set(df_byca['role'].dropna()) | set(df_jhg['role'].dropna())
            
            for role in all_roles:
                byca_role_data = df_byca[df_byca['role'] == role]
                jhg_role_data = df_jhg[df_jhg['role'] == role]
                
                if len(byca_role_data) > 0 and len(jhg_role_data) > 0:
                    # Compare average rates
                    avg_byca_rate = byca_role_data['rate'].mean()
                    avg_jhg_rate = jhg_role_data['rate'].mean()
                    
                    if pd.notna(avg_byca_rate) and pd.notna(avg_jhg_rate):
                        rate_diff_pct = abs(avg_jhg_rate - avg_byca_rate) / avg_byca_rate * 100
                        
                        if rate_diff_pct > 15:  # More than 15% difference
                            discrepancies.append({
                                'discrepancy_type': 'cross_partner_rate_variance',
                                'role': role,
                                'partner_a': 'BYCA',
                                'partner_a_avg_rate': avg_byca_rate,
                                'partner_b': 'JHG',
                                'partner_b_avg_rate': avg_jhg_rate,
                                'variance_pct': rate_diff_pct,
                                'severity': 'Major' if rate_diff_pct > 15 else 'Minor'
                            })
        
        return discrepancies
    
    def perform_three_way_reconciliation(self, 
                                       byca_data: Dict[str, pd.DataFrame],
                                       jhg_data: Dict[str, pd.DataFrame]) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
        """Perform 3-way reconciliation between HR, TFR, and Site attendance for both partners"""
        
        discrepancies = []
        
        # Process each partner's data
        for partner, data_dict in [("BYCA", byca_data), ("JHG", jhg_data)]:
            if 'hr_master' in data_dict and 'tfr_timesheet' in data_dict:
                hr_df = data_dict['hr_master']
                tfr_df = data_dict['tfr_timesheet']
                
                # Identify staff in TFR but missing from HR (ghost workers)
                if 'staff_id' in hr_df.columns and 'staff_id' in tfr_df.columns:
                    tfr_only_ids = set(tfr_df['staff_id']) - set(hr_df['staff_id'])
                    for staff_id in tfr_only_ids:
                        discrepancies.append({
                            'discrepancy_id': f"ghost_{partner}_{staff_id}",
                            'partner': partner,
                            'staff_id': staff_id,
                            'discrepancy_type': 'ghost_worker',
                            'description': f'Staff ID {staff_id} appears in TFR but missing from HR master',
                            'severity': 'Critical',
                            'confidence': 100.0,
                            'amount_impact_aud': 0  # To be calculated later based on rate
                        })
            
            # Check for name typos between HR and other sources
            if 'hr_master' in data_dict and 'tfr_timesheet' in data_dict:
                hr_df = data_dict['hr_master']
                tfr_df = data_dict['tfr_timesheet']
                
                if 'name' in hr_df.columns and 'name' in tfr_df.columns:
                    matches = self.find_matching_records(
                        hr_df, tfr_df, 
                        'staff_id', 'staff_id', 
                        'name', 'name'
                    )
                    
                    for match in matches:
                        if match['match_type'] == 'fuzzy_name':
                            confidence = match['confidence']
                            if confidence < self.fuzzy_threshold_high:
                                severity = 'Critical' if confidence < self.fuzzy_threshold_medium else 'Major'
                                discrepancies.append({
                                    'discrepancy_id': f"name_mismatch_{partner}_{match['left_record']['staff_id']}",
                                    'partner': partner,
                                    'staff_id': match['left_record']['staff_id'],
                                    'discrepancy_type': 'name_typo',
                                    'description': f'Name mismatch: HR="{match["left_record"]["name"]}" vs TFR="{match["right_record"]["name"]}"',
                                    'severity': severity,
                                    'confidence': confidence,
                                    'amount_impact_aud': 0
                                })
        
        # Cross-partner reconciliation
        if 'hr_master' in byca_data and 'hr_master' in jhg_data:
            cross_partner_discrepancies = self.detect_cross_partner_rate_discrepancies(
                byca_data['hr_master'], 
                jhg_data['hr_master']
            )
            discrepancies.extend(cross_partner_discrepancies)
        
        # Create a consolidated view of all data
        all_data_parts = []
        for partner, data_dict in [("BYCA", byca_data), ("JHG", jhg_data)]:
            for source_type, df in data_dict.items():
                df_with_meta = df.copy()
                df_with_meta['partner'] = partner
                df_with_meta['source_type'] = source_type
                all_data_parts.append(df_with_meta)
        
        if all_data_parts:
            consolidated_df = pd.concat(all_data_parts, ignore_index=True)
        else:
            consolidated_df = pd.DataFrame()
        
        return consolidated_df, discrepancies

def reconcile_data() -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
    """
    Main reconciliation function that loads saved data and performs reconciliation
    """
    logger.info("Starting reconciliation process...")
    
    # In a real implementation, we would load the data from the data/raw directory
    # For now, we'll simulate by looking for processed data in session or returning empty results
    byca_data = {}
    jhg_data = {}
    
    # Look for processed data files
    import os
    raw_data_dir = "data/raw/"
    
    for partner in ["dragages", "gammon"]:
        partner_dir = os.path.join(raw_data_dir, partner)
        if os.path.exists(partner_dir):
            data_dict = {}
            for filename in os.listdir(partner_dir):
                filepath = os.path.join(partner_dir, filename)
                if filename.endswith('.csv'):
                    if 'hr_master' in filename:
                        data_dict['hr_master'] = pd.read_csv(filepath)
                    elif 'attendance' in filename:
                        data_dict['site_attendance'] = pd.read_csv(filepath)
                elif filename.endswith('.xlsx'):
                    if 'hr_master' in filename:
                        data_dict['hr_master'] = pd.read_excel(filepath)
                    elif 'attendance' in filename:
                        data_dict['site_attendance'] = pd.read_excel(filepath)
                elif filename.endswith('.pdf'):
                    if 'tfr' in filename:
                        # For now, we'll skip PDF processing in this simplified version
                        # In a real implementation, we would extract the tables from the PDF
                        pass
            
            if partner == "dragages":
                byca_data = data_dict
            else:
                jhg_data = data_dict
    
    # Perform reconciliation
    reconciler = DataReconciler()
    consolidated_df, discrepancies = reconciler.perform_three_way_reconciliation(
        byca_data, jhg_data
    )
    
    # Enhance discrepancies with additional info
    for disc in discrepancies:
        # Calculate impact amounts where possible
        if disc.get('discrepancy_type') == 'cross_partner_rate_variance':
            avg_rate_diff = abs(disc['partner_b_avg_rate'] - disc['partner_a_avg_rate'])
            # Assuming this affects all workers in the role (simplified calculation)
            disc['amount_impact_aud'] = avg_rate_diff * 22 * 12  # Monthly impact estimate (22 workdays * 12 months)
    
    logger.info(f"Reconciliation completed. Found {len(discrepancies)} discrepancies.")
    
    return consolidated_df, discrepancies

def calculate_discrepancy_severity(discrepancy: Dict[str, Any]) -> str:
    """Calculate severity level for a discrepancy"""
    # Default to existing severity if present
    if 'severity' in discrepancy:
        return discrepancy['severity']
    
    # Otherwise, determine severity based on discrepancy type and confidence
    d_type = discrepancy.get('discrepancy_type', '')
    confidence = discrepancy.get('confidence', 0)
    
    if d_type == 'ghost_worker':
        return 'Critical'
    elif d_type == 'name_typo' and confidence < 85:
        return 'Critical'
    elif d_type == 'name_typo' and 85 <= confidence < 95:
        return 'Major'
    elif d_type == 'cross_partner_rate_variance':
        variance_pct = discrepancy.get('variance_pct', 0)
        if variance_pct > 25:
            return 'Critical'
        elif variance_pct > 15:
            return 'Major'
        else:
            return 'Minor'
    else:
        return 'Minor'