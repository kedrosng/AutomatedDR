"""
Reporting module for JV Cost Control System
Generates Excel/PDF discrepancy reports and evidence packs
"""

import pandas as pd
from datetime import datetime
import os
from typing import List, Dict, Any
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import json

logger = logging.getLogger(__name__)

def generate_discrepancy_report(discrepancies: List[Dict[str, Any]], 
                              output_path: str = None) -> str:
    """
    Generate a comprehensive discrepancy report with multiple sheets
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"outputs/discrepancy_report_{timestamp}.xlsx"
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Create sheets for different discrepancy types
    critical_items = [d for d in discrepancies if d.get('severity', '').lower() == 'critical']
    major_items = [d for d in discrepancies if d.get('severity', '').lower() == 'major']
    minor_items = [d for d in discrepancies if d.get('severity', '').lower() == 'minor']
    
    # Sheet 1: Critical Items
    if critical_items:
        ws_critical = wb.create_sheet(title="Critical Items")
        df_critical = pd.DataFrame(critical_items)
        populate_worksheet(ws_critical, df_critical, "Critical Discrepancies")
    
    # Sheet 2: Major Items
    if major_items:
        ws_major = wb.create_sheet(title="Major Items")
        df_major = pd.DataFrame(major_items)
        populate_worksheet(ws_major, df_major, "Major Discrepancies")
    
    # Sheet 3: Raw Data Snapshot
    ws_raw = wb.create_sheet(title="Raw Data")
    df_all = pd.DataFrame(discrepancies)
    populate_worksheet(ws_raw, df_all, "All Discrepancies Raw Data")
    
    # Sheet 4: Summary
    ws_summary = wb.create_sheet(title="Summary")
    create_summary_sheet(ws_summary, critical_items, major_items, minor_items)
    
    wb.save(output_path)
    logger.info(f"Discrepancy report generated: {output_path}")
    
    return output_path

def populate_worksheet(worksheet, df: pd.DataFrame, title: str):
    """
    Populate a worksheet with DataFrame data and apply formatting
    """
    # Add title
    worksheet.cell(row=1, column=1, value=title).font = Font(size=14, bold=True)
    
    # Add a blank row
    worksheet.append([])
    
    # Add DataFrame headers
    for col_num, value in enumerate(df.columns.values, 1):
        cell = worksheet.cell(row=worksheet.max_row + 1, column=col_num, value=value)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    # Add DataFrame rows
    for row in dataframe_to_rows(df, index=False, header=False):
        worksheet.append(row)
    
    # Apply borders
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in worksheet.iter_rows(min_row=3, max_row=worksheet.max_row, min_col=1, max_col=df.shape[1]):
        for cell in row:
            cell.border = thin_border

def create_summary_sheet(worksheet, critical_items, major_items, minor_items):
    """
    Create a summary sheet with key metrics
    """
    worksheet.cell(row=1, column=1, value="Discrepancy Report Summary").font = Font(size=16, bold=True)
    
    # Add summary data
    summary_data = [
        ["Category", "Count", "Percentage"],
        ["Critical", len(critical_items), f"{len(critical_items)/(len(critical_items)+len(major_items)+len(minor_items))*100:.1f}%" if (len(critical_items)+len(major_items)+len(minor_items)) > 0 else "0%"],
        ["Major", len(major_items), f"{len(major_items)/(len(critical_items)+len(major_items)+len(minor_items))*100:.1f}%" if (len(critical_items)+len(major_items)+len(minor_items)) > 0 else "0%"],
        ["Minor", len(minor_items), f"{len(minor_items)/(len(critical_items)+len(major_items)+len(minor_items))*100:.1f}%" if (len(critical_items)+len(major_items)+len(minor_items)) > 0 else "0%"],
        ["Total", len(critical_items)+len(major_items)+len(minor_items), "100%"]
    ]
    
    for row_data in summary_data:
        worksheet.append(row_data)
    
    # Format headers
    for col_num in range(1, 4):
        cell = worksheet.cell(row=2, column=col_num)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    # Format total row
    for col_num in range(1, 4):
        cell = worksheet.cell(row=6, column=col_num)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

def generate_evidence_pack(discrepancy: Dict[str, Any], 
                          raw_sources: Dict[str, pd.DataFrame] = None,
                          output_path: str = None) -> str:
    """
    Generate an evidence pack PDF combining raw TFR page + HR record + site log for a single staff ID
    For now, we'll create a basic PDF with the information using fpdf2 library
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        staff_id = discrepancy.get('staff_id', 'unknown')
        output_path = f"outputs/evidence_pack_{staff_id}_{timestamp}.pdf"
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Since we don't have the fpdf2 library installed, we'll create a placeholder
    # In a real implementation, we would use fpdf2 to generate the actual PDF
    # For now, we'll just create a JSON file with the evidence data
    evidence_data = {
        "discrepancy_info": discrepancy,
        "raw_sources": {k: v.to_dict() for k, v in raw_sources.items()} if raw_sources else {},
        "generated_at": datetime.now().isoformat(),
        "report_type": "evidence_pack"
    }
    
    evidence_path = output_path.replace('.pdf', '_evidence.json')
    with open(evidence_path, 'w') as f:
        json.dump(evidence_data, f, indent=2, default=str)
    
    logger.info(f"Evidence pack generated: {evidence_path}")
    
    return evidence_path

def generate_compliance_report(df: pd.DataFrame, output_path: str = None) -> str:
    """
    Generate a compliance report for HK Labour Department
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"outputs/compliance_report_{timestamp}.xlsx"
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Create compliance issues sheet
    ws_compliance = wb.create_sheet(title="Compliance Issues")
    
    # Identify compliance issues
    compliance_issues = []
    
    if 'fdh_permit' in df.columns and 'nationality' in df.columns:
        # Check for missing FDH permits for foreign workers
        foreign_workers = df[
            (~df['nationality'].str.contains('hk|china|hong kong', case=False, na=False))
        ]
        
        missing_permits = foreign_workers[
            foreign_workers['fdh_permit'].isna() | 
            (foreign_workers['fdh_permit'] == "") |
            (foreign_workers['fdh_permit'] == "None")
        ]
        
        for _, row in missing_permits.iterrows():
            compliance_issues.append({
                'issue_type': 'Missing FDH Permit',
                'staff_id': row.get('staff_id', 'Unknown'),
                'name': row.get('name', 'Unknown'),
                'nationality': row.get('nationality', 'Unknown'),
                'role': row.get('role', 'Unknown'),
                'date_identified': datetime.now().strftime('%Y-%m-%d')
            })
    
    # Add other compliance checks here as needed
    
    if compliance_issues:
        df_compliance = pd.DataFrame(compliance_issues)
        populate_worksheet(ws_compliance, df_compliance, "HK Labour Compliance Issues")
    else:
        ws_compliance.append(["No compliance issues identified"])
        ws_compliance.cell(row=1, column=1).font = Font(bold=True)
    
    # Create summary sheet
    ws_summary = wb.create_sheet(title="Compliance Summary")
    ws_summary.cell(row=1, column=1, value="Compliance Report Summary").font = Font(size=16, bold=True)
    
    summary_data = [
        ["Issue Category", "Count"],
        ["Missing FDH Permits", len(compliance_issues)],
        ["Total Workers", len(df)],
        ["Foreign Workers", len(df) if 'nationality' in df.columns else 0]
    ]
    
    for row_data in summary_data:
        ws_summary.append(row_data)
    
    # Format headers
    for col_num in range(1, 3):
        cell = ws_summary.cell(row=2, column=col_num)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
    
    wb.save(output_path)
    logger.info(f"Compliance report generated: {output_path}")
    
    return output_path

def format_currency_hkd(amount: float) -> str:
    """
    Format amount as HKD currency
    """
    return f"HKD {amount:,.2f}"

def create_dashboard_metrics(discrepancies: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Create dashboard metrics from discrepancies
    """
    total_discrepancies = len(discrepancies)
    critical_count = sum(1 for d in discrepancies if d.get('severity', '').lower() == 'critical')
    major_count = sum(1 for d in discrepancies if d.get('severity', '').lower() == 'major')
    minor_count = sum(1 for d in discrepancies if d.get('severity', '').lower() == 'minor')
    
    # Calculate potential financial impact
    total_potential_impact = sum(
        d.get('amount_impact_hkd', 0) for d in discrepancies 
        if d.get('amount_impact_hkd') is not None
    )
    
    return {
        'total_discrepancies': total_discrepancies,
        'critical_count': critical_count,
        'major_count': major_count,
        'minor_count': minor_count,
        'total_potential_impact_hkd': total_potential_impact,
        'average_impact_per_discrepancy': total_potential_impact / total_discrepancies if total_discrepancies > 0 else 0
    }

def export_forecast_data(df: pd.DataFrame, output_format: str = 'excel', 
                        output_path: str = None) -> str:
    """
    Export forecast data in specified format
    """
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ext = '.xlsx' if output_format.lower() == 'excel' else '.csv'
        output_path = f"outputs/forecast_data_{timestamp}{ext}"
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    if output_format.lower() == 'excel':
        df.to_excel(output_path, index=False)
    else:
        df.to_csv(output_path, index=False)
    
    logger.info(f"Forecast data exported: {output_path}")
    
    return output_path