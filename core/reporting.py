"""
Report Generation Module for JV Cost Reconciliation (Australia)

Generates:
- Multi-tab Excel discrepancy reports
- PDF evidence packs for individual staff
- Fair Work / Compliance reports (Visa, Superannuation)
"""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, Image
)

# Configure logging
logger = logging.getLogger(__name__)


@dataclass
class ReportConfig:
    """Configuration for report generation."""
    company_name: str = "JV Cost Control"
    project_name: str = "Australian Infrastructure Project"
    output_dir: Path = Path("outputs")
    logo_path: Optional[Path] = None


class ReportGenerator:
    """
    Generates various report formats for JV reconciliation.
    
    Supports:
    - Excel reports with critical/major breakdowns
    - PDF evidence packs
    - Compliance reports for Fair Work / Project Controls
    """
    
    def __init__(self, config: Optional[ReportConfig] = None):
        """Initialize the report generator."""
        self.config = config or ReportConfig()
        self.config.output_dir = Path(self.config.output_dir)
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Excel styles
        self._header_font = Font(bold=True, color="FFFFFF")
        self._header_fill = PatternFill("solid", fgColor="003366") # Navy Blue
        self._critical_fill = PatternFill("solid", fgColor="D30026") # BYCA Red
        self._major_fill = PatternFill("solid", fgColor="F37021") # Orange
        self._minor_fill = PatternFill("solid", fgColor="FFD700") # Gold
        self._border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
    def generate_excel_report(
        self,
        discrepancies: List[Any],
        raw_data: Optional[Dict[str, pd.DataFrame]] = None,
        filename: Optional[str] = None
    ) -> Path:
        """Generate multi-tab Excel discrepancy report."""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"discrepancy_report_au_{timestamp}.xlsx"
            
        output_path = self.config.output_dir / filename
        
        # Convert discrepancies to DataFrame
        disc_data = []
        for d in discrepancies:
            disc_data.append({
                'ID': d.id,
                'Severity': d.severity.value,
                'Type': d.type.value,
                'Partner': d.partner,
                'Staff ID': d.staff_id,
                'Staff Name': d.staff_name,
                'Description': d.description,
                'Confidence %': d.confidence,
                'Amount (AUD)': d.amount_hkd, # Reusing internal field
                'Source A': d.source_a,
                'Source B': d.source_b,
                'Value A': d.value_a,
                'Value B': d.value_b,
                'Status': d.status,
                'Created At': d.created_at
            })
            
        df_all = pd.DataFrame(disc_data)
        
        wb = Workbook()
        wb.remove(wb.active)
        
        # Tab 1: Critical Items
        df_critical = df_all[df_all['Severity'] == 'critical'] if len(df_all) > 0 else pd.DataFrame()
        self._add_sheet(wb, "Critical Items", df_critical, self._critical_fill)
        
        # Tab 2: Major Items
        df_major = df_all[df_all['Severity'] == 'major'] if len(df_all) > 0 else pd.DataFrame()
        self._add_sheet(wb, "Major Items", df_major, self._major_fill)
        
        # Tab 3: All Discrepancies
        self._add_sheet(wb, "All Discrepancies", df_all, self._header_fill)
        
        # Save workbook
        wb.save(output_path)
        logger.info(f"Generated Excel report: {output_path}")
        
        return output_path
    
    def _add_sheet(
        self,
        wb: Workbook,
        sheet_name: str,
        df: pd.DataFrame,
        header_fill: PatternFill
    ) -> None:
        """Add a formatted sheet to workbook."""
        ws = wb.create_sheet(sheet_name)
        
        if df.empty:
            ws.cell(row=1, column=1, value="No data available")
            return
            
        # Write headers
        for col_idx, column in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=str(column))
            cell.font = self._header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            cell.border = self._border
            
        # Write data
        for row_idx, row in enumerate(df.values, 2):
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self._border
                
        # Auto-adjust column widths
        for col_idx, column in enumerate(df.columns, 1):
            max_length = max(
                len(str(column)),
                df[column].astype(str).apply(len).max() if len(df) > 0 else 0
            )
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_length + 2, 50)
            
    def generate_evidence_pack(
        self,
        staff_id: str,
        staff_name: str,
        partner: str,
        hr_record: Optional[Dict] = None,
        tfr_record: Optional[Dict] = None,
        attendance_record: Optional[Dict] = None,
        discrepancy: Optional[Any] = None,
        filename: Optional[str] = None
    ) -> Path:
        """Generate PDF evidence pack."""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d")
            safe_name = "".join(c for c in staff_name if c.isalnum() or c in ' -_').strip()[:30]
            filename = f"evidence_au_{staff_id}_{safe_name}_{timestamp}.pdf"
            
        output_path = self.config.output_dir / filename
        
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            rightMargin=20*mm,
            leftMargin=20*mm,
            topMargin=20*mm,
            bottomMargin=20*mm
        )
        
        styles = getSampleStyleSheet()
        elements = []
        
        # Title
        elements.append(Paragraph(f"Evidence Pack: {staff_name}", styles['Title']))
        elements.append(Paragraph(
            f"Staff ID: {staff_id} | Partner: {partner} | Generated: {datetime.now().strftime('%d-%b-%Y')}",
            styles['Normal']
        ))
        elements.append(Spacer(1, 12))
        
        # Discrepancy Details
        if discrepancy:
            elements.append(Paragraph("Discrepancy Details", styles['Heading2']))
            disc_data = [
                ["Field", "Value"],
                ["ID", discrepancy.id],
                ["Type", discrepancy.type.value],
                ["Description", discrepancy.description],
            ]
            if discrepancy.amount_hkd:
                disc_data.append(["Amount (AUD)", f"${discrepancy.amount_hkd:,.2f}"])
                
            t = Table(disc_data, colWidths=[60*mm, 110*mm])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#003366')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ]))
            elements.append(t)
            elements.append(Spacer(1, 20))
            
        # Disclaimer
        elements.append(Paragraph(
            "Generated by JV Cost Control System (Australian Operations)",
            styles['Italic']
        ))
        
        doc.build(elements)
        logger.info(f"Generated evidence pack: {output_path}")
        return output_path
    
    def generate_compliance_report(
        self,
        visa_issues: List[Any],
        super_issues: List[Any],
        filename: Optional[str] = None
    ) -> Path:
        """Generate Fair Work / Compliance report."""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d")
            filename = f"compliance_report_au_{timestamp}.pdf"
            
        output_path = self.config.output_dir / filename
        
        doc = SimpleDocTemplate(str(output_path), pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []
        
        elements.append(Paragraph("Compliance Report: Visa & Superannuation", styles['Title']))
        elements.append(Spacer(1, 12))
        
        # Visa Issues
        if visa_issues:
            elements.append(Paragraph("Visa / Work Rights Issues", styles['Heading2']))
            data = [["Staff", "Issue"]]
            for i in visa_issues:
                data.append([i.staff_name, i.message[:60]])
            
            t = Table(data, colWidths=[50*mm, 120*mm])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D30026')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ]))
            elements.append(t)
            elements.append(Spacer(1, 12))
            
        # Super Issues
        if super_issues:
            elements.append(Paragraph("Superannuation Issues", styles['Heading2']))
            data = [["Staff", "Issue"]]
            for i in super_issues:
                data.append([i.staff_name, i.message[:60]])
            
            t = Table(data, colWidths=[50*mm, 120*mm])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F37021')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ]))
            elements.append(t)
            
        doc.build(elements)
        logger.info(f"Generated compliance report: {output_path}")
        return output_path

    def export_reconciled_data(self, df: pd.DataFrame, filename: Optional[str] = None) -> Path:
        """Export reconciled data."""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d")
            filename = f"reconciled_data_au_{timestamp}.csv"
            
        output_path = self.config.output_dir / filename
        df.to_csv(output_path, index=False)
        return output_path
