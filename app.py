"""
JV Cost Reconciliation System - Streamlit Application

Main entry point for the cost reconciliation dashboard.
Run with: streamlit run app.py

Tabs:
1. Data Ingestion - Upload and preview partner files
2. Discrepancy Dashboard - View and resolve discrepancies
3. Audit Trail - View snapshots and change history
4. Compliance Check - Fair Work / Visa & Super compliance
5. Forecast Export - Download reconciled data
"""

import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional, List, Any
import json

import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# Local imports
from core.ingestion import DataIngestion
from core.validation import DataValidator, ValidationSeverity
from core.reconciliation import ReconciliationEngine, DiscrepancySeverity
from core.audit import AuditTrail
from core.reporting import ReportGenerator, ReportConfig

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ============================================================================
# PAGE CONFIG
# ============================================================================

st.set_page_config(
    page_title="JV Cost Reconciliation (BYCA/JHG)",
    page_icon="🇦🇺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# PATHS
# ============================================================================

BASE_PATH = Path(__file__).parent
CONFIG_PATH = BASE_PATH / "config"
DATA_PATH = BASE_PATH / "data"
OUTPUT_PATH = BASE_PATH / "outputs"

# Ensure directories exist
(DATA_PATH / "raw" / "byca").mkdir(parents=True, exist_ok=True)
(DATA_PATH / "raw" / "jhg").mkdir(parents=True, exist_ok=True)
(DATA_PATH / "curated").mkdir(parents=True, exist_ok=True)
(DATA_PATH / "archive").mkdir(parents=True, exist_ok=True)
OUTPUT_PATH.mkdir(parents=True, exist_ok=True)

# ============================================================================
# SESSION STATE INITIALIZATION
# ============================================================================

if 'ingested_data' not in st.session_state:
    st.session_state.ingested_data = {}  # {partner: {source_type: df}}

if 'discrepancies' not in st.session_state:
    st.session_state.discrepancies = []

if 'validation_results' not in st.session_state:
    st.session_state.validation_results = {}

if 'reconciliation_complete' not in st.session_state:
    st.session_state.reconciliation_complete = False

if 'reconciled_data' not in st.session_state:
    st.session_state.reconciled_data = None

# ============================================================================
# INITIALIZE SERVICES
# ============================================================================

@st.cache_resource
def get_ingestion_engine():
    """Get cached ingestion engine."""
    return DataIngestion(CONFIG_PATH / "partners.yaml")

@st.cache_resource
def get_validator():
    """Get cached validator."""
    return DataValidator(CONFIG_PATH / "rules.yaml")

@st.cache_resource
def get_reconciliation_engine():
    """Get cached reconciliation engine."""
    return ReconciliationEngine(CONFIG_PATH / "rules.yaml")

@st.cache_resource
def get_audit_trail():
    """Get cached audit trail."""
    return AuditTrail(
        archive_path=DATA_PATH / "archive",
        rules_path=CONFIG_PATH / "rules.yaml"
    )

@st.cache_resource
def get_report_generator():
    """Get cached report generator."""
    return ReportGenerator(ReportConfig(
        company_name="JV Cost Control (BYCA/JHG)",
        project_name="Australian Infrastructure Project",
        output_dir=OUTPUT_PATH
    ))

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def save_uploaded_file(uploaded_file, partner: str) -> Path:
    """Save uploaded file to raw data directory."""
    partner_path = DATA_PATH / "raw" / partner.lower()
    file_path = partner_path / uploaded_file.name
    
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    return file_path

def discrepancy_to_dict(d) -> dict:
    """Convert discrepancy object to dict for AgGrid."""
    return {
        'ID': d.id,
        'Partner': d.partner,
        'Staff ID': d.staff_id,
        'Staff Name': d.staff_name,
        'Type': d.type.value.replace('_', ' ').title(),
        'Severity': d.severity.value.upper(),
        'Confidence %': f"{d.confidence:.1f}" if d.confidence else "N/A",
        'Amount (AUD)': f"${d.amount_hkd:,.0f}" if d.amount_hkd else "N/A",  # reusing field name internally but label as AUD
        'Description': d.description[:80] + "..." if len(d.description) > 80 else d.description,
        'Status': d.status,
        '_full_description': d.description,
        '_discrepancy_obj': d
    }

# ============================================================================
# CSS STYLING
# ============================================================================

st.markdown("""
<style>
    .stMetricValue {
        font-size: 2.5rem !important;
    }
    .critical-card {
        background-color: #D30026; /* BYCA Red */
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .major-card {
        background-color: #F37021; /* Orange */
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .minor-card {
        background-color: #FFD700;
        padding: 20px;
        border-radius: 10px;
        color: black;
        text-align: center;
    }
    .success-card {
        background-color: #008000; /* Green */
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    div[data-testid="stMetricValue"] {
        font-size: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# SIDEBAR
# ============================================================================

with st.sidebar:
    st.title("🇦🇺 JV Cost Control")
    st.caption("BYCA / JHG Reconciliation System")
    st.markdown("---")
    
    st.subheader("📊 System Status")
    
    # Data status
    data_loaded = bool(st.session_state.ingested_data)
    recon_done = st.session_state.reconciliation_complete
    
    st.markdown(f"**Data Loaded:** {'✅' if data_loaded else '❌'}")
    st.markdown(f"**Reconciliation:** {'✅' if recon_done else '❌'}")
    
    if data_loaded:
        for partner, sources in st.session_state.ingested_data.items():
            with st.expander(f"📁 {partner.upper()}"):
                for source, df in sources.items():
                    if df is not None:
                        st.markdown(f"• {source}: {len(df)} rows")
    
    st.markdown("---")
    
    # Quick stats if reconciliation done
    if recon_done and st.session_state.discrepancies:
        st.subheader("⚠️ Discrepancy Summary")
        discs = st.session_state.discrepancies
        critical = len([d for d in discs if d.severity == DiscrepancySeverity.CRITICAL])
        major = len([d for d in discs if d.severity == DiscrepancySeverity.MAJOR])
        minor = len([d for d in discs if d.severity == DiscrepancySeverity.MINOR])
        
        st.metric("Critical", critical, delta_color="inverse")
        st.metric("Major", major)
        st.metric("Minor", minor)
    
    st.markdown("---")
    st.caption(f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ============================================================================
# MAIN CONTENT - TABS
# ============================================================================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 Data Ingestion",
    "🔍 Discrepancy Dashboard",
    "📜 Audit Trail",
    "✅ Compliance (Visa/Super)",
    "📤 Forecast Export"
])

# ============================================================================
# TAB 1: DATA INGESTION
# ============================================================================

with tab1:
    st.header("📥 Data Ingestion")
    st.markdown("Upload partner files for BYCA and JHG reconciliation.")
    
    col1, col2 = st.columns(2)
    
    # BYCA upload
    with col1:
        st.subheader("🏗️ BYCA (Partner A)")
        
        byca_hr = st.file_uploader("HR Master (CSV)", type=['csv'], key="byca_hr")
        byca_tfr = st.file_uploader("TFR Timesheet (PDF)", type=['pdf'], key="byca_tfr")
        byca_att = st.file_uploader("Site Attendance (Excel)", type=['xlsx', 'xls'], key="byca_att")
        
        if 'byca' in st.session_state.ingested_data:
            with st.expander("📊 Preview BYCA Data"):
                for source, df in st.session_state.ingested_data['byca'].items():
                    if df is not None:
                        st.markdown(f"**{source}** ({len(df)} rows)")
                        st.dataframe(df.head(5), use_container_width=True)
    
    # JHG upload
    with col2:
        st.subheader("🏗️ JHG (Partner B)")
        
        jhg_hr = st.file_uploader("HR Master (CSV)", type=['csv'], key="jhg_hr")
        jhg_tfr = st.file_uploader("TFR Timesheet (PDF)", type=['pdf'], key="jhg_tfr")
        jhg_att = st.file_uploader("Site Attendance (Excel)", type=['xlsx', 'xls'], key="jhg_att")
        
        if 'jhg' in st.session_state.ingested_data:
            with st.expander("📊 Preview JHG Data"):
                for source, df in st.session_state.ingested_data['jhg'].items():
                    if df is not None:
                        st.markdown(f"**{source}** ({len(df)} rows)")
                        st.dataframe(df.head(5), use_container_width=True)
    
    st.markdown("---")
    
    # Buttons
    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
    
    with col_btn1:
        if st.button("📂 Load Files", type="primary", use_container_width=True):
            ingestion = get_ingestion_engine()
            with st.spinner("Processing files..."):
                # BYCA
                if 'byca' not in st.session_state.ingested_data: st.session_state.ingested_data['byca'] = {}
                if byca_hr:
                    res = ingestion.load_csv(save_uploaded_file(byca_hr, "byca"), "byca", "hr_master")
                    if res.success: st.session_state.ingested_data['byca']['hr_master'] = res.data
                if byca_tfr:
                    res = ingestion.load_pdf(save_uploaded_file(byca_tfr, "byca"), "byca", "tfr")
                    if res.success: st.session_state.ingested_data['byca']['tfr'] = res.data
                if byca_att:
                    res = ingestion.load_excel(save_uploaded_file(byca_att, "byca"), "byca", "attendance")
                    if res.success: st.session_state.ingested_data['byca']['attendance'] = res.data
                
                # JHG
                if 'jhg' not in st.session_state.ingested_data: st.session_state.ingested_data['jhg'] = {}
                if jhg_hr:
                    res = ingestion.load_csv(save_uploaded_file(jhg_hr, "jhg"), "jhg", "hr_master")
                    if res.success: st.session_state.ingested_data['jhg']['hr_master'] = res.data
                if jhg_tfr:
                    res = ingestion.load_pdf(save_uploaded_file(jhg_tfr, "jhg"), "jhg", "tfr")
                    if res.success: st.session_state.ingested_data['jhg']['tfr'] = res.data
                if jhg_att:
                    res = ingestion.load_excel(save_uploaded_file(jhg_att, "jhg"), "jhg", "attendance")
                    if res.success: st.session_state.ingested_data['jhg']['attendance'] = res.data
                
                st.success("✅ Files processed")
    
    with col_btn2:
        if st.button("🔄 Run Reconciliation", type="secondary", use_container_width=True):
            if not st.session_state.ingested_data:
                st.warning("⚠️ Load files first")
            else:
                with st.spinner("Reconciling..."):
                    engine = get_reconciliation_engine()
                    validator = get_validator()
                    audit = get_audit_trail()
                    
                    # Validate
                    for p, sources in st.session_state.ingested_data.items():
                        for t, df in sources.items():
                            if df is not None:
                                res = validator.validate_all(df, t, p)
                                st.session_state.validation_results[f"{p}_{t}"] = res
                                audit.create_snapshot(df, p, t, f"Pre-recon {p} {t}")
                    
                    # Reconcile
                    res = engine.reconcile_all(st.session_state.ingested_data)
                    if res.success:
                        st.session_state.discrepancies = res.discrepancies
                        st.session_state.reconciliation_complete = True
                        st.session_state.reconciled_data = res.reconciled_data
                        audit.log_change("reconciliation", "system", "full_recon", f"Found {len(res.discrepancies)} issues")
                        st.success(f"✅ Found {len(res.discrepancies)} discrepancies")
                    else:
                        st.error("❌ Reconciliation failed")

    with col_btn3:
        if st.button("📁 Load Pre-Generated Data", use_container_width=True):
            ingestion = get_ingestion_engine()
            with st.spinner("Loading dummy data..."):
                # BYCA
                p = DATA_PATH / "raw" / "byca"
                if (p / "hr_master.csv").exists():
                    if 'byca' not in st.session_state.ingested_data: st.session_state.ingested_data['byca'] = {}
                    st.session_state.ingested_data['byca']['hr_master'] = ingestion.load_csv(p/"hr_master.csv", "byca", "hr_master").data
                    if (p/"tfr_feb2026.pdf").exists():
                        st.session_state.ingested_data['byca']['tfr'] = ingestion.load_pdf(p/"tfr_feb2026.pdf", "byca", "tfr").data
                    if (p/"site_attendance.xlsx").exists():
                        st.session_state.ingested_data['byca']['attendance'] = ingestion.load_excel(p/"site_attendance.xlsx", "byca", "attendance").data
                
                # JHG
                p = DATA_PATH / "raw" / "jhg"
                if (p / "hr_master.csv").exists():
                    if 'jhg' not in st.session_state.ingested_data: st.session_state.ingested_data['jhg'] = {}
                    st.session_state.ingested_data['jhg']['hr_master'] = ingestion.load_csv(p/"hr_master.csv", "jhg", "hr_master").data
                    if (p/"tfr_feb2026.pdf").exists():
                        st.session_state.ingested_data['jhg']['tfr'] = ingestion.load_pdf(p/"tfr_feb2026.pdf", "jhg", "tfr").data
                    if (p/"site_attendance.xlsx").exists():
                        st.session_state.ingested_data['jhg']['attendance'] = ingestion.load_excel(p/"site_attendance.xlsx", "jhg", "attendance").data
                
                st.success("✅ Loaded pre-generated Australian data!")
                st.rerun()

# ============================================================================
# TAB 2: DISCREPANCY DASHBOARD
# ============================================================================

with tab2:
    st.header("🔍 Discrepancy Dashboard")
    
    if not st.session_state.reconciliation_complete:
        st.info("ℹ️ Run reconciliation first.")
    else:
        discs = st.session_state.discrepancies
        col1, col2, col3, col4 = st.columns(4)
        
        crit = len([d for d in discs if d.severity == DiscrepancySeverity.CRITICAL])
        maj = len([d for d in discs if d.severity == DiscrepancySeverity.MAJOR])
        min_ = len([d for d in discs if d.severity == DiscrepancySeverity.MINOR])
        res = len([d for d in discs if d.status == "Resolved"])
        
        col1.markdown(f'<div class="critical-card"><h3>🚨 Critical</h3><h1>{crit}</h1></div>', unsafe_allow_html=True)
        col2.markdown(f'<div class="major-card"><h3>⚠️ Major</h3><h1>{maj}</h1></div>', unsafe_allow_html=True)
        col3.markdown(f'<div class="minor-card"><h3>📋 Minor</h3><h1>{min_}</h1></div>', unsafe_allow_html=True)
        col4.markdown(f'<div class="success-card"><h3>✅ Resolved</h3><h1>{res}</h1></div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Grid
        if discs:
            df_disp = pd.DataFrame([discrepancy_to_dict(d) for d in discs]).drop(columns=['_full_description', '_discrepancy_obj'])
            gb = GridOptionsBuilder.from_dataframe(df_disp)
            gb.configure_selection('multiple', use_checkbox=True)
            gb.configure_column("ID", pinned='left', width=130)
            gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
            grid = AgGrid(df_disp, gridOptions=gb.build(), update_mode=GridUpdateMode.SELECTION_CHANGED, height=400)
            
            sel = grid['selected_rows']
            
            c1, c2, c3 = st.columns(3)
            if c1.button("✅ Resolve Selected"):
                if sel is not None and len(sel) > 0:
                    ids = [r['ID'] for r in sel.to_dict('records')]
                    for d in st.session_state.discrepancies:
                        if d.id in ids: d.status = "Resolved"
                    st.success("Resolved!")
                    st.rerun()
            
            if c2.button("📄 Evidence Pack"):
                if sel is not None and len(sel) > 0:
                    d_id = sel.to_dict('records')[0]['ID']
                    d_obj = next((d for d in discs if d.id == d_id), None)
                    path = get_report_generator().generate_evidence_pack(d_obj.staff_id, d_obj.staff_name, d_obj.partner, discrepancy=d_obj)
                    with open(path, "rb") as f:
                        st.download_button("📥 Download PDF", f, file_name=path.name)

            if c3.button("📊 Export Report"):
                path = get_report_generator().generate_excel_report(discs)
                with open(path, "rb") as f:
                    st.download_button("📥 Download Excel", f, file_name=path.name)

# ============================================================================
# TAB 3: AUDIT TRAIL
# ============================================================================

with tab3:
    st.header("📜 Audit Trail")
    audit = get_audit_trail()
    c1, c2 = st.columns([2, 1])
    
    with c1:
        st.subheader("📝 Change Log")
        for e in audit.get_change_log(limit=50):
            with st.expander(f"🕒 {e.timestamp} | {e.action}"):
                st.write(e)
    
    with c2:
        st.subheader("📸 Snapshots")
        for s in audit.list_snapshots(limit=10):
            st.text(f"{s.timestamp} - {s.partner} {s.source_type}")

# ============================================================================
# TAB 4: COMPLIANCE
# ============================================================================

with tab4:
    st.header("✅ Fair Work / Compliance")
    st.markdown("Monitor Visa and Superannuation compliance for Australian standards.")
    
    if st.session_state.validation_results:
        visa_issues = []
        super_issues = []
        
        for k, res in st.session_state.validation_results.items():
            for i in res.issues:
                if 'visa' in i.category.lower(): visa_issues.append(i)
                elif 'super' in i.category.lower(): super_issues.append(i)
        
        c1, c2 = st.columns(2)
        if visa_issues:
            c1.markdown(f'<div class="critical-card"><h3>🛂 Visa Issues</h3><h1>{len(visa_issues)}</h1></div>', unsafe_allow_html=True)
            st.table(pd.DataFrame([{'Staff': i.staff_name, 'Issue': i.message} for i in visa_issues]))
        else:
            c1.markdown('<div class="success-card"><h3>🛂 Visa Status</h3><h1>✓</h1></div>', unsafe_allow_html=True)
            
        if super_issues:
            c2.markdown(f'<div class="major-card"><h3>💰 Super Issues</h3><h1>{len(super_issues)}</h1></div>', unsafe_allow_html=True)
            st.table(pd.DataFrame([{'Staff': i.staff_name, 'Issue': i.message} for i in super_issues]))
        else:
            c2.markdown('<div class="success-card"><h3>💰 Super Status</h3><h1>✓</h1></div>', unsafe_allow_html=True)
                        
        if st.button("📄 Generate Compliance Report"):
            path = get_report_generator().generate_compliance_report(visa_issues, super_issues)
            with open(path, "rb") as f:
                st.download_button("📥 Download Report", f, file_name=path.name)

# ============================================================================
# TAB 5: FORECAST
# ============================================================================

with tab5:
    st.header("📤 Forecast & Data Export")
    if st.session_state.reconciled_data is not None:
        st.dataframe(st.session_state.reconciled_data.head())
        if st.button("Download Reconciled Data"):
            path = get_report_generator().export_reconciled_data(st.session_state.reconciled_data)
            with open(path, "rb") as f:
                st.download_button("📥 Download CSV", f, file_name=path.name)
    else:
        st.info("Run reconciliation first.")
