import streamlit as st
import pandas as pd
from core.ingestion import load_partner_data
from core.reconciliation import reconcile_data
from core.audit import save_snapshot
from core.reporting import generate_discrepancy_report
import os
from datetime import datetime

st.set_page_config(
    page_title="JV Cost Control System",
    page_icon="🏗️",
    layout="wide"
)

def main():
    st.title("🏗️ JV Construction Cost Control System")
    st.subheader("Australia Infrastructure Megaproject Cost Reconciliation Platform")
    
    # Initialize session state
    if 'reconciled_data' not in st.session_state:
        st.session_state.reconciled_data = None
    if 'discrepancies' not in st.session_state:
        st.session_state.discrepancies = None
    
    # Create tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Data Ingestion", 
        "Discrepancy Dashboard", 
        "Audit Trail", 
        "Compliance Check", 
        "Forecast Export"
    ])
    
    with tab1:
        st.header("Data Ingestion")
        
        # Partner selection
        partner = st.selectbox("Select Partner", ["BYCA", "JHG"])
        source_type = st.selectbox("Select Source Type", ["HR Master", "TFR Timesheet", "Site Attendance"])
        
        uploaded_file = st.file_uploader(
            f"Upload {source_type} file for {partner}", 
            type=["csv", "xlsx", "xls", "pdf"]
        )
        
        if uploaded_file is not None:
            with st.spinner(f"Processing {uploaded_file.name}..."):
                try:
                    df = load_partner_data(uploaded_file, partner.lower(), source_type.lower())
                    st.success(f"Successfully loaded {len(df)} rows from {uploaded_file.name}")
                    st.dataframe(df.head())
                    
                    if st.button("Run Reconciliation"):
                        reconciled_data, discrepancies = reconcile_data()
                        st.session_state.reconciled_data = reconciled_data
                        st.session_state.discrepancies = discrepancies
                        
                        # Save audit snapshot
                        save_snapshot(reconciled_data, discrepancies)
                        
                        st.success(f"Reconciliation completed! Found {len(discrepancies)} discrepancies.")
                        
                except Exception as e:
                    st.error(f"Error processing file: {str(e)}")
    
    with tab2:
        st.header("Discrepancy Dashboard")
        
        if st.session_state.discrepancies is not None:
            df_disc = pd.DataFrame(st.session_state.discrepancies)
            
            # Summary cards
            col1, col2, col3 = st.columns(3)
            with col1:
                critical_count = len(df_disc[df_disc['severity'] == 'Critical'])
                st.metric("Critical Issues", critical_count, delta=None)
            with col2:
                major_count = len(df_disc[df_disc['severity'] == 'Major'])
                st.metric("Major Issues", major_count, delta=None)
            with col3:
                minor_count = len(df_disc[df_disc['severity'] == 'Minor'])
                st.metric("Minor Issues", minor_count, delta=None)
            
            # Display discrepancies table
            st.subheader("Discrepancy Details")
            st.dataframe(df_disc)
            
            # Bulk actions
            if st.button("Export Discrepancy Report"):
                report_path = generate_discrepancy_report(st.session_state.discrepancies)
                with open(report_path, "rb") as file:
                    st.download_button(
                        label="Download Report",
                        data=file,
                        file_name=f"discrepancy_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.info("Run reconciliation in the Data Ingestion tab to see discrepancies here.")
    
    with tab3:
        st.header("Audit Trail")
        
        # Show available snapshots
        archive_dir = "data/archive/"
        if os.path.exists(archive_dir):
            snapshots = os.listdir(archive_dir)
            snapshots.sort(reverse=True)
            st.write("Available Snapshots:")
            for snap in snapshots[:10]:  # Show last 10
                st.write(f"- {snap}")
        else:
            st.info("No snapshots available yet. Run reconciliation to create snapshots.")
    
    with tab4:
        st.header("Compliance Check")
        
        # Placeholder for compliance checks
        st.warning("Compliance checks require actual data. Upload files and run reconciliation first.")
    
    with tab5:
        st.header("Forecast Export")
        
        if st.session_state.reconciled_data is not None:
            st.write("Preview of Reconciled Data:")
            st.dataframe(st.session_state.reconciled_data.head())
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Export as Excel"):
                    excel_path = "outputs/reconciled_data.xlsx"
                    os.makedirs(os.path.dirname(excel_path), exist_ok=True)
                    st.session_state.reconciled_data.to_excel(excel_path, index=False)
                    with open(excel_path, "rb") as file:
                        st.download_button(
                            label="Download Excel",
                            data=file,
                            file_name="reconciled_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            with col2:
                if st.button("Export as CSV"):
                    csv_path = "outputs/reconciled_data.csv"
                    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
                    st.session_state.reconciled_data.to_csv(csv_path, index=False)
                    with open(csv_path, "rb") as file:
                        st.download_button(
                            label="Download CSV",
                            data=file,
                            file_name="reconciled_data.csv",
                            mime="text/csv"
                        )
        else:
            st.info("Run reconciliation to see export options.")

if __name__ == "__main__":
    main()