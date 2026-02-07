# JV Cost Reconciliation System

A production-ready Streamlit application for reconciling staff cost data across Joint Venture partners in Hong Kong infrastructure megaprojects.

## 🚀 Quick Start

### 1. Install Dependencies

```bash
cd cost_control_jv
pip install -r requirements.txt
```

### 2. Generate Dummy Data (Optional - for Testing)

```bash
python utils/dummy_data_generator.py
```

This creates realistic test data with embedded discrepancies in `data/raw/`.

### 3. Run the Application

```bash
streamlit run app.py
```

The application will open at http://localhost:8501

---

## 📁 Project Structure

```
cost_control_jv/
├── app.py                    # Main Streamlit application
├── core/
│   ├── ingestion.py          # PDF/CSV/Excel data loaders
│   ├── validation.py         # Schema & business rule checks
│   ├── reconciliation.py     # 3-way matching + fuzzy logic
│   ├── audit.py              # SHA3-256 snapshots & change log
│   └── reporting.py          # Excel/PDF report generation
├── config/
│   ├── partners.yaml         # Partner column mappings
│   ├── rules.yaml            # Rate bands, tolerances
│   └── contacts.yaml         # Escalation contacts
├── data/
│   ├── raw/                  # Upload partner files here
│   │   ├── dragages/
│   │   └── gammon/
│   ├── curated/              # Processed data
│   └── archive/              # Audit snapshots
├── outputs/                  # Generated reports
└── utils/
    └── dummy_data_generator.py
```

---

## 📥 Adding Real Partner Files

1. Place files in the appropriate folder:
   - `data/raw/dragages/` for Dragages files
   - `data/raw/gammon/` for Gammon files

2. Expected file formats:
   - **HR Master**: CSV with columns: Staff_ID/Employee_No, Name, Role, Daily_Rate
   - **TFR Timesheet**: PDF with tabular data
   - **Site Attendance**: Excel (.xlsx) with clock-in/out times

3. Alternatively, upload files directly via the **Data Ingestion** tab in the UI.

---

## 🔧 Features

### Tab 1: Data Ingestion
- Upload CSV, Excel, PDF files per partner
- Auto-normalize column names across partners
- Preview loaded data with row counts

### Tab 2: Discrepancy Dashboard
- Filter by severity (Critical/Major/Minor)
- Interactive AG-Grid table with multi-select
- Bulk resolve discrepancies
- 1-click evidence pack generation

### Tab 3: Audit Trail
- View all system changes with timestamps
- SHA3-256 hashed snapshots
- Restore previous states

### Tab 4: Compliance Check
- FDH permit validation (foreign workers)
- MPF registration checks
- Labour Dept. compliance report export

### Tab 5: Forecast Export
- Download reconciled data as Excel/CSV
- Ready for P6, SAP, or ERP integration

---

## 🛠️ Troubleshooting

### PDF table not detected
- Ensure tables have clear borders
- Check page margins aren't cutting off content
- Try scanning at higher DPI (300+)

### Column mapping issues
- Verify column names match `config/partners.yaml`
- Check for trailing whitespace in CSV headers
- Use UTF-8 encoding for CSV files

### Fuzzy matching too aggressive/conservative
- Adjust thresholds in `config/rules.yaml`:
  ```yaml
  fuzzy_matching:
    auto_accept_threshold: 95  # Increase for stricter matching
    review_threshold: 85       # Decrease for more auto-accepts
  ```

### Rate band validation failures
- Update rate ranges in `config/rules.yaml` under `rate_bands`

---

## ⚙️ Configuration

### Partner Column Mappings (`config/partners.yaml`)
Maps each partner's column names to normalized names:
```yaml
partners:
  dragages:
    column_mappings:
      hr_master:
        Staff_ID: staff_id
        Staff_Name: name
        # ...
```

### Business Rules (`config/rules.yaml`)
- Fuzzy matching thresholds
- Rate deviation tolerance (default: 15%)
- Rate bands per role
- FDH permit format regex

---

## 📊 Embedded Test Discrepancies

The dummy data generator creates these test cases:

| Type | Count | Severity |
|------|-------|----------|
| Ghost workers | 12 | Critical |
| Name typos | 8 | Major |
| Rate deviations >25% | 5 | Major |
| Missing FDH permits | 3 | Critical |
| Duplicate IDs | 1 | Critical |

---

## 📝 License

Internal use only - JV Cost Control Team

---

## 📞 Support

- Email: cost-control@project.hk
- Internal Wiki: [link]
