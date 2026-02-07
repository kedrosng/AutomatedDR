# JV Construction Cost Control System

A complete, production-ready system for reconciling staff data from multiple partners on large-scale construction projects in Hong Kong. Designed for a solo cost controller managing a HK$10B+ infrastructure project with 2 partners.

## Overview

This system addresses critical pain points in joint venture cost control:
- Manual errors in data reconciliation
- Format mismatches between partner systems
- Lack of audit trail for JV disputes
- Non-compliance with HK labour regulations

## Features

- **3-way reconciliation**: Matches HR master lists, TFR timesheets, and site attendance logs
- **Fuzzy matching**: Handles name typos and variations using rapidfuzz
- **HK compliance checks**: Validates FDH permits, MPF contributions, and rate bands
- **Audit trail**: SHA3-256 snapshots with complete change tracking
- **Discrepancy dashboard**: Visualizes critical, major, and minor issues
- **Report generation**: Creates evidence packs and compliance reports

## Prerequisites

- Python 3.10+
- Windows or Linux system

## Setup Instructions

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Generate dummy data** (for testing):
   ```bash
   python utils/dummy_data_generator.py
   ```

3. **Run the application**:
   ```bash
   streamlit run app.py
   ```

## How to Add Real Partner Files

Place your partner files in the appropriate directories under `data/raw/`:

```
data/raw/
├── dragages/
│   ├── hr_master.csv          # HR master list (staff_id, name, role, rate)
│   ├── tfr_feb2026.pdf        # TFR timesheet (staff_id, date, hours)
│   └── site_attendance.xlsx   # Site attendance log (staff_id, date, status)
└── gammon/
    ├── hr_master.csv
    ├── tfr_feb2026.pdf
    └── site_attendance.xlsx
```

### Supported File Formats
- CSV for HR master and attendance data
- Excel (.xlsx, .xls) for attendance data
- PDF for TFR timesheets (tables extracted automatically)

### Expected Columns by Source Type

**HR Master**:
- `staff_id` or `employee_no` (unique identifier)
- `name` (full name)
- `role` (job position)
- `rate` (daily/hourly wage)
- `nationality` (for compliance checks)
- `fdh_permit` (for foreign workers)

**TFR Timesheet**:
- `staff_id` (identifier)
- `date` (work date)
- `hours` (hours worked)
- `rate` (applicable rate)

**Site Attendance**:
- `staff_id` (identifier)
- `date` (attendance date)
- `status` (present, absent, leave, etc.)

## How to Resolve Common Errors

### PDF Table Not Detected
- Check if the PDF has proper table structure (lines separating cells)
- Verify page margins aren't cutting off table edges
- Try converting to image-based PDF if text extraction fails

### Column Name Mismatches
- The system automatically normalizes common variations:
  - Staff ID: `staff_id`, `employee_no`, `emp_id`, `Staff_ID`, etc.
  - Name: `name`, `full_name`, `Name`, `Full_Name`, etc.
- Update `config/partners.yaml` for custom mappings

### Performance Issues with Large Files
- Files with >5000 rows may take longer to process
- Consider splitting large files into smaller chunks
- Ensure adequate RAM (4GB+ recommended)

## System Architecture

### Core Modules
- `core/ingestion.py`: Handles file loading and preprocessing
- `core/validation.py`: Enforces schema and business rules
- `core/reconciliation.py`: Performs 3-way matching with fuzzy logic
- `core/audit.py`: Manages snapshots and change tracking
- `core/reporting.py`: Generates Excel/PDF reports

### Configuration Files
- `config/partners.yaml`: Partner metadata and file patterns
- `config/rules.yaml`: Rate bands, tolerances, compliance rules
- `config/contacts.yaml`: Escalation contacts by discrepancy type

### Data Flow
1. **Ingestion**: Load files from `data/raw/` with automatic format detection
2. **Validation**: Check schemas, mandatory fields, and compliance rules
3. **Reconciliation**: Match across sources with fuzzy logic for typos
4. **Audit**: Create SHA3-256 snapshots of all processing steps
5. **Reporting**: Generate discrepancy reports and evidence packs

## Security & Privacy

- **Zero cloud dependencies**: All processing happens locally
- **Data isolation**: No external API calls or data transmission
- **Complete audit trail**: Every change tracked with cryptographic hashes

## Troubleshooting

If you encounter issues:

1. Check that all required columns are present in your files
2. Verify file encodings (UTF-8 recommended)
3. Confirm that partner names match exactly: "Dragages" and "Gammon"
4. Review the console output for specific error messages

For additional support, contact your internal development team.