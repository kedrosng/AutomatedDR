"""
Dummy Data Generator for JV Cost Reconciliation Testing (Australia)

Generates realistic Australian construction JV data with:
- 3000 HR master records (1500 per partner)
- ~20000 site attendance records (~10000 per partner)
- 2 TFR PDF files with table layouts

Includes embedded discrepancies for testing:
- 12 ghost workers (in TFR but not HR)
- 8 name typos for fuzzy matching
- 5 rate deviations >25%
- 3 missing Visa details
- 1 duplicate ID across partners
"""

import random
import string
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Tuple
import logging

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ============================================================================
# REALISTIC AUSTRALIAN CONSTRUCTION DATA
# ============================================================================

SURNAMES = [
    "SMITH", "JONES", "WILLIAMS", "BROWN", "WILSON", "TAYLOR", "JOHNSON",
    "WHITE", "MARTIN", "ANDERSON", "THOMPSON", "NGUYEN", "LEE", "WANG",
    "SINGH", "PATEL", "KELLY", "O'BRIEN", "RYAN", "WALKER"
]

GIVEN_NAMES = [
    "JAMES", "JOHN", "ROBERT", "MICHAEL", "WILLIAM", "DAVID", "RICHARD",
    "THOMAS", "CHARLES", "DANIEL", "MATTHEW", "ANTHONY", "MARK", "DONALD",
    "STEVEN", "PAUL", "ANDREW", "JOSHUA", "KENNETH", "KEVIN", "BRIAN",
    "JESSICA", "SARAH", "EMILY", "JENNIFER", "LISA", "ELIZABETH"
]

# Roles with AUD daily rate ranges
ROLES_BYCA = {
    "General Labour": (350, 550),
    "Skilled Labour": (450, 650),
    "Rigger": (500, 800),
    "Dogman": (480, 750),
    "Scaffolder": (500, 800),
    "Electrician": (600, 950),
    "Carpenter": (550, 850),
    "Leading Hand": (700, 1100),
    "Foreman": (800, 1200),
    "Site Engineer": (800, 1400),
    "Safety Officer": (600, 1000),
}

# JHG uses slightly different naming/rates
ROLES_JHG = {
    "Labourer": (350, 550),
    "Trade Assistant": (450, 650),
    "Rigger": (520, 820),
    "Dogman": (500, 780),
    "Scaffolder": (520, 820),
    "Electrician": (620, 980),
    "Carpenter": (570, 880),
    "Senior Supervisor": (750, 1150),
    "Site Manager": (900, 1500),
    "Project Engineer": (850, 1450),
    "HSE Advisor": (650, 1050),
}

NATIONALITIES = [
    ("Australia", 0.60, False),
    ("UK", 0.10, True),
    ("Ireland", 0.10, True),
    ("New Zealand", 0.05, False),
    ("Philippines", 0.05, True),
    ("China", 0.03, True),
    ("India", 0.03, True),
    ("France", 0.04, True),
]

SUPER_FUNDS = ["AusSuper", "CBUS", "Hostplus", "Rest", "UniSuper", "ART"]


def generate_staff_id(partner: str, index: int) -> str:
    """Generate staff ID in partner format."""
    if partner == "byca":
        return f"BYCA{index:05d}"
    else:
        return f"JHG{index:05d}"


def generate_name(nationality: str) -> str:
    """Generate realistic name based on nationality."""
    surname = random.choice(SURNAMES)
    given = random.choice(GIVEN_NAMES)
    return f"{surname} {given}"


def generate_visa_status(nationality: str) -> str:
    """Generate Visa status/number."""
    if nationality in ["Australia", "New Zealand"]:
        return "Citizen/PR"
    
    visa_types = ["482 TSS", "417 Working Holiday", "400 Short Stay", "500 Student"]
    visa = random.choice(visa_types)
    grant_number = "".join(random.choices(string.digits, k=13))
    return f"{visa} - {grant_number}"


def generate_super_fund() -> str:
    """Generate Superannuation fund details."""
    fund = random.choice(SUPER_FUNDS)
    member_id = "".join(random.choices(string.digits, k=9))
    return f"{fund} ({member_id})"


def introduce_name_typo(name: str) -> str:
    """Introduce a common typo in a name."""
    typo_types = [
        lambda n: n.replace(" ", ""),               # Remove space
        lambda n: n[:-1] if len(n) > 2 else n,      # Drop last char
        lambda n: n[:1] + n[2:] if len(n) > 2 else n,  # Drop second char
        lambda n: n.replace("S", "5") if "S" in n else n,
        lambda n: n + " " if not n.endswith(" ") else n,  # Trailing space
    ]
    return random.choice(typo_types)(name)


def generate_hr_master(
    partner: str,
    count: int,
    start_id: int,
    roles: dict,
    ghost_workers: List[Tuple[str, str]] = None,
    missing_visa_indices: List[int] = None,
    rate_deviation_indices: List[int] = None
) -> pd.DataFrame:
    """Generate HR master data for a partner."""
    ghost_workers = ghost_workers or []
    missing_visa_indices = missing_visa_indices or []
    rate_deviation_indices = rate_deviation_indices or []
    
    records = []
    
    for i in range(count):
        idx = start_id + i
        staff_id = generate_staff_id(partner, idx)
        
        # Skip if ghost worker
        if any(staff_id == gw[0] for gw in ghost_workers):
            continue
        
        # Random nationality
        nat_choice = random.random()
        cumulative = 0
        nationality = "Australia"
        requires_visa = False
        
        for nat, weight, needs_visa in NATIONALITIES:
            cumulative += weight
            if nat_choice <= cumulative:
                nationality = nat
                requires_visa = needs_visa
                break
        
        name = generate_name(nationality)
        role = random.choice(list(roles.keys()))
        min_rate, max_rate = roles[role]
        daily_rate = round(random.uniform(min_rate, max_rate), 0)
        
        # Rate deviation
        if i in rate_deviation_indices:
            deviation = random.choice([0.30, 0.35, 0.40])
            daily_rate = round(daily_rate * (1 + deviation), 0)
        
        # Visa
        if requires_visa:
            if i in missing_visa_indices:
                visa_status = ""  # Missing for testing
            else:
                visa_status = generate_visa_status(nationality)
        else:
            visa_status = "Citizen/PR"
        
        super_fund = generate_super_fund()
        join_date = datetime.now() - timedelta(days=random.randint(30, 1500))
        
        if partner == "byca":
            records.append({
                "Staff_ID": staff_id,
                "Staff_Name": name,
                "Role": role,
                "Daily_Rate_AUD": daily_rate,
                "Nationality": nationality,
                "Visa_Status": visa_status,
                "Super_Fund": super_fund,
                "Join_Date": join_date.strftime("%Y-%m-%d"),
            })
        else:  # jhg
            records.append({
                "Employee_No": staff_id,
                "Full_Name": name,
                "Position": role,
                "Rate_Per_Day": daily_rate,
                "Country": nationality,
                "Visa_Type": visa_status,
                "Super_Account": super_fund,
                "Start_Date": join_date.strftime("%Y-%m-%d"),
            })
    
    return pd.DataFrame(records)


def generate_tfr_data(
    partner: str,
    hr_df: pd.DataFrame,
    month: str,
    ghost_workers: List[Tuple[str, str]] = None,
    name_typos: List[int] = None,
    rate_deviations: List[int] = None
) -> pd.DataFrame:
    """Generate TFR timesheet data."""
    ghost_workers = ghost_workers or []
    name_typos = name_typos or []
    rate_deviations = rate_deviations or []
    
    records = []
    
    # Get columns
    if partner == "byca":
        id_col, name_col, rate_col = "Staff_ID", "Staff_Name", "Daily_Rate_AUD"
    else:
        id_col, name_col, rate_col = "Employee_No", "Full_Name", "Rate_Per_Day"
    
    # Sample from HR
    sample_size = min(int(len(hr_df) * 0.80), len(hr_df))
    sampled = hr_df.sample(n=sample_size)
    
    for i, (_, row) in enumerate(sampled.iterrows()):
        staff_id = row[id_col]
        name = row[name_col]
        daily_rate = row[rate_col]
        
        if i in name_typos:
            name = introduce_name_typo(name)
        
        if i in rate_deviations:
            deviation = random.choice([0.20, 0.25, 0.30])
            daily_rate = round(daily_rate * (1 + random.choice([-1, 1]) * deviation), 0)
        
        num_days = random.randint(15, 25)
        base_date = datetime(2026, 2, 1)
        
        for day_offset in random.sample(range(1, 29), num_days):
            work_date = base_date + timedelta(days=day_offset)
            hours = round(random.uniform(8, 10), 1)
            ot_hours = round(random.uniform(0, 3), 1) if random.random() > 0.6 else 0
            
            if partner == "byca":
                records.append({
                    "Staff_ID": staff_id,
                    "Name": name,
                    "Date": work_date.strftime("%Y-%m-%d"),
                    "Hours_Worked": hours,
                    "OT_Hours": ot_hours,
                    "Daily_Rate": daily_rate,
                })
            else:
                records.append({
                    "Emp_No": staff_id,
                    "Employee_Name": name,
                    "Work_Date": work_date.strftime("%Y-%m-%d"),
                    "Std_Hours": hours,
                    "Overtime": ot_hours,
                    "Day_Rate": daily_rate,
                })
    
    # Add ghost workers
    for ghost_id, ghost_name in ghost_workers:
        num_days = random.randint(10, 20)
        base_date = datetime(2026, 2, 1)
        daily_rate = round(random.uniform(400, 800), 0)
        
        for day_offset in random.sample(range(1, 29), num_days):
            work_date = base_date + timedelta(days=day_offset)
            hours = round(random.uniform(8, 10), 1)
            
            if partner == "byca":
                records.append({
                    "Staff_ID": ghost_id,
                    "Name": ghost_name,
                    "Date": work_date.strftime("%Y-%m-%d"),
                    "Hours_Worked": hours,
                    "OT_Hours": 0,
                    "Daily_Rate": daily_rate,
                })
            else:
                records.append({
                    "Emp_No": ghost_id,
                    "Employee_Name": ghost_name,
                    "Work_Date": work_date.strftime("%Y-%m-%d"),
                    "Std_Hours": hours,
                    "Overtime": 0,
                    "Day_Rate": daily_rate,
                })
    
    return pd.DataFrame(records)


def generate_attendance_data(
    partner: str,
    hr_df: pd.DataFrame
) -> pd.DataFrame:
    """Generate site attendance data."""
    records = []
    
    if partner == "byca":
        id_col, name_col = "Staff_ID", "Staff_Name"
    else:
        id_col, name_col = "Employee_No", "Full_Name"
    
    sampled = hr_df
    site_zones = ["Zone A - Tunnel", "Zone B - Station Box", "Zone C - Viaduct", "Zone D - Depot"]
    
    for _, row in sampled.iterrows():
        staff_id = row[id_col]
        name = row[name_col]
        
        num_days = random.randint(5, 9)
        base_date = datetime(2026, 2, 1)
        
        for day_offset in random.sample(range(1, 29), num_days):
            work_date = base_date + timedelta(days=day_offset)
            
            clock_in_hour = random.randint(6, 8)
            clock_in_min = random.randint(0, 59)
            clock_out_hour = random.randint(16, 19)
            clock_out_min = random.randint(0, 59)
            
            clock_in = f"{clock_in_hour:02d}:{clock_in_min:02d}"
            clock_out = f"{clock_out_hour:02d}:{clock_out_min:02d}"
            site_zone = random.choice(site_zones)
            
            if partner == "byca":
                records.append({
                    "Staff_ID": staff_id,
                    "Staff_Name": name,
                    "Date": work_date.strftime("%Y-%m-%d"),
                    "Clock_In": clock_in,
                    "Clock_Out": clock_out,
                    "Site_Zone": site_zone,
                })
            else:
                records.append({
                    "Employee_No": staff_id,
                    "Employee_Name": name,
                    "Attendance_Date": work_date.strftime("%Y-%m-%d"),
                    "Entry_Time": clock_in,
                    "Exit_Time": clock_out,
                    "Work_Area": site_zone,
                })
    
    return pd.DataFrame(records)


def generate_tfr_pdf(tfr_df: pd.DataFrame, output_path: Path, partner: str, month: str = "Feb 2026") -> None:
    """Generate a PDF with TFR table."""
    doc = SimpleDocTemplate(str(output_path), pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=20*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    elements = []
    
    title = f"Time For Record (TFR) - {partner.title()} - {month}"
    elements.append(Paragraph(title, styles['Title']))
    elements.append(Spacer(1, 10))
    
    if "Staff_ID" in tfr_df.columns:
        id_col, name_col, hours_col, rate_col = "Staff_ID", "Name", "Hours_Worked", "Daily_Rate"
    else:
        id_col, name_col, hours_col, rate_col = "Emp_No", "Employee_Name", "Std_Hours", "Day_Rate"
    
    agg_df = tfr_df.groupby([id_col, name_col]).agg({hours_col: 'sum', rate_col: 'first'}).reset_index()
    agg_df['Total_Pay'] = agg_df[hours_col] / 8 * agg_df[rate_col]
    
    agg_df = agg_df.head(2000)
    
    header = ["Staff ID", "Name", "Total Hours", "Daily Rate", "Total Pay (AUD)"]
    table_data = [header]
    
    for _, row in agg_df.iterrows():
        table_data.append([
            str(row[id_col]),
            str(row[name_col])[:25],
            f"{row[hours_col]:.1f}",
            f"${row[rate_col]:,.0f}",
            f"${row['Total_Pay']:,.0f}",
        ])
    
    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E2001A' if 'byca' in str(output_path) else '#D30026')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    
    elements.append(table)
    doc.build(elements)
    logger.info(f"Generated TFR PDF: {output_path}")


def generate_all_data(base_path: Path) -> None:
    """Generate all dummy data files."""
    raw_path = base_path / "data" / "raw"
    
    # Create directories
    (raw_path / "byca").mkdir(parents=True, exist_ok=True)
    (raw_path / "jhg").mkdir(parents=True, exist_ok=True)
    
    # Define discrepancies
    ghost_workers_byca = [(f"BYCAX{i:04d}", generate_name("Australia")) for i in range(1, 7)]
    ghost_workers_jhg = [(f"JHGX{i:04d}", generate_name("UK")) for i in range(1, 7)]
    
    name_typo_indices_byca = [5, 15, 45, 78]
    name_typo_indices_jhg = [8, 22, 56, 89]
    
    rate_deviation_indices_byca = [10, 30]
    rate_deviation_indices_jhg = [12, 35, 67]
    
    missing_visa_byca = [20, 80]
    missing_visa_jhg = [25]
    
    duplicate_id = "SHARED001"
    
    logger.info("Generating BYCA HR Master...")
    hr_byca = generate_hr_master(
        partner="byca", count=1501, start_id=1000, roles=ROLES_BYCA,
        ghost_workers=ghost_workers_byca, missing_visa_indices=missing_visa_byca,
        rate_deviation_indices=rate_deviation_indices_byca
    )
    dup_record = hr_byca.iloc[0].copy()
    dup_record["Staff_ID"] = duplicate_id
    hr_byca = pd.concat([hr_byca, pd.DataFrame([dup_record])], ignore_index=True).head(1500)
    hr_byca.to_csv(raw_path / "byca" / "hr_master.csv", index=False)
    
    logger.info("Generating JHG HR Master...")
    hr_jhg = generate_hr_master(
        partner="jhg", count=1501, start_id=3000, roles=ROLES_JHG,
        ghost_workers=ghost_workers_jhg, missing_visa_indices=missing_visa_jhg,
        rate_deviation_indices=rate_deviation_indices_jhg
    )
    dup_record = hr_jhg.iloc[0].copy()
    dup_record["Employee_No"] = duplicate_id
    hr_jhg = pd.concat([hr_jhg, pd.DataFrame([dup_record])], ignore_index=True).head(1500)
    hr_jhg.to_csv(raw_path / "jhg" / "hr_master.csv", index=False)
    
    logger.info("Generating BYCA TFR...")
    tfr_byca = generate_tfr_data(
        partner="byca", hr_df=hr_byca, month="Feb2026",
        ghost_workers=ghost_workers_byca, name_typos=name_typo_indices_byca,
        rate_deviations=rate_deviation_indices_byca
    )
    generate_tfr_pdf(tfr_byca, raw_path / "byca" / "tfr_feb2026.pdf", partner="BYCA", month="February 2026")
    
    logger.info("Generating JHG TFR...")
    tfr_jhg = generate_tfr_data(
        partner="jhg", hr_df=hr_jhg, month="Feb2026",
        ghost_workers=ghost_workers_jhg, name_typos=name_typo_indices_jhg,
        rate_deviations=rate_deviation_indices_jhg
    )
    generate_tfr_pdf(tfr_jhg, raw_path / "jhg" / "tfr_feb2026.pdf", partner="JHG", month="February 2026")
    
    logger.info("Generating BYCA Attendance...")
    att_byca = generate_attendance_data(partner="byca", hr_df=hr_byca)
    att_byca.to_excel(raw_path / "byca" / "site_attendance.xlsx", index=False)
    
    logger.info("Generating JHG Attendance...")
    att_jhg = generate_attendance_data(partner="jhg", hr_df=hr_jhg)
    att_jhg.to_excel(raw_path / "jhg" / "site_attendance.xlsx", index=False)
    
    # Print summary
    print("\n" + "="*60)
    print("DUMMY DATA GENERATION COMPLETE (AUSTRALIALISED)")
    print("="*60)
    print(f"\nFiles created in: {raw_path}")

if __name__ == "__main__":
    project_root = Path(__file__).parent.parent
    generate_all_data(project_root)
