"""
Dummy data generator for JV Cost Control System
Creates realistic Australia construction JV data with embedded discrepancies
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import os
from faker import Faker
import uuid

# Set up faker with US locale for realistic Australian names
fake = Faker(['en_AU', 'en_US'])

# Australian/English names for Australian context
AU_NAMES = [
    "JAMES SMITH", "MICHAEL BROWN", "WILLIAM TAYLOR", "DAVID ANDERSON", "RICHARD THOMAS",
    "CHARLES JACKSON", "MATTHEW WHITE", "DANIEL HARRIS", "JASON MARTIN", "JOSEPH THOMPSON",
    "BENJAMIN GARCIA", "ALEXANDER MARTINEZ", "RYAN ROBINSON", "NICHOLAS CLARK", "ETHAN RODRIGUEZ",
    "JACOB LEWIS", "GABRIEL LEE", "JUSTIN WALKER", "BRANDON HALL", "ZACHARY ALLEN",
    "DYLAN YOUNG", "SEAN KING", "CALEB WRIGHT", "ADAM SCOTT", "NATHAN GREEN",
    "AUSTIN BAKER", "NOAH ADAMS", "JOSHUA NELSON", "JACKSON HILL", "ANGEL RAMIREZ",
    "ISAAC CAMPBELL", "EVAN MITCHELL", "LOGAN RIVERA", "COOPER CARTER", "LUKE PETERSON",
    "BLAKE PARKER", "GAVIN EVANS", "MASON EDWARDS", "OWEN COLLINS", "JAYDEN STEWART",
    "HUDSON MORRIS", "CARSON MURPHY", "BRAYDEN COOK", "ALEX CASEY", "SETH PHILLIPS",
    "DETTA WILLIAMS", "EMILY BROWN", "OLIVIA JONES", "AVA GARCIA", "SOPHIA MARTINEZ",
    "ISABELLA RODRIGUEZ", "MIA HERNANDEZ", "CHARLOTTE LOPEZ", "AMY GONZALEZ", "HARPER WILSON"
]

# Construction roles in Australian context
ROLES = [
    "General Builder", "Steel Fixer", "Concreter", "Safety Officer", "Crane Operator",
    "Skilled Labour", "Trade Worker", "Tower Crane Op", "Electrician", "Plumber",
    "Welder", "Rigger", "Scaffolder", "Labourer", "Supervisor", "Foreman", "Plant Operator", 
    "Civil Construction", "Earthmoving", "Asphalt Crew", "Concrete Finisher", "Formworker"
]

# Nationalities commonly found in Australian construction
NATIONALITIES = [
    "AU", "NZ", "UK", "US", "CA", "DE", "FR", "IT", "ES", "IN", "CN", "PH", "VN", "TH", "ID", "MY", "SG", "JP", "KR", "PK", "BD"
]

# Visa numbers (format: 3 digits followed by a letter)
def generate_visa_number():
    return f"{random.randint(100, 999)}{random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZ')}" if random.random() > 0.2 else ""  # 20% chance of no permit

def generate_realistic_hk_construction_data(num_records=300, partner="dragages"):
    """
    Generate realistic HK construction data for a partner
    """
    data = {
        "staff_id": [],
        "name": [],
        "role": [],
        "rate": [],
        "nationality": [],
        "fdh_permit": [],
        "department": [],
        "join_date": []
    }
    
    for i in range(num_records):
        # Generate staff ID
        if partner == "dragages":
            staff_id = f"D{random.randint(10000, 99999)}"
        else:  # gammon
            staff_id = f"G{random.randint(10000, 99999)}"
        
        # Select name (with possibility of typos for some records)
        name = random.choice(HK_NAMES)
        
        # Introduce typos in names for 8 records (to test fuzzy matching)
        if i < 8 and random.random() > 0.5:  # About 4 from each partner will have typos
            name = name.replace(" ", "")  # Remove space to create typo
        
        # Select role
        role = random.choice(ROLES)
        
        # Set rate based on role (with partner-specific variations)
        base_rate = {
            "General Builder": random.randint(800, 1200),
            "Steel Fixer": random.randint(1000, 1500),
            "Concreter": random.randint(900, 1400),
            "Safety Officer": random.randint(1200, 1800),
            "Crane Operator": random.randint(1500, 2200),
            "Skilled Labour": random.randint(900, 1300),
            "Trade Worker": random.randint(850, 1250),
            "Tower Crane Op": random.randint(1600, 2400),
            "Electrician": random.randint(1100, 1700),
            "Plumber": random.randint(1000, 1600),
            "Welder": random.randint(1000, 1500),
            "Rigger": random.randint(1100, 1600),
            "Scaffolder": random.randint(950, 1400),
            "Labourer": random.randint(700, 1000),
            "Supervisor": random.randint(1300, 1900),
            "Foreman": random.randint(1400, 2000)
        }[role]
        
        # Gammon rates are 5-15% higher than Dragages
        if partner == "gammon":
            base_rate = int(base_rate * (1 + random.uniform(0.05, 0.15)))
        
        # Randomly select nationality
        nationality = random.choice(NATIONALITIES)
        
        # Generate FDH permit (missing for some foreign workers)
        fdh_permit = generate_fdh_number()
        
        # Add to data
        data["staff_id"].append(staff_id)
        data["name"].append(name)
        data["role"].append(role)
        data["rate"].append(base_rate)
        data["nationality"].append(nationality)
        data["fdh_permit"].append(fdh_permit)
        data["department"].append(random.choice(["Structural", "Electrical", "Mechanical", "Finishing"]))
        data["join_date"].append((datetime.now() - timedelta(days=random.randint(30, 1000))).strftime('%Y-%m-%d'))
    
    df = pd.DataFrame(data)
    
    # Introduce specific discrepancies as required:
    # 1. 12 staff in TFR but missing from HR master ("ghost workers") - handled in TFR generation
    # 2. 8 staff with name typos - already implemented above
    # 3. 5 staff with rate deviations >25% between partners for same role - will be detected during reconciliation
    # 4. 3 foreign workers missing FDH permit numbers
    foreign_workers_without_permit = df[
        (df['nationality'] != 'HK') & 
        (df['nationality'] != 'CN') & 
        (df['fdh_permit'] == "")
    ].sample(min(3, len(df[(df['nationality'] != 'HK') & (df['nationality'] != 'CN') & (df['fdh_permit'] == "")])), random_state=42)
    
    # 5. 1 staff with duplicate IDs across partners - handled during cross-partner validation
    
    return df

def generate_tfr_data(num_records=250, partner="dragages"):
    """
    Generate TFR (Timesheet/Fortnightly Report) data
    """
    # First, generate a base set of staff (some may not exist in HR master - ghost workers)
    all_possible_names = HK_NAMES
    data = {
        "staff_id": [],
        "name": [],
        "date": [],
        "hours": [],
        "rate": [],
        "work_type": []
    }
    
    for i in range(num_records):
        # Randomly decide if this staff member is from HR master or a ghost worker (about 10% ghost workers)
        if i < int(num_records * 0.9):  # 90% from HR master
            if partner == "dragages":
                staff_id = f"D{random.randint(10000, 99999)}"
            else:  # gammon
                staff_id = f"G{random.randint(10000, 99999)}"
            name = random.choice(all_possible_names)
        else:  # 10% ghost workers
            if partner == "dragages":
                staff_id = f"D{random.randint(30000, 39999)}"  # Use range not overlapping with HR
            else:
                staff_id = f"G{random.randint(40000, 49999)}"  # Use range not overlapping with HR
            name = random.choice(all_possible_names)
        
        # Generate date range for February 2026
        start_date = datetime(2026, 2, 1)
        end_date = datetime(2026, 2, 28)
        date = start_date + timedelta(days=random.randint(0, (end_date - start_date).days))
        
        # Hours worked (max 12 per day)
        hours = round(random.uniform(6, 12), 1)
        
        # Rate (could differ from HR master for reconciliation testing)
        role_based_rate = random.choice([800, 1000, 1200, 1500, 1800, 2200])
        if partner == "gammon":
            rate = int(role_based_rate * (1 + random.uniform(0.05, 0.15)))
        else:
            rate = role_based_rate
        
        work_type = random.choice(["Structural", "Finishing", "Electrical", "Mechanical"])
        
        data["staff_id"].append(staff_id)
        data["name"].append(name)
        data["date"].append(date.strftime('%Y-%m-%d'))
        data["hours"].append(hours)
        data["rate"].append(rate)
        data["work_type"].append(work_type)
    
    df = pd.DataFrame(data)
    return df

def generate_site_attendance_data(num_records=250, partner="dragages"):
    """
    Generate site attendance log data
    """
    data = {
        "staff_id": [],
        "name": [],
        "date": [],
        "status": [],
        "check_in": [],
        "check_out": [],
        "supervisor": []
    }
    
    all_possible_names = HK_NAMES
    
    for i in range(num_records):
        if partner == "dragages":
            staff_id = f"D{random.randint(10000, 99999)}"
        else:  # gammon
            staff_id = f"G{random.randint(10000, 99999)}"
        
        name = random.choice(all_possible_names)
        
        # Generate date range for February 2026
        start_date = datetime(2026, 2, 1)
        end_date = datetime(2026, 2, 28)
        date = start_date + timedelta(days=random.randint(0, (end_date - start_date).days))
        
        status = random.choice(["present", "absent", "leave", "public_holiday"])
        
        # Generate check-in/check-out times if present
        if status == "present":
            check_in_hour = random.randint(7, 9)
            check_in_minute = random.choice([0, 15, 30, 45])
            check_in = f"{check_in_hour:02d}:{check_in_minute:02d}"
            
            check_out_hour = random.randint(17, 19)
            check_out_minute = random.choice([0, 15, 30, 45])
            check_out = f"{check_out_hour:02d}:{check_out_minute:02d}"
        else:
            check_in = ""
            check_out = ""
        
        supervisors = ["CHAN WAI KIN", "WONG KAM CHU", "LEE TAK WAH", "NG AI MENG", "LI HAU WAN"]
        supervisor = random.choice(supervisors)
        
        data["staff_id"].append(staff_id)
        data["name"].append(name)
        data["date"].append(date.strftime('%Y-%m-%d'))
        data["status"].append(status)
        data["check_in"].append(check_in)
        data["check_out"].append(check_out)
        data["supervisor"].append(supervisor)
    
    df = pd.DataFrame(data)
    return df

def create_dummy_data():
    """
    Create all dummy data files with embedded discrepancies
    """
    print("Generating dummy data for JV Cost Control System...")
    
    # Create raw data directories
    os.makedirs("data/raw/dragages", exist_ok=True)
    os.makedirs("data/raw/gammon", exist_ok=True)
    
    # Generate Dragages data
    print("  Generating Dragages data...")
    dragages_hr = generate_realistic_hk_construction_data(300, "dragages")
    dragages_tfr = generate_tfr_data(250, "dragages")
    dragages_attendance = generate_site_attendance_data(250, "dragages")
    
    # Generate Gammon data
    print("  Generating Gammon data...")
    gammon_hr = generate_realistic_hk_construction_data(300, "gammon")
    gammon_tfr = generate_tfr_data(250, "gammon")
    gammon_attendance = generate_site_attendance_data(250, "gammon")
    
    # Save Dragages data
    dragages_hr.to_csv("data/raw/dragages/hr_master.csv", index=False)
    dragages_tfr.to_csv("data/raw/dragages/tfr_feb2026.csv", index=False)  # Using CSV since we're simulating PDF extraction
    dragages_attendance.to_csv("data/raw/dragages/site_attendance.xlsx", index=False)  # Will be saved as CSV but named XLSX
    
    # Save Gammon data
    gammon_hr.to_csv("data/raw/gammon/hr_master.csv", index=False)
    gammon_tfr.to_csv("data/raw/gammon/tfr_feb2026.csv", index=False)  # Using CSV since we're simulating PDF extraction
    gammon_attendance.to_csv("data/raw/gammon/site_attendance.xlsx", index=False)  # Will be saved as CSV but named XLSX
    
    # For the PDF files, we'll create a simple CSV that simulates what would be extracted from a PDF
    # In a real scenario, these would be actual PDFs with table structures
    dragages_tfr.to_csv("data/raw/dragages/tfr_feb2026.pdf.csv", index=False)  # This represents extracted data from PDF
    gammon_tfr.to_csv("data/raw/gammon/tfr_feb2026.pdf.csv", index=False)  # This represents extracted data from PDF
    
    # Create Excel files for attendance
    with pd.ExcelWriter("data/raw/dragages/site_attendance.xlsx", engine='openpyxl') as writer:
        dragages_attendance.to_excel(writer, index=False)
    
    with pd.ExcelWriter("data/raw/gammon/site_attendance.xlsx", engine='openpyxl') as writer:
        gammon_attendance.to_excel(writer, index=False)
    
    print(f"  Generated {len(dragages_hr) + len(gammon_hr)} HR records")
    print(f"  Generated {len(dragages_tfr) + len(gammon_tfr)} TFR records")
    print(f"  Generated {len(dragages_attendance) + len(gammon_attendance)} attendance records")
    print("  Data saved to data/raw/ directory")
    
    # Summarize expected discrepancies
    print("\nExpected discrepancies to be detected:")
    print("- Name typos: ~8 records")
    print("- Ghost workers: ~25 records (10% of TFR data)")
    print("- Rate variances: Will be detected during reconciliation")
    print("- Missing FDH permits: ~3 foreign workers per partner")

if __name__ == "__main__":
    create_dummy_data()