import re
import os
import csv
import glob
import hashlib
from datetime import datetime
from tabulate import tabulate
from collections import namedtuple
import tkinter as tk
from tkinter import filedialog


class Drawing:
    def __init__(self, drawing_number, title, revision, date, location_code, drawing_type, submission_date=None, submission_ref=None):
        self.drawing_number = drawing_number
        self.title = title
        self.revision = revision
        self.date = date
        self.location_code = location_code
        self.drawing_type = drawing_type
        self.submission_date = submission_date
        self.submission_ref = submission_ref

    def __str__(self):
        return (f"{self.drawing_number}: {self.title} "
                f"(Location Code: {self.location_code}, Drawing Type: {self.drawing_type}, "
                f"Revision: {self.revision}, Date: {self.date}, "
                f"Submission Date: {self.submission_date}, Submission Ref: {self.submission_ref})")

def process_all_drawings(base_folder, drawings_register):
    # Prepare the log file
    with open("log.txt", "w") as log_file:
        # Recursively process all folders and subfolders in the base_folder
        for root, dirs, files in os.walk(base_folder):
            for drawing_type_folder in dirs:
                drawing_type_path = os.path.join(root, drawing_type_folder)

                print(f"Processing drawing type folder: {drawing_type_path}")

                submission_date = datetime.fromtimestamp(os.path.getmtime(drawing_type_path)).strftime('%Y-%m-%d')

                # Find all the drawing files in the drawing_type_path
                drawing_files = glob.glob(os.path.join(drawing_type_path, "*"))

                # Filter out the files that do not match the file name format criteria
                invalid_files = []
                valid_files = []
                for drawing_file in drawing_files:
                    if re.search(r"HKT2_BYME_SDWG_([A-Za-z0-9]+)[_-]([A-Za-z0-9]+)[_-](\d+)[_-](\w)", os.path.basename(drawing_file)):
                        valid_files.append(drawing_file)
                        
                    else:
                        invalid_files.append(drawing_file)
                        log_file.write(f"Invalid file format: {drawing_file}\n")
                        

                print(f"Valid files found: {len(valid_files)}")
                print(f"Invalid files found: {len(invalid_files)}")
                if valid_files:
                    update_drawings_in_batch(drawings_register, valid_files, submission_date, drawing_type_folder)
                    
def get_drawing_info_from_file(drawing_file):
    # Extract drawing information from the file's basename
    drawing_info = parse_filename(os.path.basename(drawing_file))

    # If drawing_info is None, return None
    if drawing_info is None:
        return None

    # Extract the title from the file's basename
    drawing_title = os.path.splitext(os.path.basename(drawing_file))[0]

    # Add the title to the drawing_info dictionary
    drawing_info["title"] = drawing_title
    
    return drawing_info

def parse_filename(filename):
    pattern = r"HKT2_BYME_SDWG_([A-Za-z0-9]+)[_-]([A-Za-z0-9]+)[_-](\d+)[_-]([A-Za-z0-9]*)\.pdf"
    match = re.search(pattern, filename)
    if match:
        location_code, drawing_type, drawing_number, revision = match.groups()
        return {
            "location_code": location_code,
            "drawing_type": drawing_type,
            "drawing_number": int(drawing_number),
            "revision": revision
        }
    return None

def update_drawings_in_batch(drawings, filepaths, submission_date, submission_ref):
    updated_drawings = 0
    for filepath in filepaths:
        drawing_info = parse_filename(os.path.basename(filepath))
        if drawing_info is None:
            print(f"Error: could not extract drawing info from file {filepath}")
            continue
        
        # Find the existing drawing with the same drawing number, location code, drawing type, and revision
        existing_drawing = None
        for drawing in drawings:
            if (drawing.drawing_number == drawing_info["drawing_number"]
                    and drawing.location_code == drawing_info["location_code"]
                    and drawing.drawing_type == drawing_info["drawing_type"]
                    and drawing.revision == drawing_info["revision"]):
                existing_drawing = drawing
                break
            
        if existing_drawing:
            existing_drawing.title = os.path.splitext(os.path.basename(filepath))[0]
            existing_drawing.date = submission_date
            existing_drawing.submission_date = submission_date
            existing_drawing.submission_ref = submission_ref
            updated_drawings += 1
        else:
            # Create a new drawing object and add it to the drawings register
            new_drawing = Drawing(
                drawing_number=drawing_info["drawing_number"],
                title=os.path.splitext(os.path.basename(filepath))[0],
                revision=drawing_info["revision"],
                date=submission_date,
                location_code=drawing_info["location_code"],
                drawing_type=drawing_info["drawing_type"],
                submission_date=submission_date,
                submission_ref=submission_ref,
            )
            drawings.append(new_drawing)
            updated_drawings += 1
    
    print(f"Updated {updated_drawings} drawings from files.")
    # Save the updated drawings list to a CSV file
    with open("drawings.csv", "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["drawing_number", "title", "revision", "date", "location_code", "drawing_type", "submission_date", "submission_ref"])
        for drawing in drawings:
            writer.writerow([
                str(drawing.drawing_number).zfill(6),
                drawing.title,
                str(drawing.revision),    # Convert to string
                drawing.date,
                drawing.location_code,
                drawing.drawing_type,
                drawing.submission_date,
                drawing.submission_ref,
])

def read_drawings_from_csv(file_path):
    drawings = []
    
    if not os.path.exists(file_path):
        return drawings

    with open(file_path, "r", newline='') as csvfile:
        csv_reader = csv.reader(csvfile)
        headers = next(csv_reader)

        sanitized_headers = [header.replace(" ", "_") for header in headers]
        DrawingTuple = namedtuple("DrawingTuple", sanitized_headers)
        
        for row in csv_reader:
            drawing_tuple = DrawingTuple(*row)
            drawing = Drawing(drawing_tuple.Drawing_Number,
                              drawing_tuple.Title,
                              drawing_tuple.Revision,
                              drawing_tuple.Date,
                              drawing_tuple.Location_Code,
                              drawing_tuple.Drawing_Type,
                              drawing_tuple.Submission_Date,
                              drawing_tuple.Submission_Ref)
            drawings.append(drawing)
    return drawings


def print_drawings_table(drawings):
    headers = ["Drawing Number", "Title", "Location Code", "Drawing Type", "Revision", "Date", "Submission Date", "Submission Ref"]
    data = []
    for drawing in drawings:
        data.append([
            drawing.drawing_number,
            drawing.title,
            drawing.location_code,
            drawing.drawing_type,
            drawing.revision,
            drawing.date,
            drawing.submission_date,
            drawing.submission_ref
        ])
    print(tabulate(data, headers=headers, tablefmt="pretty"))
    
# Load the drawings register from a CSV file
csv_file_path = "drawings_register.csv"
drawings_register = read_drawings_from_csv(csv_file_path)

# Prompt the user to select the base folder
root = tk.Tk()
root.withdraw()  # Hide the Tkinter root window

base_folder = filedialog.askdirectory(title="Select the base folder")
if base_folder:
    # Process all drawings found in the folder structure
    process_all_drawings(base_folder, drawings_register)
    

    print_drawings_table(drawings_register)
else:
    print("No folder was selected. Exiting.")