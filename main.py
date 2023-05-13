import re
import os
import csv
import glob
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
                valid_files = [drawing_file for drawing_file in drawing_files
                                if re.search(r"HKT2_BYME_SDWG_([A-Za-z0-9]+)[_-]([A-Za-z0-9]+)[_-](\d+)[_-](\w)", os.path.basename(drawing_file))]

                invalid_files = set(drawing_files) - set(valid_files)

                for invalid_file in invalid_files:
                    log_file.write(f"Invalid file format: {invalid_file}\n")

                print(f"Valid files found: {len(valid_files)}")
                print(f"Invalid files found: {len(invalid_files)}")

                if valid_files:
                    update_drawings_in_batch(drawings_register, valid_files, submission_date, drawing_type_folder)


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

        title = os.path.splitext(os.path.basename(filepath))[0]

        if existing_drawing:
            existing_drawing.title = title
            existing_drawing.date = submission_date
            existing_drawing.submission_date = submission_date
            existing_drawing.submission_ref = submission_ref
            updated_drawings += 1
        else:
            # Create a new drawing object and add it to the drawings register
            new_drawing = Drawing(
                drawing_number=drawing_info["drawing_number"],
                title=title,
                revision=drawing_info["revision"],
                date=submission_date,
                location_code=drawing_info["location_code"],
                drawing_type=drawing_info["drawing_type"],
                submission_date=submission_date,
                submission_ref=submission_ref,
            )
            drawings.append(new_drawing)
            updated_drawings += 1

            print(f"Updated {updated_drawings} drawings in the batch")

def save_drawings_to_csv(drawings, output_file):
    with open(output_file, "w", newline="") as csvfile:
        fieldnames = ["drawing_number", "title", "revision", "date", "location_code", "drawing_type", "submission_date", "submission_ref"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        for drawing in drawings:
            writer.writerow({
                "drawing_number": f"{drawing.drawing_number:06d}",  # Format the drawing number as a zero-padded string with a width of 6 digits
                "title": drawing.title,
                "revision": drawing.revision,
                "date": drawing.date,
                "location_code": drawing.location_code,
                "drawing_type": drawing.drawing_type,
                "submission_date": drawing.submission_date,
                "submission_ref": drawing.submission_ref,
            })

def main():
    # Create a list to store all the drawings
    drawings_register = []

    # Prompt the user to select a folder
    root = tk.Tk()
    root.withdraw()
    base_folder = filedialog.askdirectory(title="Select a folder containing the drawing files")

    if not base_folder:
        print("No folder selected. Exiting.")
        return

    # Process all drawings in the selected folder and its subfolders
    process_all_drawings(base_folder, drawings_register)

    # Save the updated drawings register to a CSV file
    output_file = "drawings_register.csv"
    save_drawings_to_csv(drawings_register, output_file)

    print("Finished processing. Saved updated drawings register to:", output_file)

if __name__ == "__main__":
    main()