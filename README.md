# AutomatedDR Python script

This Python script is for processing and updating a register of drawings found in a specified folder hierarchy.

## Main components

* Class definition (Drawing): The Drawing class represents a single drawing with properties like drawing number, title, revision, date, location code, drawing type, submission date, and submission reference.
* Functions: There are several functions for handling tasks such as processing and updating the drawings register, reading and writing the register to a CSV file, and printing the register in a tabulated format.
* Main script: The main script first loads the drawings register from a CSV file. It then prompts the user to select a base folder, processes all drawings found in the folder structure, and updates the register accordingly. Finally, it prints the updated drawings register in tabulated format.

## High-level overview of the script's flow

* Load the drawings register from a CSV file (drawings_register.csv).
* Prompt the user to select the base folder.
* Process all drawings found in the folder structure and update the register.
* Print the updated drawings register in a tabulated format.

## Some key functions used in the script

* process_all_drawings(base_folder, drawings_register): Processes all drawings found in the folder hierarchy rooted at base_folder and updates the drawings_register.
* read_drawings_from_csv(file_path): Reads the drawings register from a CSV file and returns a list of Drawing objects.
* update_drawings_in_batch(drawings, filepaths, submission_date, submission_ref): Updates the drawings register with the information from the list of filepaths.

## Benefits of using the AutomatedDR Python script

* It can help you to keep your drawing register up-to-date.
* It can help you to find drawings quickly and easily.
* It can help you to generate reports on your drawing register.
* It can help you to automate tasks related to your drawing register.

## If you are looking for a powerful and easy-to-use tool for managing your drawing register, then I recommend using the AutomatedDR Python script.
