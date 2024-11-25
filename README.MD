# Automatic Name Badge Maker

This script generates a Word document with name badges based on an attendee list. Follow the steps below to ensure smooth execution.

## Prerequisites
- Python installed on your system.
- A folder containing the required files named as below:
  - `automatic_name_badge_maker.py` (this script).
  - `attendee_list.csv` (attendee information).
  - `logo.PNG` (RCUK logo).

## How to Use

1. **Prepare the Attendee List**:
   - Create a CSV file named `attendee_list.csv` with attendee information.
   - Use the `additional` column to add extra details (e.g., "Plenary" or "Committee").
   - Carefully review the CSV for duplicate bookings or errors. Notes on the ZSL system group bookings will all come through
   on one line, seperated by semi-colons, do not fear, the script will handle these!

2. **Set Up the Folder**:
   - Place the following files in the same folder:
     - `automatic_name_badge_maker.py` (this script).
     - `attendee_list.csv` (prepared in the previous step).
     - `logo.PNG` (RCUK logo file).

3. **Run the Script**:
   - Open a terminal and navigate to the folder containing the files.
   - Execute the following command:
     ```bash
     python automatic_name_badge_maker.py attendee_list.csv logo.PNG badges.docx
     ```

4. **Check the Output**:
   - The script will generate a Word document named `badges.docx`.
   - Carefully review the generated document for accuracy.
