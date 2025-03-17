# AKV_fix_automation
## Description
This script is designed to automate the process of correcting specific numerical values (coefficients, "AKV") associated with certain branches of an insurance contract dataset. Initially, the dataset contained values that required cleaning and adjustment due to inconsistencies in a specific column where the branches were listed. The script identifies and modifies the coefficient values based on the defined rules, and then automatically updates the Excel files accordingly. The program streamlines the process of working with large datasets, making it more efficient and less prone to manual errors.

## Functional Description
The program performs the following steps:
1. Reads multiple Excel files containing contract data.
2. Cleans and standardizes column names for easier processing.
3. Filters the data based on specific conditions related to contract status, reinsurance, tariff, and region.
4. Calculates the correct AKV value based on predefined rules.
5. Updates the original Excel files with the corrected AKV values.
6. Highlights the changed cells for easy reference.

## How It Works
1. The program scans the current working directory for Excel files.
2. It loads each file and processes the data to clean up column names.
3. Then, it applies filtering criteria to extract relevant records for further processing.
4. The `determine_kv` function is used to calculate the correct AKV value for each contract based on specific conditions (e.g., branch, date, tariff).
5. The script updates the files by inserting the corrected AKV values in the corresponding cells.
6. Finally, the modified files are saved, with changed cells being highlighted for clarity.

## Input Structure
The program processes Excel files (.xlsx) with the following structure:
1. The Excel files should contain contract data with relevant columns such as "Статус договора страхования", "Базовый тариф", "Перестрахование", and "Филиал".
2. The file names should be in a format that is recognizable by the script, which processes all `.xlsx` files in the working directory.

## Technical Requirements
To run the program, the following are required:
1. Python 3.x
2. Installed libraries: `pandas`, `openpyxl`, `warnings`, `re`
3. The Excel files should be structured properly with columns as specified in the script for processing.

## Usage
1. Place the Excel files you wish to process in the working directory.
2. Run the script. It will:
   - Automatically process each file.
   - Correct the AKV values according to the defined rules.
   - Highlight the changed cells in red.
3. The modified files will be saved with the corrected AKV values.

## Example Output
For each processed file, the following steps occur:
1. Corrected AKV values are inserted in the designated columns.
2. The cells with changes are highlighted for easy identification.

## Conclusion
This tool automates the process of correcting and updating AKV values in large insurance contract datasets, improving efficiency and reducing the likelihood of human error. It streamlines the data cleaning process and saves time when working with multiple Excel files containing contract data.
