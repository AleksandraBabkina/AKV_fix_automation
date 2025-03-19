# AKV_fix_automation

## Description

This script processes multiple Excel files to filter and modify data based on specific conditions. It reads the data from each Excel file, filters rows based on given criteria, calculates the `AКВ` coefficient for insurance contracts, and updates the corresponding cells in the file. The modified cells are highlighted in red for easy identification. The script operates on files with the `.xlsx` extension and applies custom logic to determine the `AКВ` coefficient.

## Functional Description

The script performs the following steps:
1. **Cleaning Column Names**: Removes unwanted parts from column names, including level indicators and newline characters.
2. **Calculating `КВ` Coefficient**: Based on certain conditions (such as the base tariff and other contract details), it determines the appropriate `КВ` coefficient for each row.
3. **Filtering Data**: Filters the data based on multiple conditions related to contract status, reinsurance, tariff, and other contract details.
4. **Modifying Excel Files**: Iterates through Excel files, applies the filtering and modification logic, and saves the updated files. It also highlights modified cells in red to indicate changes.

## Input Structure

1. **Files**: The script processes all `.xlsx` files in the current working directory.
2. **Columns in the Excel File**:
   - `Статус договора страхования`: Contract status.
   - `Превышение КВ согласовано с андеррайтерами`: Whether the `КВ` increase was agreed with underwriters.
   - `Перестрахование`: Reinsurance flag.
   - `Базовый тариф`: Base tariff.
   - `ЦФО`: Central Financial Organization.
   - `Срок действия договора подписание`: Contract start date.
   - `Филиал`: Branch.
   - `Договор страхования Номер`: Insurance contract number.
   
## Key Functions

1. **`clean_column_name(name)`**:
   - Removes unnecessary parts from column names like "Unnamed" and newlines, and returns a cleaned-up version.
   
2. **`determine_kv(row)`**:
   - Determines the `КВ` coefficient based on conditions:
     - **Base Tariff** (`Базовый тариф`) = 1111 or 2222.
     - For `Базовый тариф` = 2222, checks the branch and contract signing date.
     - Applies specific logic for certain branches and dates.

3. **Excel Modification**:
   - Loops through each row of the Excel file and updates the corresponding cell with the calculated `КВ` value.
   - Highlights the updated cells in red using the `openpyxl` library.
   
## Output

- The `КВ` coefficient is calculated and stored in a new column for each row.
- The script highlights cells in red where modifications were made.
- The updated data is saved back into the same Excel file.

## Example Workflow

1. The script reads an Excel file and cleans up column names.
2. It filters the data based on certain conditions, such as:
   - Contract status is "In force" (`Вступил в силу`).
   - Reinsurance status is not agreed with underwriters (`Превышение КВ согласовано с андеррайтерами` = "нет").
   - Base tariff is either 1111 or 2222.
3. The script applies the custom logic to determine the `КВ` coefficient.
4. The `КВ` values are written into the corresponding cells in the Excel file, and the updated cells are highlighted in red.
5. The modified Excel file is saved with the new data.

## Error Handling

If a file does not meet the required format or if an error occurs while processing a file, the script will print an error message and skip that file.

## Example Output

- **Excel Modifications**: After processing, cells that have been modified will be highlighted in red.
- **Final Message**: After completing the processing, the script prints.

## Conclusion

This script automates the process of filtering and modifying contract data in Excel files. It applies custom logic to determine the `КВ` coefficient, updates the data, and highlights changes for easy identification. The use of `openpyxl` allows for efficient modifications to the Excel files, ensuring data integrity and clarity.
