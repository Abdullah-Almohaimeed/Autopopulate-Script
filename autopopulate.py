from openpyxl import load_workbook  # Reading and writing without losing formatting
from fuzzywuzzy import process  # For name matching
import pandas as pd  # Data manipulation
import os  # For checking file existence

# Function to get and validate file information from the user
def get_file_info(file_purpose):
    """
    Function to get and validate file information from the user.
    It validates file format, file existence, and whether the column exists in the file.

    Args:
    file_purpose (str): Purpose of the file ("main file" or "data file").

    Returns:
    tuple: (file_name, column_name, file_flag, matcher_column)
    """
    
    # Take file name from user
    file_name = input(f"Enter the name of the {file_purpose}, make sure it's in the same directory of this script: ").strip()

    file_flag = 0   # Flag to know file format, 1 for Excel, 2 for CSV
    
    # Loop until valid file format is given
    while True:
        file_format = input(f"Enter the format of the {file_purpose}, 1 for Excel, 2 for CSV: ").strip()  # Only CSV and xlsx supported
        match file_format:  # Match the file format
            case '1':
                file_name += '.xlsx'  # Add the extension
                file_flag = 1
                break
            case '2':
                file_name += '.csv'  # Add the extension
                file_flag = 2
                break
            case _:
                print("Invalid input")

    # Check if file exists
    if not os.path.exists(file_name):
        print(f"Error: The file '{file_name}' does not exist. Please check the file name and try again. Also check if it's in the same directory as the script.")
        exit(1)

    # Take column name input from the user, strip any extra spaces
    column_name = input(f"Enter the column name of the {file_purpose}, i.e. column that will have data be taken from or added to: ").strip()

    # Check if column exists in the file
    if file_flag == 1:
        # Excel file, check the columns
        try:
            df = pd.read_excel(file_name)
        except Exception as e:
            print(f"Error reading the Excel file: {e}")
            exit(1)
    elif file_flag == 2:
        # CSV file, check the columns
        try:
            df = pd.read_csv(file_name)
        except Exception as e:
            print(f"Error reading the CSV file: {e}")
            exit(1)

    # Check if the provided column exists in the file
    if column_name not in df.columns:
        print(f"Error: The column '{column_name}' does not exist in the {file_purpose}. Please check the column name.")
        exit(1)

    # Take the column that will be used for matching
    matcher_column = input(f"Enter the column name of the {file_purpose} that will be used for matching: ").strip()

    # Check if matcher column exists in the file
    if matcher_column not in df.columns:
        print(f"Error: The column '{matcher_column}' does not exist in the {file_purpose}. Please check the column name.")
        exit(1)

    return file_name, column_name, file_flag, matcher_column


# Take inputs from the user for the main file
file_name_main, file_column_main, file_flag_main, file_matcher_main = get_file_info("main file")  # Main file, one that will be updated

# Take inputs from the user for the data file
file_name_data, file_column_data, file_flag_data, file_matcher_data = get_file_info("data file")  # Data file, one that has data to update main file

print("Processing...") # Very important for the user to know that the script is running

# If you want to preserve formatting, use openpyxl, check the last part of the script
# Otherwise, modify this script to use pandas only 
try:
    preserve_format = load_workbook(file_name_main)
except FileNotFoundError:
    print("File not found, make sure it's in the same directory as the script")
    exit(1)

# PLEASE MAKE SURE TO INCLUDE THIS LINE. IF THE SHEET IS HIDDEN, IT WILL CORRUPT THE FILE
game_sheet = preserve_format.active  # Adjust the sheet name according to your data
if game_sheet.sheet_state != 'visible':
    game_sheet.sheet_state = 'visible'

# Get the header names and their respective column indexes
header = {cell.value: idx for idx, cell in enumerate(next(game_sheet.iter_rows(min_row=1, max_row=1)), start=0)}

# Check if the required columns exist in the sheet
if file_matcher_main not in header or file_column_main not in header:
    print(f"Error: Required columns '{file_matcher_main}' or '{file_column_main}' not found in the sheet.")
    exit(1)

# Get the column index for the matching key and the column to be updated
matcher_column_index = header[file_matcher_main]  # e.g., index for the 'Game' column
update_column_index = header[file_column_main]     # e.g., index for the 'Year' column

# Load main dataset
if file_flag_main == 1:
    dataframe_main = pd.read_excel(file_name_main)
else:
    dataframe_main = pd.read_csv(file_name_main)

# Load the dataset that contains the data for main file
if file_flag_data == 1:
    dataframe_years = pd.read_excel(file_name_data)
else:
    dataframe_years = pd.read_csv(file_name_data)

# Look up values from one column and populate another based on a matching key, keep in mind this is fuzzy matching and may not be 100% accurate
def lookup_value(main_value, data_dataframe, match_column_data, return_column_data):
    best_match, score, *_ = process.extractOne(main_value, data_dataframe[match_column_data].str.lower())
    if score >= 80:  # Change the score according to your data
        return data_dataframe.loc[data_dataframe[match_column_data].str.lower() == best_match, return_column_data].values[0]
    else:
        return None


# Apply lookup function to rows where the target column is missing
dataframe_main[file_column_main] = dataframe_main.apply(
    lambda row: lookup_value(row[file_matcher_main], dataframe_years, file_matcher_data, file_column_data)
    if pd.isna(row[file_column_main])
    else row[file_column_main],
    axis=1
)


# Save updated data frame
for row in game_sheet.iter_rows(min_row=2, max_row=game_sheet.max_row):
    main_value = row[matcher_column_index].value  # Use the dynamically found index for matching
    
    # Lookup the value in the main DataFrame based on the matching key
    try:
        matched_value = dataframe_main.loc[dataframe_main[file_matcher_main] == main_value, file_column_main].values[0]
    except IndexError:
        matched_value = None
    
    if matched_value is not None:
        # Update the appropriate cell in the game sheet using the dynamic column index
        row[update_column_index].value = matched_value
    else:
        print(f"Value not found for {main_value}")
        exit(1)

# Save the workbook after updating
preserve_format.save(file_name_main)
print('Done')
