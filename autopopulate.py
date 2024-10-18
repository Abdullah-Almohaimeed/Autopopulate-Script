# This script will help me autopopulate the year column in the excel sheet
from openpyxl import load_workbook
import pandas as pd

# If you want to preserve formatting, use openpyxl, check the last part of the script
try:
    preserve_format = load_workbook('Master_List.xlsx')
    #print(preserve_format.sheetnames)
except FileNotFoundError:
    print("File not found")
    exit(1)

# PLEASE MAKE SURE TO INCLUDE THIS LINE. IF THE SHEET IS HIDDEN, IT WILL CORRUPT THE FILE
# Adjust the sheet name according to your data
game_sheet = preserve_format['Master_List']
if game_sheet.sheet_state != 'visible':
    game_sheet.sheet_state = 'visible'

# Load main dataset. Year column needs to be autopopulated
dataframe_main = pd.read_excel('Master_List.xlsx')

# Load the dataset that contains the data for the year column, I found one in github, CSV format
dataframe_years = pd.read_csv('Video_Games.csv')

#Checking
#print(dataframe_main.head())
#print(dataframe_years.head())


def lookup_release_year(game_name, years_dataframe):
    # Search for the game in the release years data frame
    game = years_dataframe[years_dataframe['Name'].str.lower() == game_name.lower()]
    
    # Return the release year if found, otherwise return None
    if not game.empty:
        return game['Year'].values[0]
    else:
        return None  # You can return a default value here if you want


    
# Apply the function to rows where 'Year' is missing
dataframe_main['Year'] = dataframe_main.apply(
    lambda row: lookup_release_year(row['Game'], dataframe_years)
    if pd.isna(row['Year'])
    else row['Year'],
    axis=1
)

#Checking
#print(dataframe_main.head())

# Save updated data frame. I'm overwriting the original file, be careful if you care about your data
# I'm using openpyxl as the engine to preserve formatting
# I'm using this dumb logic because I can't save the file with pandas without losing the formatting.
# Time complexity of O(n^godknows)
for row in game_sheet.iter_rows(min_row=2, max_row=game_sheet.max_row, min_col=1, max_col=game_sheet.max_column):
    game_name = row[0].value  # Change index values according to your data
    # Find year value for game, also handling case where game not found
    year_value = dataframe_main.loc[dataframe_main['Game'] == game_name, 'Year'].values[0]
    if year_value is not None:
        row[4].value = year_value  # Change index values according to your data
    else:
        print(f"Year not found for {game_name}") 
        exit(1) # Exiting so as not to alter the file
# Save the workbook 
preserve_format.save('Master_List.xlsx')