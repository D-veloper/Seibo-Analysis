# import the necessary data science libraries we need to analyse our data
import pandas as pd

pd.set_option('display.max_rows', None)  # set the display option to show all rows when we print our data
pd.set_option('display.max_columns', None)  # set the display option to show all columns when we print our data
pd.set_option('display.max_colwidth', None)  # set the display options to show the full width of the columns

# Load the Excel file. This code assumes the file is in same folder as this script.
file_path = 'Attendance, Enrollment and Meals Fed 2024.xlsx'  # specify the file name.
data = pd.read_excel(file_path)  # read the file and save it inside a variable called data

# Displays first few rows in the data set
print(data.head())

# Returns the number of missing values in each column.
# If a column has missing values, please do something about it before analysing the data
print("\nMissing values in each column:\n", data.isnull().sum())

# Returns the number of unique values in each column.
# Ensure the numbers make sense. Investigate anything that doesn't seem to make sense.
print("\nUnique values in each column:\n", data.nunique())