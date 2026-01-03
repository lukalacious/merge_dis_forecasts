# import libraries
import os
import pandas as pd
from datetime import datetime

print("CHECKPOINT 1: Starting Distributor Forecast Merge Script")
print("CHECKPOINT 1a: Libraries imported successfully")

# view all files in the folder
print("CHECKPOINT 2: Getting user input for folder name...")
ftw_apeq = input()

folder_path = rf"C:/Users/luke.roberts/OneDrive - ASICS Corporation/Documents/vs_code/input/{ftw_apeq}"
print(f"CHECKPOINT 2a: Folder path set to: {folder_path}")

excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')] # list of all files in the folder
print(f"CHECKPOINT 2b: Found {len(excel_files)} Excel files in folder")
print(f"Files found: {excel_files}")
excel_files

# create a dictionary to store dataframes
dataframes = {}
print("CHECKPOINT 3: Starting file import process...")

# import each file as its own dataframe variable and save in the dictionary

# loop through each file and import the specific sheet
for file in excel_files:
    file_name = file.split("--")[-1].replace(".xlsx", "").strip() # splits the file after"--" to use as name
    file_path = os.path.join(folder_path, file)
    try:
        dataframes[file_name] = pd.read_excel(file_path, sheet_name="ASSORTMENT") # import assortment sheet
        print(f"CHECKPOINT 3a: Successfully imported {file_name}") # print the file name and path
    except PermissionError as e:
        print(f"PermissionError: Could not open {file_path}. Please close the file if it is open in another program.")
    except Exception as e:
        print(f"Error importing {file_name}: {e}")

print(f"CHECKPOINT 3b: File import complete. Total dataframes created: {len(dataframes)}")

# rename dfs to their key
print("CHECKPOINT 4: Assigning dataframes to global variables...")
for key, df in dataframes.items():
    globals()[key] = df

# Get list of successfully imported dataframe names
working_dfs = list(dataframes.keys())
print(f"CHECKPOINT 4a: Working with {len(working_dfs)} dataframes: {working_dfs}")

print("CHECKPOINT 5: Displaying dataframe previews...")
for key in dataframes.keys():
    df = dataframes[key]
    print(f"DataFrame name: {key}")
    print(df.head(10))
    print("-_"*50)

# Drop rows where column 1 ("Unnamed: 1") is empty (NaN) for each dataframe
print("CHECKPOINT 6: Starting data cleaning - dropping empty rows...")
for df in list(dataframes.values()):
    df.dropna(subset=[df.columns[0]], inplace=True)
    df.reset_index(drop=True, inplace=True)
print("CHECKPOINT 6a: Empty rows dropped from all dataframes")

# set first row as header for each dataframe and reset index
print("CHECKPOINT 7: Setting first row as headers for all dataframes...")
for name in working_dfs:
    df = globals()[name]
    df.columns = df.iloc[0]
    globals()[name] = df[1:].reset_index(drop=True)
print("CHECKPOINT 7a: Headers set successfully for all dataframes")

# print shape of each dataframe
print("CHECKPOINT 8: Checking dataframe shapes after cleaning...")
for name in working_dfs:
    df = globals()[name]
    print(f"shape {name}: {df.shape}")
    print("-_"*50)

# print 2 rows and 4 columns of each dataframe
print("CHECKPOINT 9: Displaying cleaned dataframe previews...")
for name in working_dfs:
    df = globals()[name]
    print(f"{name}")
    print(df.iloc[:2, :4])
    print("-_"*50)

# add a new column to each dataframe with dataframe name set to each row
print("CHECKPOINT 10: Adding DIS column to each dataframe...")
for name in working_dfs:
    df = globals()[name]
    df['DIS'] = name  # Add a new column named 'DIS' with the dataframe name
    globals()[name] = df  # Update the global variable
    print(f"CHECKPOINT 10a: DIS column added to {name}")

# Concatenate all dataframes in working_dfs vertically
print("CHECKPOINT 11: Starting dataframe concatenation...")
dis_consolidated = pd.concat([globals()[name] for name in working_dfs], ignore_index=True)
print(f"CHECKPOINT 11a: Concatenation successful! Shape of consolidated dataframe: {dis_consolidated.shape}")
print(dis_consolidated.head())

# Use .count() instead of .counts() to get the count of non-NA cells for each column per group
print("CHECKPOINT 12: Checking distribution counts by DIS...")
counts_by_dis = dis_consolidated.groupby('DIS')['DIS'].count()
print(counts_by_dis)

# enter season
print("CHECKPOINT 13: Getting user input for export filename...")
print("enter file export name, eg SS26_FTW_MF2:")
season = input(f"file name: {str()}")

# export dis_stacked as excel file
print("CHECKPOINT 14: Starting file export...")
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
output_file = f"C:/Users/luke.roberts/OneDrive - ASICS Corporation/Documents/vs_code/output/dis_consolidated_{season}_{timestamp}.xlsx"
dis_consolidated.to_excel(output_file, index=False)

print(f"CHECKPOINT 15: EXPORT SUCCESSFUL! File saved as: {output_file}")
print("CHECKPOINT 16: Script execution completed successfully!")
print(f"Final consolidated dataframe contains {len(dis_consolidated)} rows and {len(dis_consolidated.columns)} columns")