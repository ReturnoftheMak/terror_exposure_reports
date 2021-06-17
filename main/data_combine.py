# %% Package Imports

import pandas as pd
import glob


# %% Params - change the output filepath as needed

directory = r'\\basapoprdfpsv05.apollo.local\apollo_prd_file_share\Underwriting\Crisis Management\Exposure\QUOTING'

# A quick glance at some random directories showed one with a lowercase m, these two should hopefully cover all cases.

match_1 = r'\**\*Terror_250m*.xls*'
match_2 = r'\**\*Terror_250M*.xls*'

# Change this if using
output_filepath = r'\\insert\relevant\directories\and\filename.xlsx'


# %% Get filepaths and then read them all in

files_terror = []

for file in glob.glob(directory + match_1):
    files_terror.append(file)

for file in glob.glob(directory + match_1):
    files_terror.append(file)

# // One liner without error handling
# // df = pd.concat([pd.read_excel(file, header=1) for file in files], ignore_index=True)

def concat_files(files:list) -> pd.DataFrame:

    df_list = []
    for file in files:
        try:
            df = pd.read_excel(file, header=1)
            df_list.append(df)
        except:
            pass

    df_combined = pd.concat(df_list)
    return df_combined


# %% Run + Export

df = concat_files(files_terror)
df.to_excel(output_filepath)

