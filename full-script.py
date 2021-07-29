import pandas as pd
import glob
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import numpy as np
import csv

# This creates an empty dataframe
data = pd.DataFrame(columns=['Publisher', 'Application', 'Version', 'Install Date', 'Last HW Scan'])

# The filepath needs to be where the files are located on your device
file_paths = glob.glob('/Users/jeremylee/Desktop/CS-Projects/T4SG/usda-app-rationalization-summer/data/*.xlsx')

# Reading in the data
# This might take some time, but you'll be able to see the filenames as they're being read in
print("Reading in data...")
for filepath in file_paths:
    data = data.append(pd.read_excel(filepath))
print('Data read completed.')

drop_columns = ["AD_Site_Name0", "User", "Agency"]
data.drop(columns=drop_columns, inplace=True)

print("Removing servers...")
### Remove all the rows of data from Microsoft servers, as we’re only interested in workstations
data['C054'].fillna('',inplace=True)  # replaces NaN values in column C054 with ‘’ nothing, needed for next segment of code to work
index = data[data['C054'].str.lower().str.contains('server')].index # creates an index of rows containing server in the OS column (C054)
data.drop(index, inplace=True)  # this deletes all rows where the word 'server' is in column C054, from prior line's index
print("Server removal completed.")

print("Removing GOTS applications...")
### Removing GOTS applications
data = data[data["Publisher"].str.lower().str.contains("usda") == False]
print("GOTS application removal complete.")

### Flagging utilities
print("Flagging utilities...")

# Tagging utilities based on keywords in app names
utility_keywords = ["driver", "update", "compiler", "decompiler", 
                    "installer", "utility", "plugin", "tool"]

data["Application"].fillna('',inplace=True)
data["Utility"] = 0
for keyword in utility_keywords:
    data.loc[data["Application"].str.lower().str.contains(keyword, na=False), "Utility"] = 1

# Tagging utilities based on particular publishers
utility_publishers = ["Intel", "Intel Corporation", "Intel(R) Corporation",
                      "Advanced Micro Devices, Inc.", "Advanced Micro Devices Inc.", "AMD",
                      "Dell", "Dell Inc.", "Dell, Inc.", "Dell Inc"]

# Tagging utilities with a 1 in the "Utility" column
data["Publisher"].fillna('',inplace=True)
data.loc[data.Publisher.isin(utility_publishers), "Utility"] = 1

utility_list = data[data["Utility"] == 1].drop(columns="Utility")
utility_list = utility_list[["Publisher", "Application"]].drop_duplicates()
utility_df = pd.DataFrame(utility_list)

# Getting all utilities and their installation counts
counts = []
value_counts = data.groupby("Application").count()["System Name"]
for app in utility_df["Application"]:
    counts.append(value_counts[app])
utility_df["Count"] = counts
utility_df.sort_values(by="Count", ascending=False, inplace=True)

# Exporting utility results to csvs
utility_df.to_csv("utilities.csv")
data.to_csv("flagged_utilities.csv")
print("Utility flagging complete.")

### Normalizing App Names
print("Normalizing app names...")

# Getting unique app names
main_df = data[data["Utility"] == 0].drop(columns="Utility")
business_apps = list(main_df.Application.unique())

# Creating a dataframe with unique applications and their counts (# of installations)
counts = []
value_counts = main_df.groupby("Application").count()["System Name"]
for app in business_apps:
    counts.append(value_counts[app])

normalized_df = pd.DataFrame()
normalized_df["old_name"] = business_apps
normalized_df["new_name"] = business_apps
normalized_df["count"] = counts

normalized_df.sort_values(by="count", ascending=False, inplace=True)

# Normalizing names

# Filters out fully numerical words (targeting years and version numbers)
def is_non_number(s):
    if len(s) == 0:
        return False
    if s[0].lower() == "v":
        for char in s[1:]:
            if char.isalpha():
                return True
        return False
    else:
        for char in s:
            if char.isalpha():
                return True
        return False

# Removes commas at the ends of words
def remove_comma(word):
    if len(word) == 0:
        return word
    if word[-1] == ",":
        return word[:-1]
    return word

# Removes 32-bit and 64-bit tags on app names
tags = ["ARM64", "arm64", "amd64", "arm", "ARM",
        "X64", "X86", "x64", "x86", "64-bit", "32-bit", "32bit", "64bit"]

def remove_tag_word(word):
    new_word = word
    for tag in tags:
        if "(" + tag + ")" in new_word:
            new_word = new_word.replace("(" + tag + ")", "")
        elif "_" + tag in new_word:
            new_word = new_word.replace("_" + tag, "")
        elif tag in new_word:
            new_word = word.replace(tag, "")
        
    return new_word

# Removes starting and ending parentheses for words
def remove_parentheses_word(word):
    if len(word) == 0:
        return word
    if (word[0] == "(" and ")" not in word) or (word[-1] == ")" and "(" not in word):
        return ""
    return word

# Removes the word "version"
def remove_version_word(word):
    if "version" in word.lower():
        return False
    return True

# Removes words that are empty
def remove_blank_words(word):
    return word != ""

def normalize_name(app_name):
    words = str(app_name).split()
    
    words = list(filter(is_non_number, words)) # Removing version and years
    words = list(map(remove_comma, words)) # Removing ending commas
    words = list(map(remove_tag_word, words)) # Removing 32 vs 64 bit
    words = list(map(remove_parentheses_word, words)) # Removing parentheses
    words = list(filter(remove_version_word, words)) # Removing the word "version"
    words = list(filter(remove_blank_words, words)) # Removing blank words
    
    return " ".join(words)

normalized_df["new_name"] = normalized_df.old_name.apply(normalize_name)

# Outputting normalized names to csv
normalized_df.to_csv("normalized_apps.csv")

# Comparing unique counts
print("No. of Unique (un-normalized) App Names: " + str(normalized_df["old_name"].unique().shape[0]))
print("No. of Unique Normalized App Names: " + str(normalized_df["new_name"].unique().shape[0]))

print("App normalizing complete.")

# Changing the setting so we could see all the lines of the dashboard
pd.set_option('display.max_rows', None)

# Dashboard by application name and number of installations in the whole dataset
full_data = data.groupby(['Publisher', 'Application'])
full_data = full_data.size().sort_values( ascending=False).reset_index(name='# of all entries')
dashboard = pd.DataFrame(full_data)

### Creating a dashboard without versions

print("Creating dashboard without versions...")

# Dashboard by application name and number of duplicated installations in the whole dataset
duplicates = data.where(data.duplicated(subset=['Application', 'System Name', 'Last HW Scan']) == True).dropna(how='all')
dupl_dash = duplicates.groupby(['Publisher', 'Application'])
dupl_dash = dupl_dash.size().sort_values( ascending=False).reset_index(name='# of duplicate installations')
dupl_df = pd.DataFrame(dupl_dash, columns = ['Publisher', 'Application', '# of duplicate installations'])

# Dashboard by application name and number of not duplicated installations in the whole dataset
data_no_dupl = data.where(data.duplicated(subset=['Application', 'System Name', 'Last HW Scan']) != True).dropna(how='all')
data_no_dupl = data_no_dupl.groupby(['Publisher', 'Application'])
data_no_dupl = data_no_dupl.size().sort_values( ascending=False).reset_index(name='# of unique installations')
no_dupl = pd.DataFrame(data_no_dupl, columns = ['Publisher', 'Application', '# of unique installations'])

# Putting the sub-dashboards all together
dashboard = dashboard.merge(dupl_df, how='left', on=['Publisher', 'Application'])
dashboard = dashboard.merge(no_dupl, on=['Publisher', 'Application'])
dashboard = dashboard.fillna(value=0)
dashboard['# of duplicate installations'] = dashboard['# of duplicate installations'].astype(int)

print('The number of distinct applications is:')
print(dashboard.shape[0])
print('The number of duplicated installations is:')
print(dupl_df['# of duplicate installations'].sum())

# Writing the dashboards into Excel files
dupl_df.to_excel('duplicates_07_06.xlsx')
dashboard.to_excel('full_dashboard_07_06.xlsx')

print("Dashboard without versions complete.")

### Creating dashboards that DO take into account the version number

print("Creating dashboard with versions...")

# Creating a version based duplicates dashboard
version_data = data.copy()
version_data['Version'] = version_data['Version'].astype(str)

# Full dashboard for applications with versions and the number of all installations
v_full_data = version_data.groupby(['Publisher', 'Application', 'Version'])
v_full_data = v_full_data.size().sort_values( ascending=False).reset_index(name='# of all entries')
v_dashboard = pd.DataFrame(v_full_data)

# Finding duplicates by versions
v_duplicates = version_data.where(version_data.duplicated(subset=['Application', 'Version', 'System Name', 'Last HW Scan']) == True).dropna(how='all')
v_dupl_dash = v_duplicates.groupby(['Publisher', 'Application', 'Version'])
v_dupl_dash = v_dupl_dash.size().sort_values( ascending=False).reset_index(name='# of duplicate installations')

# Finding unique installations by versions
v_no_dupl = version_data.where(version_data.duplicated(subset=['Application', 'Version', 'System Name', 'Last HW Scan']) != True).dropna(how='all')
no_dupl_by_v = v_no_dupl.groupby(['Publisher', 'Application', 'Version'])
no_dupl_by_v = no_dupl_by_v.size().sort_values( ascending=False).reset_index(name='# of unique installations')

# To verify that when summing the duplicates up by application it the numbers are similar to the previous dashboards
v_dupl_byapp = v_duplicates.groupby(['Publisher', 'Application'])
v_dupl_byapp = v_dupl_byapp.size().sort_values( ascending=False).reset_index(name='# of duplicate installations')

# Putting everything into the full dashboard
v_no_dupl_df = pd.DataFrame(no_dupl_by_v, columns = ['Publisher', 'Application', 'Version', '# of unique installations'])
v_dupl_df = pd.DataFrame(v_dupl_dash, columns = ['Publisher', 'Application', 'Version', '# of duplicate installations'])
v_full = v_dashboard.merge(v_dupl_df, how='left', on=['Publisher', 'Application', 'Version'])
v_full = v_full.merge(v_no_dupl_df, on=['Publisher', 'Application', 'Version'])
v_full = v_full.fillna(value=0)
v_full['# of duplicate installations'] = v_full['# of duplicate installations'].astype(int)

# Writing the new dashboards into Excel files
v_full.to_excel('full_dashboard_w_versions_07_06.xlsx')
v_dupl_dash.to_excel('duplicates_w_versions_07_06.xlsx')

# Sorting the dashboard by # of unique installations for the bundling algorithm
dashboard = dashboard.sort_values(by=['# of unique installations'], ascending=False).reset_index(drop=True)

print("Dashboard with versions complete.")

### Bundle-identification code

print("Bundling applications...")

import re
from collections import defaultdict

# filter out commonly used words in application titles
common_words = ['Tool', 'Module', 'Update', 'Software', '', 'App', 'Client',
                'Tools', 'for', 'and', 'in', 'Client', 'Installer', 'Drive', 
                'Driver', 'Web', 'Helper', 'Support', 'Center', 'Manager',
                'File', 'Reader', 'C', 'Launcher', 'Plugin', 'Service', 'Setup', 'Driver',
               'x86', '(x86)', 'x64', '(x64)', 'X86']

# filter out numbers and version numbers
def is_number(string):
    matching = re.match(r'^\d+(\.\d+)*$', string)
    if matching:
        return True
    return False

# cleaning the name of characters such as dash, underscore etc.
def clean_name(name):
    words = re.split(' |,|_|-', name)
    words = list(filter(lambda w: w not in common_words, words))
    words = list(filter(lambda w: not is_number(w), words))
    return words

### create a spreadsheet that lists applications and their potential bundles based on fuzzywuzzy scores

# Converts application at row index to form "[name] || [publisher] || [version]"
def stringify(dashboard, index):
    name = dashboard.loc[[index]]['Application'].iat[0]
    pub = dashboard.loc[[index]]['Publisher'].iat[0]
    version = dashboard.loc[[index]]['Version'].iat[0]
    return name + " || " + pub + " || " + version

# Confirms that cleaned potential bundled application names have common words
def checkIfSubword(name, other):
    name_clean = clean_name(name)
    other_clean = clean_name(other)
    # If any of the apps are 1 word long, don't have to have a common string
    if len(name) == 1 or len(other) == 1:
        return True
    return bool(set(name_clean) & set(other_clean))

# Checks that two applications are installed on at least 70% of the same workstations
def checkIfSimilarWorkstations(name, other):
    filtered_1 = data[data.Application == name]
    filtered_2 = data[data.Application == other]
    flist1 = filtered_1["System Name"].to_list()
    flist2 = filtered_2["System Name"].to_list()
    res = len(set(flist1) & set(flist2)) / float(len(set(flist1) | set(flist2))) * 100
    # This similarity threshold (70%) can be adjusted higher or lower
    return res >= 70

# Checks that two applications have the same version number
def checkVersion(dashboard, index, other_index):
    v1 = dashboard.loc[[index]]['Version'].iat[0]
    v2 = dashboard.loc[[other_index]]['Version'].iat[0]
    return v1 == v2

# Checks if two applications should be bundled
def checkIfBundle(dashboard, index, other_index):
    name = dashboard.loc[[index]]['Application'].iat[0]
    other = dashboard.loc[[other_index]]['Application'].iat[0]
    name_clean = clean_name(name)
    other_clean = clean_name(other)
    # Fuzzywuzzy checks that the names of the two applications are roughly similar (>50% similarity)
    return checkVersion(dashboard, index, other_index) and checkIfSubword(name, other) and fuzz.partial_ratio(name_clean, other_clean) >= 50 and checkIfSimilarWorkstations(name, other)

# Creates a csv of potential bundles, taking in a dashboard as input
def createBundle(dashboard):
    # Sorts dashboard by number of unique installations
    dashboard = dashboard.sort_values(by=['# of unique installations'], ascending=False).reset_index(drop=True)
    
    maxRows = len(dashboard.index)
    bundleList = []
    
    # Creates a dashboard column "grouped" that indicates whether an application has been bundled
    dashboard['grouped'] = False
    
    # needs to be sorted by count
    for index, row in dashboard.iterrows():
        bundle = []
        count = row['# of unique installations']
        # stop algo if # of installations is less than 100
        if count < 100:
            break
        # if the application has been grouped, skip it
        if dashboard.loc[[index]]['grouped'].iat[0] == True:
            continue
        name = row['Application']
        # parse through above rows until count >10% difference
        tempIndex = index - 1
        while tempIndex >= 0 and dashboard.loc[[tempIndex]]['# of unique installations'].iat[0] - count <= count/10:
            if checkIfBundle(dashboard, index, tempIndex):
                # Setting grouped for the application being analyzed
                dashboard.at[index, 'grouped'] = True
                dashboard.at[tempIndex, 'grouped'] = True
                bundle.append(stringify(dashboard, tempIndex))
            tempIndex -= 1
        # parse through below rows until count <10% difference
        tempIndex = index + 1
        while tempIndex < maxRows and count - dashboard.loc[[tempIndex]]['# of unique installations'].iat[0] <= count/10:
            if checkIfBundle(dashboard, index, tempIndex):
                dashboard.at[index, 'grouped'] = True
                dashboard.at[tempIndex, 'grouped'] = True
                bundle.append(stringify(dashboard, tempIndex))
            tempIndex += 1
        # indicate that the program is running
        if index % 10 == 0:
            print(index, end=' ')
        # insert the app name and the count
        bundle.insert(0, stringify(dashboard, index))
        bundle.insert(1, count)
        bundleList.append(bundle)
        
    # Creates v_bundles.csv with bundles
    df = pd.DataFrame(bundleList)
    df.to_csv('v_bundles.csv', index=False, header=False)
    
    # prints Done when algo has finished
    print("Application bundling complete.")

# Creates a csv of bundles from the v_full dashboard (defined above)
createBundle(v_full)


### Identify problematic applications that have typos in publisher name

print("Identifying problematic applications...")

# Detecting applications that have more than 100 installations
applications = dashboard[dashboard['# of unique installations'] > 100]['Application'].unique()
prob_apps = []
for app in applications:
#     detecting applications that show up in the dashboard at least twice
    if len(dashboard[(dashboard['Application'] == app) & (dashboard['# of unique installations'] > 100)]['# of unique installations']) != 1:
        if app not in prob_apps:
            prob_apps.append(app)

problem_df = pd.DataFrame(columns=['Application', 'Publishers'])

# Pairing up the different publisher names for each problematic application
for index, row in dashboard[dashboard['# of unique installations'] > 100].iterrows():
        if row['Application'] in prob_apps and row['Application'] not in list(problem_df['Application']):
#             print(row['Application'], dashboard[(dashboard['Application'] == row['Application']) & (dashboard['# of unique installations'] > 100)]['Publisher'])
            problem_df = problem_df.append({'Application': row['Application'], 'Publishers': str(list(dashboard[(dashboard['Application'] == row['Application']) & (dashboard['# of unique installations'] > 100)]['Publisher']))[1:-1].replace("'", '')}, ignore_index=True)

problem_df.to_excel('problematic_apps_07_06.xlsx')

print("Identifying problematic applications complete.")