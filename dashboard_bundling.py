import pandas as pd
import glob
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import numpy as np
import csv

# This creates an empty dataframe
data = pd.DataFrame(columns=['Publisher', 'Application', 'Version', 'Install Date', 'Last HW Scan', 'OS', 'OS Version', 'Encrypted Workstation Name'])

# The filepath needs to be where the files are located on your device
file_paths = glob.glob('/Users/jeremylee/Desktop/CS-Projects/T4SG/usda-app-rationalization-summer/data/*.xlsx')

# Reading in the data
# This might take some time, but you'll be able to see the filenames as they're being read in
for filepath in file_paths:
    data = data.append(pd.read_excel(filepath))
    print(filepath)
print('Done')

### Remove all the rows of data from Microsoft servers, as we’re only interested in workstations
data['C054'].fillna('',inplace=True)  # replaces NaN values in column C054 with ‘’ nothing, needed for next segment of code to work
index = data[data['C054'].str.lower().str.contains('server')].index # creates an index of rows containing server in the OS column (C054)
data.drop(index, inplace=True)  # this deletes all rows where the word 'server' is in column C054, from prior line's index

# Changing the setting so we could see all the lines of the dashboard
pd.set_option('display.max_rows', None)

# Dashboard by application name and number of installations in the whole dataset
full_data = data.groupby(['Publisher', 'Application'])
full_data = full_data.size().sort_values( ascending=False).reset_index(name='# of all entries')
dashboard = pd.DataFrame(full_data)

### Creating a dashboard without versions

# Dashboard by application name and number of duplicated installations in the whole dataset
duplicates = data.where(data.duplicated(subset=['Application', 'Encrypted Workstation Name', 'Last HW Scan']) == True).dropna(how='all')
dupl_dash = duplicates.groupby(['Publisher', 'Application'])
dupl_dash = dupl_dash.size().sort_values( ascending=False).reset_index(name='# of duplicate installations')
dupl_df = pd.DataFrame(dupl_dash, columns = ['Publisher', 'Application', '# of duplicate installations'])

# Dashboard by application name and number of not duplicated installations in the whole dataset
data_no_dupl = data.where(data.duplicated(subset=['Application', 'Encrypted Workstation Name', 'Last HW Scan']) != True).dropna(how='all')
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


### Creating dashboards that DO take into account the version number


# Creating a version based duplicates dashboard
version_data = data.copy()
version_data['Version'] = version_data['Version'].astype(str)

# Full dashboard for applications with versions and the number of all installations
v_full_data = version_data.groupby(['Publisher', 'Application', 'Version'])
v_full_data = v_full_data.size().sort_values( ascending=False).reset_index(name='# of all entries')
v_dashboard = pd.DataFrame(v_full_data)

# Finding duplicates by versions
v_duplicates = version_data.where(version_data.duplicated(subset=['Application', 'Version', 'Encrypted Workstation Name', 'Last HW Scan']) == True).dropna(how='all')
v_dupl_dash = v_duplicates.groupby(['Publisher', 'Application', 'Version'])
v_dupl_dash = v_dupl_dash.size().sort_values( ascending=False).reset_index(name='# of duplicate installations')

# Finding unique installations by versions
v_no_dupl = version_data.where(version_data.duplicated(subset=['Application', 'Version', 'Encrypted Workstation Name', 'Last HW Scan']) != True).dropna(how='all')
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


### Bundle-identification code


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
    flist1 = filtered_1["Encrypted Workstation Name"].to_list()
    flist2 = filtered_2["Encrypted Workstation Name"].to_list()
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
    print("Done")

# Creates a csv of bundles from the v_full dashboard (defined above)
createBundle(v_full)


### Identify problematic applications that have typos in publisher name


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
