## Setup and running the code
Both the dashboards and the bundling algorithm can be found in the notebook `full_dashboard.ipynb` and the code can be run in Jupyter notebooks. 
Install Python 3.9
Use `pip` (built-in Python package management system) or other libraries to install the following libraries: 
- Jupyter Notebook: run `pip install jupyterlab`, see https://jupyter.org/install for more details
- Pandas 1.2.3: run `pip install pandas==1.2.3`, see https://pandas.pydata.org/pandas-docs/stable/getting_started/install.html
- FuzzyWuzzy 0.18.0: run `pip install fuzzywuzzy==0.18.0`, see https://pypi.org/project/fuzzywuzzy/
- NumPy 1.20.0: run `pip install numpy`, see https://numpy.org/install/

Navigate to the notebooks/ folder in command prompt or terminal and run ‘jupyter notebook’ to open up the notebook on your local machine

The data should be downloaded as excel file(s) and placed in the data/ folder of the project repository.

## Running the full Python script

First, open full-script.py in your text editor and change the filepath on line 12 to the directory where you are storing the data.

After installing all of the required packages, go to your Terminal and navigate to the directory where the script is held. Then, type `python full-script.py` into your Terminal.

The script should perform the following processes:
1. Reading in data
2. Removing servers
3. Removing GOTS applications
4. Flagging utility applications (outputting one csv with the full dataset flagged and another csv with just the utility applications)
5. Normalizing application names
6. Creating a dashboard without versions
7. Creating a dashboard with versions
8. Bundling applications
9. Identifying problematic applications
