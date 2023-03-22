import pandas as pd
import os
from tqdm import tqdm_gui as tqdm
from sqlalchemy import create_engine
import glob
from pandas import ExcelWriter
import datetime

##
## Change Directory where your python credentials are stored
os.chdir('C:\\Users\\Username\\Documents')

## Connect to your database via:
exec(open("C:\\Users\\Username\\Documents\\passcodes\\credentials.py").read())

# lists of queries locations, could be a shared Drive if working with others!
root_folder = 'S:\\Analytics\\Reporting\\Filename\\SQL'

# here is where you can include multiple subfolders, repeat line 23
folder_list = [
                root_folder + '\\SQLSubFolderWhereSQLStored'
                ]

# Get date for last month for use in all filenames
today = datetime.date.today()
first = today.replace(day=1)
lastmonth = (first - datetime.timedelta(days=1)).strftime("%Y.%m")

for folder in folder_list:

    # find only sql queries by joining the path name in folder list
    for file in glob.glob(os.path.join(folder, '*.sql')):
        # open every sql file in the folder list and read the sql query
        with open(file, 'r') as f:
            read = f.read()
            print(f'Reading file {str(os.path.basename(file))}')
            df = pd.read_sql_query(read, con=con)

            # Change to directory where file will be saved
            os.chdir(folder)

            # variables for file naming conventions
            base_folder_name = str(os.path.basename(folder))
            filename = base_folder_name + ' ' + lastmonth + '.xlsx'

            # write executed dataframe with the name of the base folder with last months date
            writer = ExcelWriter(filename, engine='xlsxwriter',datetime_format="mm/dd/yyyy")
            df.to_excel(writer, 'Sheet 1', index=False)
            writer.save()
            print('Saving...')