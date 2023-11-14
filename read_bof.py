import pandas as pd
import re
import warnings
import sys


def read_bof(file):

    with open(file, 'r') as f:
        i = 0
        if file.endswith(".bof"):
            start_read = False
            for line in f:
                i += 1
                if 'FileName' in line:
                    start_read = True
                if start_read == True:
                    line = line.strip()
                    header_line = re.split('[,\t]', line)
                    file_header = [elem.replace('\t', '') for elem in header_line]
                    file_header = [elem.replace(' ', '') for elem in file_header]
                    break

            header = [elem for elem in file_header if elem != ""]

            warnings.filterwarnings("ignore")

            col = pd.Index(header)

            df = pd.read_csv(file, sep="\s*,\s*|\t", lineterminator='\n', header=None, names=col, index_col=False,
                             skiprows=i)
            warnings.simplefilter("default")

        elif file.endswith(".csv"):
            df = pd.read_csv(file)

        # delete all the non-numeric columns in the data
        # df_numeric = df.select_dtypes(include='number')  # select numeric columns
        # new_cols = pd.DataFrame(df.iloc[:, 0], columns=['FileName'])  # create new DataFrame with FileName column
        # df_numeric = pd.concat([new_cols, df_numeric], axis=1)  # concatenate DataFrames
        #return df_numeric

        sys.stdout.write('Getting the time from the bof file...\n')
        bof = df.copy()
        date_time_strs = bof.iloc[:, 0]
        # Split the strings into separate date and time parts
        df[['date', 'time']] = date_time_strs.str.extract(r'(\d{2}-\d{2}-\d{2}) (\d{2}.\d{2}.\d{2})')

        # Add "20" prefix to the year part of the date strings
        date_strs = "20" + df['date']
        time_strs = df['time']
        # Convert the date and time strings to datetime format
        date_time_strs = date_strs + " " + time_strs
        # Convert the datetime Series to numpy array
        analyse_time = pd.to_datetime(date_time_strs, format='%Y-%m-%d %H.%M.%S')

        bof.insert(1, 'Analyse Time', analyse_time)
        bof.insert(2, 'Batch', 'Out of Sample')
        bof = bof.sort_values('Analyse Time')
        bof = bof.reset_index(drop=True)
        cols_to_convert = bof.iloc[0:, 3:].columns
        bof[cols_to_convert] = bof[cols_to_convert].apply(pd.to_numeric, errors='coerce')

        return bof

#
# bof_file = "P:/3 - python test code(learning)/write test/0 - 090 All Oxides IN.bof"
# bof = read_bof(bof_file)
# print(bof)
