
import pandas as pd
from scipy.stats import pearsonr
import numpy as np
import warnings
import sys


def Cal_Correlation(lab_data, avg_bof):
    element_cnt = lab_data.shape[1] - 4
    bof_col_cnt = avg_bof.shape[1] - 3
    corr_LabPar = np.zeros((element_cnt, bof_col_cnt))
    warnings.filterwarnings("ignore")

    log_messages = []  # Collect log messages in a list

    sys.stdout.write('\n\nCalculating correlation of averaged analysis data and lab data...')

    for j in range(bof_col_cnt):
        corr = 0
        col_avg = avg_bof.iloc[:, j + 3]

        if len(col_avg.unique()) == 1:
            log_messages.append(f"\n{avg_bof.columns[j + 3]} in bof are same, thus its corr ratio will output 0.")
            continue

        if col_avg.isna().any():
            log_messages.append(f"\n{avg_bof.columns[j + 3]} in bof contains nan, thus its corr ratio will output 0.")
            continue

        else:
            col_avg = pd.to_numeric(col_avg, errors='coerce')

        for i in range(element_cnt):
            col = lab_data.iloc[:, i + 2]
            if col.isna().any():
                log_messages.append('\nLab data ' + lab_data.columns[i + 2] + 'contains nan')
                log_messages.append('\n\nPlease check the lab data and then run the code again!\n\n')
                sys.exit(1)

            col_lab = pd.to_numeric(col, errors='coerce')

            try:
                corr, _ = pearsonr(col_lab, col_avg)
            except ValueError:
                corr = 'N/A'
                log_messages.append(f"Error calculating correlation for columns {i+2} and {j+3}")

            corr_LabPar[i, j] = corr

    df_element_name = pd.DataFrame({"ElementName": lab_data.columns[2:-2].transpose()})
    df_corr_Data = pd.DataFrame(data=corr_LabPar, columns=avg_bof.columns[3:])
    df_corr_Data = pd.concat([df_element_name, df_corr_Data], axis=1)

    log_string = ''.join(log_messages)  # Join log messages into a single string
    sys.stdout.write(log_string)

    return df_corr_Data, log_string



# # test code
# import pandas as pd
# import numpy as np
# from read_bof import read_bof
# from Sample_Sort import Sample_Sort
# from Product_Average import Product_Average
# bof = read_bof('J:/Client Analysers/On Belt Analyser/OBA-264 Nova CimAngola - Sinoma/S01-Technical/Calibration/230810 Cal Review DR/03 CL.bof')
# Lab_data = pd.read_excel('J:/Client Analysers/On Belt Analyser/OBA-264 Nova CimAngola - Sinoma/S01-Technical/Calibration/230810 Cal Review DR/02 fixed lab data.xlsx', header=1)
# time_delay = 0
# [Processed_bof, processed_lab_data] = Sample_Sort(Lab_data, bof, time_delay)
# bof_in_sample = Processed_bof[Processed_bof['Batch'] != 'Out of Sample']
# [avg_bof, processed_lab_data, log] = Product_Average(bof_in_sample, 'S835', 'S829', processed_lab_data)
# df_valid_labdata = processed_lab_data[processed_lab_data['Batch'] != 'No analyser data']
# df_valid_labdata = df_valid_labdata.reset_index(drop=True)
# [corr_data, log] = Cal_Correlation(df_valid_labdata, avg_bof)
# print(corr_data)
