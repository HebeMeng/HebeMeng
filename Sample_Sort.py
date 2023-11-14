
import pandas as pd
import numpy as np
import sys
import math


def Sample_Sort(lab_data, bof, time_delay):

    lab_data = lab_data.sort_values('Start Time')
    lab_data = lab_data.reset_index(drop=True)
    col_del_duplicate = 'Start Time'
    lab_data = lab_data.drop_duplicates(subset=col_del_duplicate, keep='first')
    lab_data = lab_data.reset_index(drop=True)

    lab_data['Start Time'] = pd.to_datetime(lab_data['Start Time'])
    lab_data['Finish Time'] = pd.to_datetime(lab_data['Finish Time'])

    # Calculate the time gap to prevent incorrect lab data time
    gap = (lab_data['Start Time'] - lab_data['Finish Time'].shift(1)).to_frame('time_gap')
    gap['time_gap'] = gap['time_gap'].dt.total_seconds() / 60

    wrong_index = gap[gap['time_gap'] < 0].index.to_list()
    if len(wrong_index) != 0:
        for index in wrong_index:
            print('The start time of ' + str(lab_data.loc[index, 'Start Time']) + ' in the lab data is overlapping with the previous one!')
            input('Please check the lab data to redo the process!')
            sys.exit()

    sample_time_delay = time_delay
    Batch_cnt = lab_data.shape[0]

    Batch_digit = math.floor(math.log(Batch_cnt) / math.log(10) + 1)
    lab_data['Batch'] = 'No analyser data'

    sys.stdout.write('Sorting the bof according to lab data sample time...\n')
    update_frequency = 2
    time_col = bof.iloc[:, 1] - pd.Timedelta(minutes=sample_time_delay)

    for j in range(Batch_cnt):
        # if j== 37:
        #     print('pause')

        start = lab_data.iloc[j, 0]
        finish = lab_data.iloc[j, 1]

        mask = (time_col > start) & (time_col <= finish)
        sample_indices = np.where(mask)[0]  # Find matching samples indices

        if len(sample_indices) > 0:
            batch_label = "Sample" + str(j + 1).zfill(Batch_digit)
            bof.loc[sample_indices, 'Batch'] = batch_label
            lab_data.loc[j, 'Batch'] = batch_label

        # Calculate and print progress
        progress = int((j + 1) / Batch_cnt * 100)
        if progress % update_frequency == 0:
            sys.stdout.write("\r{:.2f}%".format(progress))
            sys.stdout.flush()

    bof = bof.reset_index(drop=True)

    return bof, lab_data


# # test code
# import pandas as pd
# from read_bof import read_bof
# bof = read_bof('//scantech-adc1/scantech/Client Analysers/TBL-230/TBL-068 JKN Logistics-Kazakhmys/S01-Technical/Calibration/2023-09-04-Calibration HG/0 - TBL068 values RV.bof')
# Lab_data = pd.read_excel('//scantech-adc1/scantech/Client Analysers/TBL-230/TBL-068 JKN Logistics-Kazakhmys/S01-Technical/Calibration/2023-09-04-Calibration HG/1 - whole LabData.xlsx', header=1)
# Lab_header = Lab_data.columns.to_list()
# Lab_header[0] = 'Start Time'
# Lab_header[1] = 'Finish Time'
# Lab_data.columns = Lab_header
# time_delay = 0
# [processed_bof, processed_lab_data] = Sample_Sort(Lab_data, bof, time_delay)
# print(processed_bof)
