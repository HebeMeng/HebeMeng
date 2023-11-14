import pandas as pd
import numpy as np
import sys


def Product_Average(bof, analysed_ton_name, not_analysed_ton_name, lab_data):

    # bof is the bof in sample
    # analysed_ton_name is the Symbol in PAR file that means analysed ton. Like S034 or S835
    # not_analysed_ton_name is the Symbol in PAR file that means not analysed ton. Like S029
    # lab_data is the processed lab data. So you can get the lab data sample ID with analysis data

    log_messages = []
    sys.stdout.write('\n\nCalculating the average data for each sample...')
    sample_list = [x for x in lab_data['Batch'].to_list() if x != 'No analyser data']
    lab_data['Rows of bof'] = 0
    valid_batch_cnt = len(sample_list)
    avg_header = bof.columns[2:].to_list()
    avg_header.insert(1, 'sum of analysed ton')
    df_avg = pd.DataFrame(np.nan, index=range(valid_batch_cnt), columns=avg_header)
    df_avg['Batch'] = df_avg['Batch'].astype(object)

    # drop the empty rows in advance
    nan_percentage = bof.isnull().sum(axis=1) / len(bof.columns)
    threshold = 0.8  # if the row contains over 0.8 nan data, drop this row
    rows_to_drop = bof[nan_percentage > threshold].index.to_list()
    if len(rows_to_drop) != 0:
        bof = bof.drop(rows_to_drop)

        log_messages.append(f'\nBof in sample sheet row ' + str(
            rows_to_drop) + ' +3 contains too many nan, so it is dropped in advance!')

    bof = bof.reset_index(drop=True)

    # FileName and Analyse_time not suitable for df_avg
    pure_average_flag = False
    if analysed_ton_name == '':
        pure_average_flag = True

        log_messages.append(f'\nBof does not contain ton column..\nSo the averaged bof is not based on analysed ton.')

    elif bof[analysed_ton_name].isna().any():
        pure_average_flag = True

        log_messages.append(f'\nBof does not contain proper analysed ton column..\nSo the averaged bof is not based on analysed ton.')

    else:
        analysed_ton_index = bof.columns.get_loc(analysed_ton_name)

    if not_analysed_ton_name == '':
        log_messages.append(f'\nLacking not analysed ton, so bof rows that contain high not analysed ton will also be averaged.')

    col_cnts = bof.shape[1]
    i = 0
    for sample_id in sample_list:
        # if sample_id == 'Sample52':
        #     print('pause here')
        temp_df = bof[bof['Batch'] == sample_id]
        lab_data.loc[lab_data['Batch'] == sample_id, 'Rows of bof'] = temp_df.shape[0]
        if pure_average_flag:
            temp_avg = temp_df.iloc[:, 3:].mean(axis=0).tolist()
            temp_avg.insert(0, 'wrong ton')
            temp_avg.insert(0, sample_id)
            df_avg.iloc[i, :] = temp_avg
            i += 1
            continue

        nan_percentage = temp_df.isnull().sum(axis=1) / len(temp_df.columns)
        threshold = 0.8  # if the row contains over 0.8 nan data, drop this row
        rows_to_drop = temp_df[nan_percentage > threshold].index.to_list()
        if len(rows_to_drop) != 0:
            temp_df = temp_df.drop(rows_to_drop)
            lab_data.loc[lab_data['Batch'] == sample_id, 'Rows of bof'] = temp_df.shape[0]

            log_messages.append(f'\n' + sample_id + ' in sheet of bof in sample row ' + str(
                rows_to_drop) + ' +3 contains too many nan, so it is not averaged!')

        non_analysed_ton_ratio = pd.DataFrame(index=temp_df.index)
        if not_analysed_ton_name != '':
            non_analysed_ton_ratio = temp_df[not_analysed_ton_name].div(temp_df[analysed_ton_name])
            rows_to_drop = non_analysed_ton_ratio[non_analysed_ton_ratio > 0.3].index.to_list()
        else:
            rows_to_drop = []

        if len(rows_to_drop) != 0:
            temp_df = temp_df.drop(rows_to_drop)
            log_messages.append(f'\n' + sample_id + ' in sheet of bof in sample row ' + str(
                rows_to_drop) + ' + 3 contains too high ratio of non analysed ton, so it is not averaged!')

        # how to drop some rows properly
        if len(temp_df) != 0:
            analysed_ton = temp_df.iloc[:, analysed_ton_index].values
            data = temp_df.iloc[:, 3:]
            temp_avg = (np.matmul(data.transpose().to_numpy(), analysed_ton) / sum(analysed_ton)).tolist()
            temp_avg.insert(0, sum(analysed_ton))
            temp_avg.insert(0, sample_id)
            df_avg.iloc[i, :] = temp_avg

            i += 1
        else:
            log_messages.append(f'\nNo qualified data in ' + sample_id + ', so it is not producing average data!')
            lab_data.loc[lab_data['Batch'] == sample_id, ['Batch', 'Rows of bof']] = ['No analyser data', 0]
            df_avg.drop(i, inplace=True)

    valid_lab_data = lab_data[lab_data['Batch'] != 'No analyser data'].reset_index(drop=True)

    df_avg.insert(1, 'BOF rows in use', valid_lab_data.loc[:, 'Rows of bof'])

    log_string = ''.join(log_messages)
    sys.stdout.write(log_string)

    return df_avg, lab_data, log_string


#
# # test code
# import pandas as pd
# import numpy as np
# from read_bof import read_bof
# from Sample_Sort import Sample_Sort
#
# bof = read_bof('J:/Client Analysers/TBL-230/TBL-068 JKN Logistics-Kazakhmys/S01-Technical/Calibration/2023-09-04-Calibration HG/0 - TBL068 values RV.bof')
# Lab_data = pd.read_excel('J:/Client Analysers/TBL-230/TBL-068 JKN Logistics-Kazakhmys/S01-Technical/Calibration/2023-09-04-Calibration HG/1 - prod 1 - Shatyrkul -LabData.xlsx', header=1)
# time_delay = 0
# [Processed_bof, processed_lab_data] = Sample_Sort(Lab_data, bof, time_delay)
# bof_in_sample = Processed_bof[Processed_bof['Batch'] != 'Out of Sample']
# [avg_bof, lab_data, log] = Product_Average(bof_in_sample, 'S034', '', processed_lab_data)
# print('\n')
# print(avg_bof)



