import pandas as pd
import scipy.stats as stats


def CochranC_test(data):

    d2 = data ** 2
    dmax2 = d2.max()
    max_index = d2.idxmax()
    sum_dev_2 = d2.sum()
    c_value = dmax2 / sum_dev_2
    dfd = d2.shape[0]
    q = 1 - 0.01 / dfd
    dfn = 2
    f_cric = stats.f.ppf(q, dfn - 1, (dfd - 1) * (dfn - 1))
    c_cric = pow((1 + (dfd - 1) / f_cric), -1)
    critical_value = round(c_cric, 4)

    # Identify outlier
    if c_value > critical_value:
        outlier_index = max_index

    else:
        outlier_index = 'NO'

    return outlier_index


def mark_outlier(data):
    i = 1
    outlier_index = True
    outlier_index_list =[]
    while outlier_index != 'NO':
        data_left = data.loc[data.index >= 0]
        outlier_index = CochranC_test(data_left)
        if outlier_index != 'NO':
            outlier_index_list.append(outlier_index)
        data.rename(index={outlier_index: -1 * i}, inplace=True)
        i += 1

    return outlier_index_list

# import os
# os.chdir('J:/Client Analysers/Analyser PSA Report/Dynamic Calibration Templates/calibration python code HG/outlier_debug_test')
# lab_data = pd.read_excel('lab.xlsx')
# analyser_data = pd.read_excel('analyser.xlsx')
# dif = lab_data['SI'] - analyser_data['SI']
# x = mark_outlier(dif)
# print(x)
