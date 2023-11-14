import re
import pandas as pd


def read_PAR_symbols(filename):
    with open(filename, 'r') as file:
        found_calibration = False
        symbol_section = []

        for line in file:
            if not found_calibration:
                if "[CalibrationSymbols]" in line:
                    found_calibration = True
            elif not line.strip():
                continue
            elif "Symbol_" not in line:
                break
            else:
                symbol_section.append(line.strip())

    pattern = r"= "
    value = []
    for line in symbol_section:
        match = re.search(pattern, line)
        if match:
            value.append(line[match.start() + 2:])
    l = len(value)
    i = 0
    var_tbl = pd.DataFrame(index=range(l), columns=['ParmID', 'Meaning'])
    for ele in value:
        [var_tbl.iloc[i, 0], var_tbl.iloc[i, 1]] = ele.split(',')
        i += 1
    return var_tbl



# result = tt.loc[tt['Meaning'].str.contains('Dyn'), 'Meaning']
#
# print(tt.iloc[result.index[0],0])


def read_PAR_formula(filename):
    found_calib_equations = False
    formula_list: list[str] = []
    i = 0
    with open(filename, 'r') as f:
        for line in f:
            if found_calib_equations:
                if not line.strip():
                    continue
                elif "=" in line:
                    formula_list.append(line.strip())

                elif "Expression_" not in line:
                    break

            elif line.strip() == '[CalibrationEquations]':
                found_calib_equations = True
    element_dict = {}
    for element in formula_list:
        sub_elements = element.split()  # Split the element into sub-elements
        element_dict[sub_elements[2]] = sub_elements[4:]

    return element_dict


def read_file_without_empty_lines(file_path):
    with open(file_path, 'r') as file:
        lines = [line.strip() for line in file if line.strip()]
    return lines


def get_formula_layer(my_dict):
    dict_layer_group = []
    remaining_dict = my_dict
    source_keys = list(my_dict.keys())
    assigned_keys = []
    remaining_keys = source_keys
    while len(remaining_keys) != 0:
        dict_layer_new = {}
        keys = list(remaining_dict.keys())
        flag_found = False
        for key in remaining_dict:
            if flag_found:
                break
            for element in remaining_dict[key]:
                if element in keys:
                    if '/' not in remaining_dict[key]:
                        flag_found = True
                        break
            else:
                dict_layer_new[key] = my_dict[key]

        dict_layer_group.append(dict_layer_new)
        assigned_keys.extend(list(dict_layer_new.keys()))
        remaining_keys = [x for x in source_keys if x not in assigned_keys]
        remaining_dict = {key: my_dict[key] for key in remaining_keys}
    return dict_layer_group



#
# search_key = ['R030', 'R031', 'R032']  # belong to FSA_ID
# BLC_R = 'S028'  # the last column header of the lab data
# R_tuple = ()
# BLC_Slope_tuple = ()
# for sub_search in search_key:
#     R_revise = []  # record the R numbers in the bof that need to be recalculated with the assigned BLC and static weight
#     BLC_slope = []  # record the BLC slope that correspond one by one to the R_2revise list
#     found_value = ' '
#     target = [sub_search]
#     relevant_R = target  # record for each search key which R values or S values are correlated with this search_key, it is a tempeary value that only fit to the newest search
#     searched = []  # record for each search loop, which R numbers have already been searched and don't repeat
#     end_flag = 0  # when search reaches the second layer, stops for the first layer is not BLC dependent
#     found = False
#     for ele in target:
#         if end_flag == 0:
#
#             for dictionary in layer_group:
#                 if ele in dictionary:
#                     searched.append(ele)
#                     found = True
#                     found_value = dictionary[ele]
#                     RXXX = [element for element in found_value if 'R' in element or 'S' in element]
#                     if BLC_R in RXXX:
#
#                         R_revise.append(ele)
#                         BLC_index = found_value.index(BLC_R)
#                         if BLC_index - 3 >= 0:
#                             if found_value[BLC_index - 1] == '*':
#                                 str_slope = (found_value[BLC_index - 3]+found_value[BLC_index - 2])
#                                 BLC_slope.append(float(str_slope))
#                         if BLC_index + 3 < len(found_value):
#                             if found_value[BLC_index + 1] == '*':
#                                 str_slope = (found_value[BLC_index + 2]+found_value[BLC_index + 3])
#                                 BLC_slope.append(float(str_slope))
#                         elif BLC_index + 2 < len(found_value):
#                             if found_value[BLC_index + 1] == '*':
#                                 BLC_slope.append(float(found_value[BLC_index + 2]))
#
#                     relevant_R.extend(RXXX)
#                     if dictionary == layer_group[1]:
#                         end_flag = 1
#
#                     break
#
#             target = [ele for ele in relevant_R if ele not in searched]
#         else:
#             break
#
#     R_tuple += (R_revise,)
#     BLC_Slope_tuple += (BLC_slope,)
# end_t = time.time()
# dt = end_t - start_t
# print(dt)
# for i in range(len(R_tuple)):
#     for j in range(len(R_tuple[i])-1, -1, -1):
#         temp_str = R_dict[R_tuple[i][j]]
#         new_temp = []
#         for element in temp_str:
#             if 'R' in element or 'S' in element:
#                 new_temp.append('bof["' + element + '"]')
#             else:
#                 new_temp.append(element)
#
#         print('bof["' + str(R_tuple[i][j]) + '"]' + '=' + ''.join(new_temp))
#         # print('temp = ' + "bof['" + str(R_tuple[i][j]) + "'] + bof['" +str(BLC_R) + "'] * (" + str(BLC_Slope_tuple[i][j]) + ')')
#         # print("bof['" + str(R_tuple[i][j]) + "'] = temp")

# Example usage
# filename = "J:/Client Analysers/CS-2100/C21-140 PT Asmin Bara Bronang/S01-Technical/Calibration/2023-10-25 Ash/CS2120.par"  # Replace with your file name
# tt = read_PAR_symbols(filename)
# print(tt)
# R_dict = read_PAR_formula(filename)
# print(R_dict)
# layer_group = get_formula_layer(R_dict)
# print(layer_group)