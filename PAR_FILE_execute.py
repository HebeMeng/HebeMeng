

def IF(expression, exec_0, exec_1, exec_2=None, exec_3=None, exec_4=None, exec_5=None, exec_6=None, exec_7=None, exec_8=None, exec_9=None):
    try:
        result = int(expression)
        if result < 0 or result > 9:
            raise ValueError("Expression out of range. Just support 0-9.")
    except ValueError:
        print("Invalid expression. Just support 0-9.")
        return

    actions = [exec_0, exec_1, exec_2, exec_3, exec_4, exec_5, exec_6, exec_7, exec_8, exec_9]
    action = actions[result]

    if action is None:
        print(f"No action provided for expression: {result}")
        return

    return action


def test():
    def action0():
        return "Action 0 executed"

    def action1():
        return "Action 1 executed"

    print(IF("0", action0, action1))  # Should print "Action 0 executed"
    print(IF("1", action0, action1))  # Should print "Action 1 executed"
    print(IF("10", action0, action1))  # Should print "Invalid expression"
    print(IF("-1", action0, action1))  # Should print "Invalid expression"
    print(IF("2", action0, action1))  # Should print "No action provided for expression: 2"


# test()
# A = 2
# C = 1
# B = IF(A, 0 * C, 1 * C, 2 * C)
# print(B)  # Output: 2


def par_run(variable_name):

    return


# import re
#
# text = "R030 = IF(S030, R031 * 2 + 1.23, R032)"
# pattern = r'\b[A-Z]+\d+\b'
#
# variables = re.findall(pattern, text)
# print(variables)
#
# pattern = r'\b\d+\.\d+|\b\d+\b'
# coefficients = re.findall(pattern, text)
# print(coefficients)
#
# # text = "R030 = IF(S030, R031 * 2, R032)"
# # pattern = r'\b(?!IF\b)[A-Z]+\d+\b'
#
# variables = re.findall(pattern, text)
# print(variables)

import pandas as pd



# Convert the columns to datetime

