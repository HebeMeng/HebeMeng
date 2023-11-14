import sys

from datetime import datetime
import os
import win32com.client as win32
import numpy as np
import pandas as pd
import xlwings as xw
from PyQt6.QtGui import QFont
from PySide6.QtCore import *
from PySide6.QtGui import *
from PySide6.QtWidgets import *
from read_bof import read_bof
from read_PAR import read_PAR_symbols
from read_PAR import read_PAR_formula
from read_PAR import get_formula_layer
from Sample_Sort import Sample_Sort
from Product_Average import Product_Average
from Cal_Correlation import Cal_Correlation


class Ui_calWindow(object):

    def __init__(self):
        self.delay_rangeLow = None
        self.spinBox_rangeHigh = None
        self.text_static = None
        self.checkBox_corr = None
        self.line_delay = None
        self.spinBox_rangeLow = None
        self.lineEdit_static = None
        self.checkBox_static = None
        self.pTE_PAR = None
        self.PAR = None
        self.pTE_lab = None
        self.pTE_bof = None
        self.static_calibration = False
        self.pushButton_bof = None
        self.pushButton_PAR = None
        self.lineEdit_4 = None
        self.lineEdit_2 = None
        self.lineEdit = None
        self.pushButton_close = None
        self.font1 = None
        self.spinBox = None
        self.signvalue = 0
        self.minutes = 0
        self.pushButton_lab = None
        self.pushButton_ok = None
        self.comboBox = None
        self.lab_file = None
        self.bof_file = None
        self.flag = 0
        self.flag_min = 0
        self.static_calibration = False
        self.best_corr_search = False  # mark whether to try to find a best correlation in a loop
        self.delay_rangeLow = 0  # set the range of possible delay in the search
        self.delay_rangeHigh = 0
        self.time_step = 1  # set the time step in searching best correlation

    def setupUi(self, calWindow):
        if not calWindow.objectName():
            calWindow.setObjectName(u"calWindow")
        calWindow.setEnabled(True)
        calWindow.resize(465, 490)
        calWindow.setAutoFillBackground(True)
        self.pushButton_bof = QPushButton(calWindow)
        self.pushButton_bof.setObjectName(u"pushButton_bof")
        self.pushButton_bof.setGeometry(40, 30, 140, 40)
        font = QFont('Arial', 10)
        fontS = QFont('Arial', 8)
        fontB = QFont('Arial', 10)
        fontB.setBold(True)
        fontBL = QFont('Arial', 12)
        fontBL.setBold(True)
        self.pushButton_bof.setFont(fontB)
        self.pTE_bof = QPlainTextEdit(calWindow)
        self.pTE_bof.setObjectName(u"pTE_bof")
        self.pTE_bof.setGeometry(190, 30, 240, 40)
        self.pushButton_lab = QPushButton(calWindow)
        self.pushButton_lab.setObjectName(u"pushButton_lab")
        self.pushButton_lab.setGeometry(40, 90, 140, 40)
        self.pushButton_lab.setFont(fontB)
        self.pTE_lab = QPlainTextEdit(calWindow)
        self.pTE_lab.setObjectName(u"pTE_lab")
        self.pTE_lab.setGeometry(190, 90, 240, 40)
        self.pushButton_PAR = QPushButton(calWindow)
        self.pushButton_PAR.setObjectName(u"pushButton_PAR")
        self.pushButton_PAR.setFont(fontB)
        self.pushButton_PAR.setGeometry(40, 150, 140, 40)
        self.pTE_PAR = QPlainTextEdit(calWindow)
        self.pTE_PAR.setObjectName(u"pTE_lab")
        self.pTE_PAR.setGeometry(190, 150, 240, 40)
        self.pushButton_ok = QPushButton(calWindow)
        self.pushButton_ok.setObjectName(u"pushButton_ok")
        self.pushButton_ok.setGeometry(100, 400, 140, 40)
        self.pushButton_ok.setFont(fontB)
        self.pushButton_close = QPushButton(calWindow)
        self.pushButton_close.setObjectName(u"pushButton_close")
        self.pushButton_close.setGeometry(260, 400, 140, 40)
        self.pushButton_close.setFont(fontB)
        self.pushButton_close.setEnabled(False)
        self.comboBox = QComboBox(calWindow)
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.setObjectName(u"comboBox")
        self.comboBox.setGeometry(260, 230, 61, 22)
        self.comboBox.setFont(fontB.toString())
        self.comboBox.setEditable(True)
        self.lineEdit = QLineEdit(calWindow)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setEnabled(False)
        self.lineEdit.setGeometry(40, 230, 110, 20)
        self.lineEdit.setFont(font)
        self.lineEdit.setFrame(False)
        self.lineEdit_2 = QLineEdit(calWindow)
        self.lineEdit_2.setObjectName(u"lineEdit_2")
        self.lineEdit_2.setEnabled(False)
        self.lineEdit_2.setGeometry(200, 230, 50, 20)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setAutoFillBackground(True)
        self.lineEdit_2.setFrame(False)
        self.lineEdit_4 = QLineEdit(calWindow)
        self.lineEdit_4.setObjectName(u"lineEdit_4")
        self.lineEdit_4.setEnabled(False)
        self.lineEdit_4.setGeometry(330, 230, 120, 20)
        self.lineEdit_4.setFont(font)
        self.lineEdit_4.setFrame(False)
        self.spinBox = QSpinBox(calWindow)
        self.spinBox.setObjectName(u"spinBox")
        self.spinBox.setGeometry(150, 230, 45, 20)
        self.spinBox.setMaximum(1000)
        self.spinBox.setSingleStep(1)
        self.spinBox.setValue(0)
        self.checkBox_static = QCheckBox(calWindow)
        self.checkBox_static.setObjectName(u"checkBox_static")
        self.checkBox_static.setGeometry(40, 270, 200, 20)
        self.checkBox_static.setFont(fontB)
        self.checkBox_static.setEnabled(False)
        self.checkBox_corr = QCheckBox(calWindow)
        self.checkBox_corr.setObjectName(u"checkBox_corr")
        self.checkBox_corr.setGeometry(QRect(200, 270, 200, 20))
        self.checkBox_corr.setFont(fontB)
        self.text_static = QTextEdit(calWindow)
        self.text_static.setObjectName(u"text_static")
        self.text_static.setEnabled(False)
        self.text_static.setGeometry(40, 300, 140, 70)
        self.text_static.setFont(fontB)
        self.text_static.setAutoFillBackground(True)
        self.text_static.setFrameShape(QFrame.NoFrame)
        self.text_static.setFont(font)
        self.text_static.setReadOnly(True)
        self.line_delay = QLineEdit(calWindow)
        self.line_delay.setObjectName(u"line_delay")
        self.line_delay.setEnabled(False)
        self.line_delay.setGeometry(197, 300, 240, 20)
        self.line_delay.setFont(font)
        self.line_delay.setFrame(False)
        self.line_delay_2 = QLineEdit(calWindow)
        self.line_delay_2.setObjectName(u"line_delay_2")
        self.line_delay_2.setEnabled(False)
        self.line_delay_2.setGeometry(QRect(300, 325, 251, 31))
        self.line_delay_2.setFont(font)
        self.line_delay_2.setFrame(False)

        self.spinBox_rangeLow = QSpinBox(calWindow)
        self.spinBox_rangeLow.setObjectName(u"spinBox_range")
        self.spinBox_rangeLow.setGeometry(330, 300, 42, 22)
        self.spinBox_rangeLow.setInputMethodHints(Qt.ImhDigitsOnly | Qt.ImhHiddenText)
        self.spinBox_rangeLow.setWrapping(True)
        self.spinBox_rangeLow.setFrame(True)
        self.spinBox_rangeLow.setVisible(False)
        self.spinBox_rangeLow.setMinimum(-1000)
        self.spinBox_rangeLow.setMaximum(1000)
        self.spinBox_rangeLow.setSingleStep(1)
        self.spinBox_rangeLow.setValue(0)

        self.spinBox_rangeHigh = QSpinBox(calWindow)
        self.spinBox_rangeHigh.setObjectName(u"spinBox_range")
        self.spinBox_rangeHigh.setGeometry(330, 330, 42, 22)
        self.spinBox_rangeHigh.setInputMethodHints(Qt.ImhDigitsOnly | Qt.ImhHiddenText)
        self.spinBox_rangeHigh.setWrapping(True)
        self.spinBox_rangeHigh.setFrame(True)
        self.spinBox_rangeHigh.setVisible(False)
        self.spinBox_rangeHigh.setMinimum(-1000)
        self.spinBox_rangeHigh.setMaximum(1000)
        self.spinBox_rangeHigh.setSingleStep(1)
        self.spinBox_rangeHigh.setValue(0)

        self.lineEdit_step = QLineEdit(calWindow)
        self.lineEdit_step.setObjectName(u"lineEdit_step")
        self.lineEdit_step.setEnabled(False)
        self.lineEdit_step.setGeometry(216, 360, 250, 20)
        self.lineEdit_step.setFont(font)
        self.lineEdit_step.setFrame(False)
        self.spinBox_step = QSpinBox(calWindow)
        self.spinBox_step.setObjectName(u"spinBox_step")
        self.spinBox_step.setVisible(False)
        self.spinBox_step.setGeometry(330, 360, 42, 22)
        self.spinBox_step.setMaximum(120)
        self.spinBox_step.setMinimum(1)
        self.spinBox_step.setSingleStep(1)
        self.spinBox_step.setValue(0)
        self.retranslateUi(calWindow)
        self.checkBox_corr.setText(QCoreApplication.translate("calWindow", u"Try to find best correlation", None))
        self.line_delay.setText('')
        self.lineEdit_step.setText('')
        # retranslateUi

        # Connect the slot function to the released signal of the pushButton
        self.pushButton_bof.released.connect(self.on_pushButton_bof_released)
        self.pushButton_lab.released.connect(self.on_pushButton_lab_released)
        self.pushButton_PAR.released.connect(self.on_pushButton_PAR_released)
        self.pushButton_ok.released.connect(self.on_pushButton_ok_released)
        self.comboBox.activated.connect(self.comboBox_setCurrentIndex)
        self.spinBox.valueChanged.connect(self.on_spin_box_value_changed)
        self.pushButton_close.released.connect(self.on_pushButton_close_released)
        self.checkBox_static.stateChanged.connect(self.checkBox_static_statechange)
        self.checkBox_corr.stateChanged.connect(self.checkBox_findbest_corr)
        self.spinBox_rangeLow.valueChanged.connect(self.set_delay_rangeLow)
        self.spinBox_rangeHigh.valueChanged.connect(self.set_delay_rangeHigh)
        self.spinBox_step.valueChanged.connect(self.set_time_step)

        QMetaObject.connectSlotsByName(calWindow)

        # setupUi

    def retranslateUi(self, calWindow):
        calWindow.setWindowTitle(QCoreApplication.translate("calWindow", "Calibration Setup V3", disambiguation=""))
        icon = QIcon()
        icon.addFile(u"cool-dude.png", QSize(120, 120), QIcon.Normal, QIcon.Off)
        calWindow.setWindowIcon(icon)
        self.pushButton_bof.setText(QCoreApplication.translate("calWindow", u"Select the BOF file", disambiguation=""))
        self.pushButton_lab.setText(QCoreApplication.translate("calWindow", u"Select the Lab Data", None))
        self.pushButton_ok.setText(QCoreApplication.translate("calWindow", u"CONFIRM", None))
        self.comboBox.setItemText(0, QCoreApplication.translate("calWindow", u"select", None))
        self.comboBox.setItemText(1, QCoreApplication.translate("calWindow", u"earlier", None))
        self.comboBox.setItemText(2, QCoreApplication.translate("calWindow", u"later", None))
        self.comboBox.setCurrentText(QCoreApplication.translate("calWindow", u"select", None))
        self.lineEdit.setText(QCoreApplication.translate("calWindow", u"Sampling time is", None))
        self.lineEdit_2.setText(QCoreApplication.translate("calWindow", u"minutes", None))
        self.lineEdit_4.setText(QCoreApplication.translate("calWindow", u"than Analyzing time", None))
        self.pushButton_close.setText(QCoreApplication.translate("calWindow", u"RUN", None))
        self.pushButton_PAR.setText(QCoreApplication.translate("calWindow", u"Select the PAR file", None))
        self.checkBox_static.setText(QCoreApplication.translate("calWindow", u"static calibration", None))
        self.text_static.setText("")
        self.checkBox_corr.setText(QCoreApplication.translate("calWindow", u"Try to find best correlation", None))
        self.line_delay.setText("")
        self.line_delay_2.setText("")
        self.lineEdit_step.setText("")

    # retranslateUi

    def on_pushButton_bof_released(self):
        self.bof_file, _ = QFileDialog.getOpenFileNames(QWidget(), "Select BOF File", "",
                                                        "BOF Files (*.bof);;All Files (*.*)")
        temp = self.bof_file
        bof_name = [temp[i][temp[i].rfind("/") + 1:] for i in range(len(temp))]
        delimiter = '; '
        self.pTE_bof.setPlainText(delimiter.join(bof_name))

        # If a file was selected, print its path to the console
        if self.bof_file:
            return self.bof_file

    def on_pushButton_PAR_released(self):
        self.PAR, _ = QFileDialog.getOpenFileName(QWidget(), "Select PAR File", "",
                                                  "PAR Files (*.par);;All Files (*.*)")
        temp = self.PAR
        index = temp.rfind("/")
        PAR_name = temp[index + 1:]
        self.pTE_PAR.setPlainText(PAR_name)
        if self.pTE_PAR:
            return self.pTE_PAR

    def on_pushButton_lab_released(self):
        self.lab_file, _ = QFileDialog.getOpenFileName(QWidget(), "Select Lab Data File", "",
                                                       "Lab Data Files (*.xlsx);;All Files (*.*)")
        temp = self.lab_file
        index = temp.rfind("/")
        lab_name = temp[index + 1:]
        self.pTE_lab.setPlainText(lab_name)
        # If a file was selected, print its path to the console
        if self.lab_file:
            return self.lab_file

    def on_pushButton_ok_released(self):
        if self.minutes != 0 and self.signvalue == 0:
            sys.stdout.write('\nPlease choose earlier or later, or set the minutes to be 0!\n')
        else:
            if self.flag == 0:
                self.signvalue = str(0)
            if self.flag_min == 0:
                self.minutes = str(0)
            if self.bof_file and self.lab_file and self.PAR:
                self.run_calibration()
            fontB = QFont('Arial', 10)
            fontB.setBold(True)
            self.pTE_bof.setReadOnly(True)
            self.pTE_bof.setBackgroundVisible(False)
            self.pTE_bof.setFont(fontB.toString())
            self.pTE_lab.setReadOnly(True)
            self.pTE_lab.setBackgroundVisible(False)
            self.pTE_lab.setFont(fontB.toString())
            self.pTE_PAR.setReadOnly(True)
            self.pTE_PAR.setBackgroundVisible(False)
            self.pTE_PAR.setFont(fontB.toString())
            self.spinBox.setReadOnly(True)
            self.comboBox.setEnabled(False)
            self.checkBox_static.setEnabled(False)
            self.checkBox_corr.setEnabled(False)
            self.pushButton_lab.setEnabled(False)
            self.pushButton_bof.setEnabled(False)
            self.pushButton_close.setEnabled(True)
            self.pushButton_ok.setEnabled(False)
            self.pushButton_PAR.setEnabled(False)

    def checkBox_static_statechange(self):
        if self.checkBox_static.isChecked():  # Qt.CheckState.Checked:
            self.text_static.setText('Please confirm the lab data file contains weight data')
            self.static_calibration = True
            self.checkBox_corr.setEnabled(False)
            print("static calibration checked, doesn't support multi delay time")
        else:
            self.text_static.setText('')
            self.static_calibration = False
            self.checkBox_corr.setEnabled(True)
            print('static calibration unchecked')

    def checkBox_findbest_corr(self):
        if self.checkBox_corr.isChecked():
            self.line_delay.setText('Delay time range is from                   minutes')
            self.line_delay_2.setText('To                  minutes')
            self.lineEdit_step.setText('Set time step as                  minutes')
            self.spinBox_step.setVisible(True)
            self.spinBox_rangeLow.setVisible(True)
            self.spinBox_rangeHigh.setVisible(True)
            self.checkBox_static.setEnabled(False)
            self.best_corr_search = True
            self.spinBox.setEnabled(False)
            self.comboBox.setEnabled(False)
            self.signvalue = 0

            print(
                'Will output multi delay results on each delay time in different sheets, please decide based on your '
                'judge')
        else:
            self.line_delay.setText('')
            self.line_delay_2.setText('')
            self.lineEdit_step.setText('')
            self.spinBox_step.setVisible(False)
            self.spinBox_rangeLow.setVisible(False)
            self.spinBox_rangeHigh.setVisible(False)
            self.checkBox_static.setEnabled(True)
            self.best_corr_search = False
            self.spinBox.setEnabled(True)
            self.comboBox.setEnabled(True)
            print('Will output single result based on the time delay')

    def on_spin_box_value_changed(self):
        self.minutes = self.spinBox.value()
        self.flag_min = 1

    def comboBox_setCurrentIndex(self):
        self.signvalue = self.comboBox.currentIndex()
        self.flag = 1

    def set_delay_rangeLow(self):
        self.delay_rangeLow = self.spinBox_rangeLow.value()

    def set_delay_rangeHigh(self):
        self.delay_rangeHigh = self.spinBox_rangeHigh.value()

    def set_time_step(self):
        self.time_step = self.spinBox_step.value()

    def run_calibration(self):
        if self.bof_file and self.lab_file:
            # Pass the filenames as command-line arguments to the other Python script
            param1 = self.bof_file
            param2 = self.lab_file
            param3 = self.minutes
            param4 = self.signvalue
            param5 = self.PAR
            param6 = self.static_calibration
            param7 = self.time_step
            param8 = self.delay_rangeLow
            param9 = self.delay_rangeHigh
            param10 = self.best_corr_search
            return param1, param2, param3, param4, param5, param6, param7, param8, param9, param10

    def on_pushButton_close_released(self):
        window.close()


# Create the application
app = QApplication(sys.argv)

# Create the main window
window = QMainWindow()

# Create an instance of the Ui_calWindow class and set up the user interface
ui = Ui_calWindow()
ui.setupUi(window)

window.show()

status = app.exec()
parm_list = ui.run_calibration()
# print(parm_list)

# sys.exit(status)
[param1, param2, param3, param4, param5, param6, param7, param8, param9, param10] = parm_list

bof_file = param1  # full location of bof file
lab_file = param2  # full location of lab file
sample_time_delay = int(param3)  # the time unit of argv is minutes
sign_value = int(param4)  # it gets 1, means sample time is earlier, if it gets 2, means sample time is later.
if sign_value == 1:
    sample_time_delay = sample_time_delay
elif sign_value == 2:
    sample_time_delay = -1 * sample_time_delay
else:
    sample_time_delay = 0

par_file = param5  # full location of par file
static_calibration = param6  # Boolean type
time_step = param7  # time step in calculating the best corr
time_rangeLow = param8  # time delay start value
time_rangeHigh = param9  # time delay finish value
corr_search = param10  # Boolean type

bof_cnt = len(bof_file)
index = bof_file[0].rfind("/")
path = bof_file[0][:index]
delimiter = '; '
bof_names = delimiter.join(bof_file)

if not os.path.exists(path):
    print('Exit for some wrong inputs')
    import time

    time.sleep(6)
    sys.exit()

if sign_value == 1:
    sign_symbol = 1
    time_note = "_Sample(earlier)"
elif sign_value == 2:
    sign_symbol = -1
    time_note = "_Sample(later)"
else:
    sign_symbol = 0
    time_note = "_Sample(same_time)"

# read lab data
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
print('Reading the lab data...')
os.chdir(path)
df_lab = pd.read_excel(lab_file, header=None)
sample_ID = df_lab.iloc[0, 2:].values
Lab_data = pd.read_excel(lab_file, header=1)
Lab_header = Lab_data.columns.to_list()
Lab_header[0] = 'Start Time'
Lab_header[1] = 'Finish Time'
Lab_data.columns = Lab_header

# read bof
print('Reading bof file...')
if bof_cnt == 1:
    bof = read_bof(bof_file[0])
else:
    bof_group = {}
    for i in range(bof_cnt):
        bof_group[i] = read_bof(bof_file[i])
    headers = bof_group[0].columns

    same_headers = all(df.columns.equals(headers) for df in list(bof_group.values()))
    if same_headers:
        bof = pd.concat(bof_group, ignore_index=True)
        col_del_duplicate = 'FileName'
        bof = bof.drop_duplicates(subset=col_del_duplicate, keep='first')
        bof = bof.reset_index(drop=True)

    else:
        input('These bof files have different headers!')
        sys.exit(0)

# check whether the sample ID are all included in the bof
bof_headers = bof.columns
correct_header_flag = all(ele in bof_headers for ele in sample_ID)
if not correct_header_flag:
    print('\nSample ID in the lab data contain out of bof variables, please check!\n')
    sys.exit(0)

# read PAR file to get the calibration symbol table
print('Reading the PAR file...')
var_tbl = read_PAR_symbols(par_file)
if static_calibration:
    R_dict = read_PAR_formula(par_file)
    layer_group = get_formula_layer(R_dict)

# set an excel files to save in each step
now = datetime.now()
date_without_year = datetime.now().date().strftime("%m-%d")
current_t = str(date_without_year) + " " + str(now.hour).zfill(2) + str(now.minute).zfill(2)
output_file = current_t + " Calibration.xlsx"

log_string_1 = 'This calibration result is based on the following files:\n' + bof_names + '\n' + lab_file + '\n' + par_file
if sign_symbol == 1:
    txt = 'considered ' + str(abs(sample_time_delay)) + ' minutes earlier than the analyser time'
elif sign_symbol == -1:
    txt = 'considered ' + str(abs(sample_time_delay)) + ' minutes later than the analyser time'
elif sign_symbol == 0:
    txt = 'considered the same time as the analyser time'
else:
    txt = 'unrecorded'

log_string_2 = '\n\nThe sampling time is ' + txt
log_name = output_file + '.log'
with open(log_name, "w") as file:
    file.write(log_string_1)
    file.write(log_string_2)

# Sorting bof based on lab data
if not corr_search:
    print('Output excel file name will be ' + output_file[:-1] + 'm')
    print('The log file with the same name records the input parameters in case you need to trace back.')
    try:
        [processed_bof, processed_lab_data] = Sample_Sort(Lab_data, bof, sample_time_delay)
        bof_in_sample = processed_bof[processed_bof['Batch'] != 'Out of Sample']
        bof_in_sample = bof_in_sample.reset_index(drop=True)

        # getting different ton names based on different CsSchedule version
        bof_header = bof_in_sample.columns
        not_analysed_ton_name = ''
        analysed_ton_name = ''

        if 'S835' in bof_header and 'S828' in bof_header:
            analysed_ton_name = "S835"
            weight_name = 'S828'
            if 'S829' in bof_header:
                not_analysed_ton_name = 'S829'
            else:
                response = input(
                    '\nNo not analysed ton in bof. Suggest to re extract bof.\nEnter 0 to continue anyway; Or press '
                    'other button to re extract bof.\n ').strip()
                if response != '0':
                    sys.exit()
        elif 'S034' in bof_header and 'S028' in bof_header:
            analysed_ton_name = 'S034'
            weight_name = 'S028'
            if 'S029' in bof_header:
                not_analysed_ton_name = 'S029'
            else:
                response = input(
                    '\nNo not analysed ton in bof. Suggest to re extract bof.\nEnter 0 to continue anyway; Or press '
                    'other button to re extract bof.\n ').strip()
                if response != '0':
                    sys.exit()
        elif 'S803' in bof_header and 'S906' in bof_header:
            analysed_ton_name = 'S803'
            weight_name = 'S906'
            if 'S821' in bof_header:
                not_analysed_ton_name = 'S821'
            else:
                response = input(
                    '\nNo not analysed ton in bof. Suggest to re extract bof.\nEnter 0 to continue anyway; Or press '
                    'other button to re extract bof.\n ').strip()
                if response != '0':
                    sys.exit()
        else:
            print('\n\nbof does not contain analysed ton data or belt load data')
            print('Suggest to re extract bof...')
            response = input('\nEnter 0 to continue anyway; Or press other button to re extract bof.\n ').strip()
            if response != '0':
                sys.exit()

        # calculate the averaged results
        [avg_bof, processed_lab_data, add_log] = Product_Average(bof_in_sample, analysed_ton_name,
                                                                 not_analysed_ton_name, processed_lab_data)

        avg_bof = avg_bof.reset_index(drop=True)
        processed_lab_data = processed_lab_data.reset_index(drop=True)

        with open(log_name, "a") as file:
            file.write(add_log)
        del add_log

        # calculate the correlation of data
        valid_lab_data = processed_lab_data[processed_lab_data['Batch'] != 'No analyser data']
        valid_lab_data = valid_lab_data.reset_index(drop=True)
        # Do the outlier test for the difference
        outlier_flag = input('\n\nPlease enter 1 to do the outlier test and 0 to pass. \n')

        df1 = pd.DataFrame(data=avg_bof[sample_ID])
        df1.columns = valid_lab_data.columns[2: -2]
        df2 = pd.DataFrame(data=valid_lab_data.iloc[:, 2:-2])
        dif_data = df1 - df2
        outlier_table = pd.DataFrame(columns=dif_data.columns.to_list())
        if int(outlier_flag) == 1:
            with open(log_name, "a") as file:
                file.write('\n\nThe outlier for difference between lab data and analyser data is processed.\n')

            from CochranC_module import mark_outlier
            from CochranC_module import CochranC_test
            for col in dif_data.columns.to_list():
                outlier_sample = 'no outlier'
                test_data = dif_data[col]
                index = mark_outlier(test_data)
                if len(index) != 0:
                    outlier_sample = ','.join(valid_lab_data.loc[index, 'Batch'])
                outlier_table.loc[0, col] = outlier_sample

        # calculate the correlation ratios
        [corr_data, add_log] = Cal_Correlation(valid_lab_data, avg_bof)
        with open(log_name, "a") as file:
            file.write(add_log)
        del add_log

        output_bof_in_sample = bof_in_sample.iloc[:, 2:]

        # add meaning to the output data
        col_cnt = processed_bof.shape[1]
        row_meaning = ['Undefined'] * col_cnt
        row_meaning[0] = 'meaning'
        for i in range(1, col_cnt):
            if processed_bof.columns[i] in var_tbl.iloc[:, 0].to_list():
                row_meaning[i] = var_tbl[var_tbl.iloc[:, 0] == processed_bof.columns[i]].iloc[:, 1].to_list()[0]
        output_processed_bof = processed_bof.copy()
        output_processed_bof.loc[-1] = row_meaning
        output_processed_bof.index = output_processed_bof.index + 1
        output_processed_bof = output_processed_bof.sort_index()

        col_cnt = output_bof_in_sample.shape[1]
        row_meaning = ['Undefined'] * col_cnt
        row_meaning[0] = 'meaning'
        for i in range(1, col_cnt):
            if output_bof_in_sample.columns[i] in var_tbl.iloc[:, 0].to_list():
                row_meaning[i] = var_tbl[var_tbl.iloc[:, 0] == output_bof_in_sample.columns[i]].iloc[:, 1].to_list()[0]
        print_bof_in_sample = output_bof_in_sample.copy()
        print_bof_in_sample.loc[-1] = row_meaning
        print_bof_in_sample.index = print_bof_in_sample.index + 1
        print_bof_in_sample = print_bof_in_sample.sort_index()

        col_cnt = avg_bof.shape[1]
        row_meaning = ['Undefined'] * col_cnt
        row_meaning[0] = 'meaning'
        for i in range(1, col_cnt):
            if avg_bof.columns[i] in var_tbl.iloc[:, 0].to_list():
                row_meaning[i] = var_tbl[var_tbl.iloc[:, 0] == avg_bof.columns[i]].iloc[:, 1].to_list()[0]
        output_avg_bof = avg_bof.copy()
        output_avg_bof.loc[-1] = row_meaning
        output_avg_bof.index = output_avg_bof.index + 1
        output_avg_bof = output_avg_bof.sort_index()

        col_cnt = corr_data.shape[1]
        row_meaning = ['Undefined'] * col_cnt
        row_meaning[0] = 'meaning'

        for i in range(1, col_cnt):
            if corr_data.columns[i] in var_tbl.iloc[:, 0].to_list():
                row_meaning[i] = var_tbl[var_tbl.iloc[:, 0] == corr_data.columns[i]].iloc[:, 1].to_list()[0]
        output_corr_data = corr_data.copy()
        output_corr_data.loc[-1] = row_meaning
        output_corr_data.index = output_corr_data.index + 1
        output_corr_data = output_corr_data.sort_index()

        # write the output dataframe into Excel
        os.chdir(path)
        print("\n")
        print("Writing to Excel file: " + output_file[:-1] + 'm')
        print("Please be patient if the bof is large")
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            output_processed_bof.to_excel(writer, sheet_name='Processed bof', index=False)
            print_bof_in_sample.to_excel(writer, sheet_name='Bof in sample', index=False)
            output_avg_bof.to_excel(writer, sheet_name='Batch Average', index=False)
            processed_lab_data.to_excel(writer, sheet_name='Lab Data', index=False)
            output_corr_data.to_excel(writer, sheet_name='Correlation Ratio', index=False)
            if int(outlier_flag) == 1:
                outlier_table.to_excel(writer, sheet_name='Outlier table for reference', index=False)

        del output_corr_data
        del output_avg_bof
        del output_processed_bof
        del output_bof_in_sample
        del print_bof_in_sample

    except Exception as e:
        print(f"An error occurred: {e}")

    # generating the Calibration sheet
    os.chdir(path)
    wb = xw.Book(output_file)

    lab_data_in_batch = processed_lab_data[processed_lab_data['Batch'] != 'No analyser data']
    lab_data_in_batch = lab_data_in_batch.reset_index(drop=True)
    element_cnt = lab_data_in_batch.shape[1] - 4

    if 'weight_name' in locals():
        lab_data_in_batch.insert(element_cnt + 2, 'Belt Load', avg_bof.loc[:, weight_name])

    sheet = wb.sheets.add(name='Calibration')
    sheet.range('A1').options(index=False).value = "Lab Data"
    sheet.range('A6').options(index=False).value = lab_data_in_batch
    col_gap = element_cnt + 2

    sheet.range(1, col_gap + 5).options(index=False).value = "Analyser Data"
    bof_header = avg_bof.columns
    intersection_sample_bof = [item for item in sample_ID if item in bof_header]
    sheet.range(6, col_gap + 5).options(index=False).value = avg_bof.loc[:, intersection_sample_bof]
    cnt_FSA = len(intersection_sample_bof)
    cnt_sample = avg_bof.shape[0] - 1
    sheet.range(1, col_gap * 2 + 5).options(index=False).value = "Original Difference"
    sheet.range(6, col_gap * 2 + 5).options(index=False).value = np.array(lab_data_in_batch.columns[2:-2]).transpose()
    sheet.range(3, col_gap * 2 + 4).options(index=False).value = "StdDev"
    sheet.range(4, col_gap * 2 + 4).options(index=False).value = "Bias"
    sheet.range(1, col_gap * 3 + 5).options(index=False).value = "Analyser Data Corrected"
    sheet.range(3, col_gap * 3 + 4).options(index=False).value = "Slope"
    sheet.range(4, col_gap * 3 + 4).options(index=False).value = "Offset/Intercept"
    sheet.range(5, col_gap * 3 + 4).options(index=False).value = "BLC"
    sheet.range(6, col_gap * 3 + 5).options(index=False).value = np.array(lab_data_in_batch.columns[2:-2]).transpose()
    sheet.range(1, col_gap * 4 + 5).options(index=False).value = "Corrected Difference"
    sheet.range(6, col_gap * 4 + 5).options(index=False).value = np.array(lab_data_in_batch.columns[2:-2]).transpose()
    sheet.range(3, col_gap * 4 + 4).options(index=False).value = "StdDev"
    sheet.range(4, col_gap * 4 + 4).options(index=False).value = "Bias"
    for i in range(cnt_FSA):
        sheet.range(3, col_gap * 3 + 5 + i).value = 1
        sheet.range(4, col_gap * 3 + 5 + i).value = 0
        sheet.range(5, col_gap * 3 + 5 + i).value = 0
    wb.save()
    wb.close()

    # save the previous .xlsx file to xlsm file with macro code

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    full_name = path + '/' + output_file
    workbook = excel.Workbooks.Open(full_name)
    fName = os.path.splitext(workbook.FullName)[0] + '.xlsm'
    workbook.SaveAs(Filename=fName, FileFormat=52)  # 52 is for xlOpenXMLWorkbookMacroEnabled
    workbook.Close(True)

    # delete the .xlsx file after it save as an xlsm file
    os.remove(full_name)

    # Define the VBA macro code as a string

    # Connect to an existing Excel application or create a new one
    xl_app = win32.Dispatch('Excel.Application')

    # Connect to an existing workbook or create a new one
    wb = xl_app.Workbooks.Open(fName)
    # Define the macro code

    macro_code = '''
        Sub Calibration_process(E_cnt As Integer, B_cnt As Integer)
    
            'Set ws = ActiveWorkbook.Worksheets("Lab Data")
            'E_cnt = ws.UsedRange.Columns.Count - 4
            'Set ws = ActiveWorkbook.Worksheets("Batch Average")
            'B_cnt = ws.UsedRange.Rows.Count - 2
    
            ActiveWorkbook.Sheets("Calibration").Activate
            Columns("A:A").EntireColumn.AutoFit
            Columns("B:B").EntireColumn.AutoFit
            Cells(7, 3).Resize(B_cnt, E_cnt + 1).NumberFormat = "0.00"
            Cells(7, E_cnt + 7).Resize(B_cnt, E_cnt).NumberFormat = "0.00"
            Col_1 = (E_cnt + 2) * 2 + 5
            Cells(7, Col_1).FormulaR1C1 = "=RC[" & (-1 * ((E_cnt + 2) * 2 + 2)) & "]-RC[" & -1 * (E_cnt + 2) & "]"
            Cells(7, Col_1).AutoFill Destination:=Cells(7, Col_1).Resize(B_cnt, 1), Type:=xlFillDefault
            If E_cnt > 1 Then
                Range(Cells(7, Col_1), Cells(B_cnt + 6, Col_1)).AutoFill Destination:=Cells(7, Col_1).Resize(B_cnt, E_cnt), Type:=xlFillDefault
            End If
            Cells(7, Col_1).Resize(B_cnt, E_cnt).NumberFormat = "0.00"
    
            Cells(3, Col_1).FormulaR1C1 = "=STDEV.P(R[4]C:R[" & B_cnt + 3 & "]C)"
            If E_cnt > 1 Then
                Cells(3, Col_1).AutoFill Destination:=Cells(3, Col_1).Resize(1, E_cnt), Type:=xlFillDefault
            End If
            Cells(3, Col_1).Resize(1, E_cnt).NumberFormat = "0.00"
    
            Cells(4, Col_1).FormulaR1C1 = "=AVERAGE(R[3]C:R[" & B_cnt + 2 & "]C)"
            If E_cnt > 1 Then
                Cells(4, Col_1).AutoFill Destination:=Cells(4, Col_1).Resize(1, E_cnt), Type:=xlFillDefault
            End If
            Cells(4, Col_1).Resize(1, E_cnt).NumberFormat = "0.00"
    
            'Revise the Analyser Data
            Col_2 = (E_cnt + 2) * 3 + 5
            Cells(7, Col_2).FormulaR1C1 = "=R3C*RC[-" & (E_cnt + 2) * 2 & "]+R4C+R5C*RC" & E_cnt + 3
            Cells(7, Col_2).AutoFill Destination:=Cells(7, Col_2).Resize(B_cnt, 1), Type:=xlFillDefault
            If E_cnt > 1 Then
                Range(Cells(7, Col_2), Cells(B_cnt + 6, Col_2)).AutoFill Destination:=Cells(7, Col_2).Resize(B_cnt, E_cnt), Type:=xlFillDefault
            End If
            Cells(7, Col_2).Resize(B_cnt, E_cnt).NumberFormat = "0.00"
    
            'Calculate the revised difference
            Col_3 = (E_cnt + 2) * 4 + 5
            Cells(7, (E_cnt + 2) * 4 + 5).FormulaR1C1 = "=RC[" & (-1 * ((E_cnt + 2) * 4 + 2)) & "]-RC[" & -1 * (E_cnt + 2) & "]"
            Cells(7, (E_cnt + 2) * 4 + 5).AutoFill Destination:=Cells(7, (E_cnt + 2) * 4 + 5).Resize(B_cnt, 1), Type:=xlFillDefault
            If E_cnt > 1 Then
                Range(Cells(7, (E_cnt + 2) * 4 + 5), Cells(B_cnt + 6, (E_cnt + 2) * 4 + 5)).AutoFill Destination:=Cells(7, (E_cnt + 2) * 4 + 5).Resize(B_cnt, E_cnt), Type:=xlFillDefault
            End If
            Cells(7, (E_cnt + 2) * 4 + 5).Resize(B_cnt + 1, E_cnt).NumberFormat = "0.00"
    
            Cells(3, (E_cnt + 2) * 4 + 5).FormulaR1C1 = "=STDEV.P(R[4]C:R[" & B_cnt + 3 & "]C)"
            If E_cnt > 1 Then
                Cells(3, (E_cnt + 2) * 4 + 5).AutoFill Destination:=Cells(3, (E_cnt + 2) * 4 + 5).Resize(1, E_cnt), Type:=xlFillDefault
            End If
            Cells(3, (E_cnt + 2) * 4 + 5).Resize(1, E_cnt).NumberFormat = "0.00"
    
            Cells(4, (E_cnt + 2) * 4 + 5).FormulaR1C1 = "=AVERAGE(R[3]C:R[" & B_cnt + 2 & "]C)"
            If E_cnt > 1 Then
                Cells(4, (E_cnt + 2) * 4 + 5).AutoFill Destination:=Cells(4, (E_cnt + 2) * 4 + 5).Resize(1, E_cnt), Type:=xlFillDefault
            End If
            Cells(4, (E_cnt + 2) * 4 + 5).Resize(1, E_cnt).NumberFormat = "0.00"
    
            ' Create the correlation matrix
            Rng = ConvertToRangeFormat(B_cnt + 8, 1)
            Range(Rng).Value = "CorrelationMatrix"
    
            Range("C6").Select
            Selection.Resize(1, E_cnt + 1).Copy
            Rng = ConvertToRangeFormat(B_cnt + 9, 1)
            Range(Rng).Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
    
    
            For i = 0 To E_cnt - 1 'represent col
                For j = 0 To E_cnt 'represent row
                Rng = ConvertToRangeFormat(B_cnt + 9 + j, (E_cnt + 2) * 3 + 5 + i)
                row_dif = Trim(Str(B_cnt + 2 + j))
                col_dif = Trim(Str(3 * E_cnt + 8 - j + i))
                Range(Rng).FormulaR1C1 = "=CORREL(R[-" & row_dif & "]C:R[" & Str(-3 - j) & "]C,R[-" & row_dif & "]C[-" & col_dif & "]:R[" & Str(-3 - j) & "]C[-" & col_dif & "])"
                Range(Rng).NumberFormat = "0.00"
                Next j
            Next i
    
            'Create the chart
    
            Dim lineCharts(0 To 100) As ChartObject
            Dim xyChart(0 To 100) As ChartObject
    
            chrt_width_1 = Cells(B_cnt + 8, (E_cnt + 2) * 3 + 12).Width * 11
            chrt_height_1 = Cells(B_cnt + 8, (E_cnt + 2) * 3 + 12).Height * 20
            chrt_width_2 = Cells(B_cnt + 8, (E_cnt + 2) * 3 + 12).Width * 8
            chrt_height_2 = Cells(B_cnt + 8, (E_cnt + 2) * 3 + 12).Height * 20
    
            For i = 0 To E_cnt - 1
    
                Set lineCharts(i) = ActiveSheet.ChartObjects.Add( _
                Left:=ActiveSheet.Cells(B_cnt + E_cnt + 11 + 20 * i, (E_cnt + 2) * 3 + 5).Left, _
                Top:=ActiveSheet.Cells(B_cnt + E_cnt + 11 + 20 * i, (E_cnt + 2) * 3 + 5).Top, _
                Width:=chrt_width_1, _
                Height:=chrt_height_1 _
                )
                lineCharts(i).Chart.ChartType = xlLine
                lineCharts(i).Select
                ActiveChart.ApplyLayout (1)
                ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = ""
                ActiveChart.ChartType = xlLineMarkers
                ActiveChart.Axes(xlCategory).Select
                Selection.Delete
                With lineCharts(i)
    
                    .Chart.SetSourceData Source:=Union(Cells(7, (E_cnt + 2) * 3 + 5 + i).Resize(B_cnt, 1), Cells(7, 3 + i).Resize(B_cnt, 1))
                    .Chart.FullSeriesCollection(1).Name = "=""LabData"""
                    .Chart.FullSeriesCollection(1).Format.Line.ForeColor.RGB = RGB(190, 75, 72)
                    .Chart.FullSeriesCollection(1).MarkerBackgroundColor = RGB(190, 75, 72)
                    .Chart.FullSeriesCollection(1).Format.Line.Weight = 1
                    .Chart.FullSeriesCollection(2).Name = "=""AnalyserData"""
                    .Chart.FullSeriesCollection(2).Format.Line.ForeColor.RGB = RGB(74, 126, 187)
                    .Chart.FullSeriesCollection(2).MarkerBackgroundColor = RGB(74, 126, 187)
                    .Chart.FullSeriesCollection(2).Format.Line.Weight = 1
                    .Chart.ChartTitle.Text = ActiveSheet.Cells(6, 3 + i).Value
                    .Chart.Legend.Position = xlTop
                End With
                Set myseries = lineCharts(i).Chart.FullSeriesCollection(1)
                myseries.MarkerSize = 3
                Set myseries = lineCharts(i).Chart.FullSeriesCollection(2)
                myseries.MarkerSize = 3
    
            Next i
            For i = 0 To E_cnt - 1
    
                Set xyChart(i) = ActiveSheet.ChartObjects.Add( _
                Left:=ActiveSheet.Cells(B_cnt + E_cnt + 11 + 20 * i, (E_cnt + 2) * 3 + 16).Left, _
                Top:=ActiveSheet.Cells(B_cnt + E_cnt + 11 + 20 * i, (E_cnt + 2) * 3 + 16).Top, _
                Width:=chrt_width_1, _
                Height:=chrt_height_1 _
                )
                xyChart(i).Chart.ChartType = xlXYScatter
                xyChart(i).Select
                ActiveChart.ApplyLayout (1)
    
                With xyChart(i)
                    .Chart.SetSourceData Source:=Union(Cells(7, 3 + i).Resize(B_cnt, 1), Cells(7, (E_cnt + 2) * 3 + 5 + i).Resize(B_cnt, 1))
    
                    .Chart.Axes(xlCategory).MinimumScale = WorksheetFunction.Floor_Math(WorksheetFunction.Max(0, WorksheetFunction.Min(Cells(7, 3 + i).Resize(B_cnt, 1)) * 0.95))
                    .Chart.Axes(xlCategory).MaximumScale = WorksheetFunction.Ceiling_Math(WorksheetFunction.Max(Cells(7, 3 + i).Resize(B_cnt, 1)) * 1.05)
    
                    .Chart.Axes(xlValue).MinimumScale = .Chart.Axes(xlCategory).MinimumScale
                    .Chart.Axes(xlValue).MaximumScale = .Chart.Axes(xlCategory).MaximumScale
                    .Chart.Axes(xlValue).MinimumScaleIsAuto = False
                    .Chart.Axes(xlValue).MaximumScaleIsAuto = False
                    .Chart.Axes(xlCategory).AxisTitle.Text = "Lab Data"
                    .Chart.Axes(xlValue).AxisTitle.Text = "Analyser Data"
                    .Chart.ChartTitle.Text = ActiveSheet.Cells(6, 3 + i).Value
                    .Chart.Legend.Delete
    
                End With
    
                'Creation of the line x=y, the perfect fit line
                Set mychart = ActiveChart
                mychart.FullSeriesCollection(1).MarkerSize = 3
                mychart.Shapes.AddLine(ActiveChart.PlotArea.InsideLeft, _
                    mychart.PlotArea.InsideHeight + ActiveChart.PlotArea.InsideTop, _
                    mychart.PlotArea.InsideWidth + ActiveChart.PlotArea.InsideLeft, _
                    mychart.PlotArea.InsideTop).Name = "xy_line"
                mychart.SetElement (msoElementPrimaryCategoryGridLinesMajor)
    
            Next i
    
    
           'Freeze the window
            ActiveWorkbook.Sheets("Processed bof").Activate
            Range("D3").Select
            ActiveWindow.FreezePanes = True
            Columns("A:A").EntireColumn.AutoFit
            Columns("B:B").EntireColumn.AutoFit
            Columns("C:C").EntireColumn.AutoFit
            ActiveWorkbook.Sheets("Bof in sample").Activate
            Range("B3").Select
            ActiveWindow.FreezePanes = True
    
            ActiveWorkbook.Sheets("Lab Data").Activate
            Range("C2").Select
            ActiveWindow.FreezePanes = True
            Columns("A:A").EntireColumn.AutoFit
            Columns("B:B").EntireColumn.AutoFit
            ActiveWorkbook.Sheets("Batch Average").Activate
            Range("B3").Select
            ActiveWindow.FreezePanes = True
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).NumberFormat = "0.00"
            ActiveWorkbook.Sheets("Correlation Ratio").Activate
            Range("B3").Select
            ActiveWindow.FreezePanes = True
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).NumberFormat = "0.00"
    
            ' Highlight the correlation result
            Dim searchValue As String
            Dim foundCell_1 As Range
            Dim foundCell_2 As Range
            Dim columnNumber As Integer
            Dim rowNumber As Integer
            Set ws = ActiveWorkbook.Worksheets("Lab Data")
            E_cnt = ws.UsedRange.Columns.Count - 4
            Set ws = ActiveWorkbook.Worksheets("Batch Average")
            B_cnt = ws.UsedRange.Rows.Count - 2
    
            ' Set the value to search
            For i = 0 To E_cnt - 1
                searchValue_1 = ActiveWorkbook.Sheets("Calibration").Cells(6, E_cnt + 7 + i).Value
                searchValue_2 = ActiveWorkbook.Sheets("Calibration").Cells(6, 3 + i).Value
                ' Find the value in the first row of the active worksheet
                Set foundCell_1 = ActiveWorkbook.Sheets("Correlation Ratio").Rows(1).Find(What:=searchValue_1, LookIn:=xlValues, LookAt:=xlWhole)
                Set foundCell_2 = ActiveWorkbook.Sheets("Correlation Ratio").Columns(1).Find(What:=searchValue_2, LookIn:=xlValues, LookAt:=xlWhole)
    
            ' Check if the value is found
    
                ' Get the column number of the found cell
                columnNumber = foundCell_1.Column
                rowNumber = foundCell_2.Row
                ActiveWorkbook.Sheets("Correlation Ratio").Activate
                Cells(rowNumber, columnNumber).Select
                With Selection.Interior
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
    
            Next i
    
            'Show the chart directly
            ActiveWorkbook.Sheets("Calibration").Activate
            Columns("A:A").EntireColumn.AutoFit
            Columns("B:B").EntireColumn.AutoFit
            Range("C7").Select
            ActiveWindow.FreezePanes = True
            Cells(B_cnt + 8, (E_cnt + 2) * 3 + 12).Select
            'ActiveWindow.WindowState = xlMaximized
    
        End Sub
    
        Function ConvertToRangeFormat(rowNumber As Integer, columnNumber As Integer) As String
            Dim columnLetter As String
            columnLetter = Split(Cells(1, columnNumber).Address, "$")(1)
            ConvertToRangeFormat = columnLetter & rowNumber
        End Function
    
        '''

    # Define the module name
    module_name = 'MyModule'

    # Add the macro code to the workbook
    xl_mod = wb.VBProject.VBComponents.Add(1)  # Add module
    xl_mod.Name = module_name  # Set module name
    xl_mod.CodeModule.AddFromString(macro_code)  # Add code to module

    E_cnt = element_cnt
    B_cnt = lab_data_in_batch.shape[0]

    wb = xw.Book(fName)
    wb.macro('Calibration_process')(E_cnt, B_cnt)
    wb.save()

    # Close the workbook and clear the memory
    del app
    del wb
    del xl_app

    # app = xw.apps.active
    input('Process finished. \nPress Enter to exit...')

if corr_search:
    now = datetime.now()
    date_without_year = datetime.now().date().strftime("%m-%d")
    current_t = str(date_without_year) + " " + str(now.hour).zfill(2) + str(now.minute).zfill(2)
    corr_file = current_t + "corr_compare.xlsx"
    writer = pd.ExcelWriter(corr_file,
                            engine='xlsxwriter')  # This file record the correlation of different time delay

    corr_search_header = [ele for ele in sample_ID]
    corr_search_header.insert(0, 'Delay Time')
    df_corr_search = pd.DataFrame(columns=corr_search_header)

    rng = range(time_rangeLow, time_rangeHigh + time_step, time_step)
    trial_times = len(rng)
    trial_time = 0
    for delay_time in rng:
        trial_time += 1
        sys.stdout.write(
            '\n\n' + str(trial_times) + ' times of trial in total, processing No.' + str(trial_time) + '...\n\n')
        [processed_bof, processed_lab_data] = Sample_Sort(Lab_data, bof, delay_time)
        bof_in_sample = processed_bof[processed_bof['Batch'] != 'Out of Sample']

        # getting different ton names based on different CsSchedule version
        bof_header = bof_in_sample.columns

        not_analysed_ton_name = ''
        analysed_ton_name = ''

        if 'S835' in bof_header and 'S828' in bof_header:
            analysed_ton_name = "S835"
            weight_name = 'S828'
            if 'S829' in bof_header:
                not_analysed_ton_name = 'S829'
            else:
                response = input(
                    '\nNo not analysed ton in bof. Suggest to re extract bof.\nEnter 0 to continue anyway; Or press other button to re extract bof.\n ').strip()
                if response != '0':
                    sys.exit()
        elif 'S034' in bof_header and 'S028' in bof_header:
            analysed_ton_name = 'S034'
            weight_name = 'S028'
            if 'S029' in bof_header:
                not_analysed_ton_name = 'S029'
            else:
                response = input(
                    '\nNo not analysed ton in bof. Suggest to re extract bof.\nEnter 0 to continue anyway; Or press other button to re extract bof.\n ').strip()
                if response != '0':
                    sys.exit()
        elif 'S803' in bof_header and 'S906' in bof_header:
            analysed_ton_name = 'S803'
            weight_name = 'S906'
            if 'S821' in bof_header:
                not_analysed_ton_name = 'S821'
            else:
                response = input(
                    '\nNo not analysed ton in bof. Suggest to re extract bof.\nEnter 0 to continue anyway; Or press other button to re extract bof.\n ').strip()
                if response != '0':
                    sys.exit()
        else:
            print('bof does not contain analysed ton data or belt load data')
            print('Suggest to re extract bof...')
            response = input('Enter 0 to continue anyway; Or press other button to re extract bof.\n ').strip()
            if response != '0':
                sys.exit()

        [avg_bof, processed_lab_data, add_log] = Product_Average(bof_in_sample, analysed_ton_name,
                                                                 not_analysed_ton_name, processed_lab_data)
        valid_lab_data = processed_lab_data[processed_lab_data['Batch'] != 'No analyser data']
        valid_lab_data = valid_lab_data.reset_index(drop=True)

        avg_bof_refer = avg_bof[sample_ID]

        lab_vs_mean_bof = pd.concat([valid_lab_data, avg_bof_refer], axis=1)
        if delay_time < 0:
            vs_sheetname = str(abs(delay_time)) + 'earlier lab vs. average'
        else:
            vs_sheetname = str(abs(delay_time)) + 'later lab vs. average'

        lab_vs_mean_bof.to_excel(writer, sheet_name=vs_sheetname, index=False)

        # make avg_bof_refer to be in same format as the original one to fit the Cal_Correlation function
        avg_bof_refer.insert(0, 'Batch', avg_bof['Batch'])
        avg_bof_refer.insert(1, 'BOF rows in use', avg_bof['BOF rows in use'])
        avg_bof_refer.insert(2, 'sum of analysed ton', avg_bof['sum of analysed ton'])

        [corr_data, add_log] = Cal_Correlation(valid_lab_data, avg_bof_refer)
        corr_size = corr_data.shape[1] - 1
        temp_list = [str(delay_time) + 'minutes']
        for i in range(corr_size):
            temp_list.append(corr_data.iloc[i, i + 1])

        df_temp_list = pd.DataFrame(temp_list).transpose()
        df_temp_list.columns = corr_search_header
        df_corr_search = pd.concat([df_corr_search, df_temp_list], axis=0)
        del df_temp_list
        del temp_list
        del lab_vs_mean_bof

    corr_row_sum = df_corr_search.iloc[:, 1:].sum(axis=1)
    df_corr_search['Sum'] = corr_row_sum
    df_corr_search.to_excel(writer, sheet_name='corr for each time', index=False)
    writer._save()
    current_dir = os.getcwd()

    # save the previous .xlsx file to xlsm file with macro code
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    full_name = current_dir + '/' + corr_file
    workbook = excel.Workbooks.Open(full_name)
    corr_macrofile = os.path.splitext(workbook.FullName)[0] + '.xlsm'
    workbook.SaveAs(Filename=corr_macrofile, FileFormat=52)  # 52 is for xlOpenXMLWorkbookMacroEnabled
    workbook.Close(True)

    # Connect to an existing Excel application or create a new one
    xl_app = win32.Dispatch('Excel.Application')

    # Connect to an existing workbook or create a new one
    wb = xl_app.Workbooks.Open(corr_macrofile)
    # Define the macro code

    macro_code = '''
            Sub Plot_comparation()'
                Dim sheetCnt As Long
                Dim lineCharts(0 To 100) As ChartObject
                sheetCnt = Worksheets.Count
                col_cnt = Worksheets(1).UsedRange.Columns.Count
                row_cnt = Worksheets(1).UsedRange.Rows.Count
                chart_cnt = (col_cnt - 4) / 2
                increment_cnt = col_cnt / 2
                chrt_width = Cells(1, 1).Width * 8
                chrt_height = Cells(1, 1).Height * 15

                Dim ws As Worksheet

                For i = 1 To sheetCnt - 1 Step 1
                    Set ws = ActiveWorkbook.Worksheets(i)
                    ws.Activate
                    For j = 1 To chart_cnt
                        Set lineCharts(j) = ws.ChartObjects.Add( _
                            Left:=ActiveSheet.Cells(row_cnt + 1 + 15 * (j - 1), 1).Left, _
                            Top:=ActiveSheet.Cells(row_cnt + 1 + 15 * (j - 1), 1).Top, _
                            Width:=chrt_width, _
                            Height:=chrt_height _
                            )
                        lineCharts(j).Chart.ChartType = xlLine
                        lineCharts(j).Select
                        ActiveChart.ApplyLayout (1)
                        ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = ""
                        ActiveChart.ChartType = xlLineMarkers
                        ActiveChart.Axes(xlCategory).Select
                        Selection.Delete
                        With lineCharts(j)

                            .Chart.SetSourceData Source:=Union(Cells(2, 2 + j).Resize(row_cnt - 1, 1), Cells(2, 2 + increment_cnt + j).Resize(row_cnt - 1, 1))
                            .Chart.FullSeriesCollection(1).Name = "=""LabData"""
                            .Chart.FullSeriesCollection(1).Format.Line.ForeColor.RGB = RGB(190, 75, 72)
                            .Chart.FullSeriesCollection(1).MarkerBackgroundColor = RGB(190, 75, 72)
                            .Chart.FullSeriesCollection(1).Format.Line.Weight = 1
                            .Chart.FullSeriesCollection(2).Name = "=""AnalyserData"""
                            .Chart.FullSeriesCollection(2).Format.Line.ForeColor.RGB = RGB(74, 126, 187)
                            .Chart.FullSeriesCollection(2).MarkerBackgroundColor = RGB(74, 126, 187)
                            .Chart.FullSeriesCollection(2).Format.Line.Weight = 1
                            .Chart.ChartTitle.Text = ActiveSheet.Cells(1, 2 + j).Value
                            .Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
                            '.Chart.Axes(xlCategory).AxisTitle.Delete
                        End With

                    Next j
                    Columns("A:A").EntireColumn.AutoFit
                    Columns("B:B").EntireColumn.AutoFit

                Next i
                ActiveWorkbook.Worksheets(sheetCnt).Activate
                col_cnt = Worksheets(sheetCnt).UsedRange.Columns.Count
                For i = 2 To col_cnt
                    Cells(2, i).Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.FormatConditions.AddColorScale ColorScaleType:=3
                    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
                        xlConditionValueLowestValue
                    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
                        .Color = 7039480
                        .TintAndShade = 0
                    End With
                Next i
                ActiveWindow.WindowState = xlMaximized
                ActiveWorkBook.Save
                End Sub
            '''

    # Define the module name
    module_name = 'MyModule'

    # Add the macro code to the workbook
    xl_mod = wb.VBProject.VBComponents.Add(1)  # Add module
    xl_mod.Name = module_name  # Set module name
    xl_mod.CodeModule.AddFromString(macro_code)  # Add code to module
    wb = xw.Book(corr_macrofile)
    wb.macro('Plot_comparation')()

    print('\nPlease refer to curves in each delay to run the calibration again')
    # Close the workbook and clear the memory
    del writer
    del app
    del wb
    del xl_app
    # delete the .xlsx file after it save as an xlsm file
    os.remove(current_dir + '/' + corr_file)

    input('Process finished. Press Enter to exit...')
