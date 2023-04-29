import os
import sys
from datetime import datetime
from itertools import zip_longest
from math import isclose

import xlrd
import xlwt

basedir = os.getcwd()


class ZipExhausted(Exception):
    pass


def xlsm_xlsx_compare_new_30(file_1: object, file_2: object, _precision_: object, base_dir: object) -> object:
    return compare(file_1, file_2, _precision_, base_dir)


def txt_compare(file_1, file_2):
    # noinspection PyGlobalUndefined
    global sheet_diff, row_rb2, row_rb1, sheet_found1
    head, tail = os.path.split(file_1)
    file_n_arr = tail.split(".")
    f_name_no_ext = ''
    for name_len in range(0, len(file_n_arr) - 1):
        if f_name_no_ext == '':
            f_name_no_ext = file_n_arr[name_len]
        else:
            f_name_no_ext = f_name_no_ext + "." + file_n_arr[name_len]
    differences = 0
    file_name_for_logs = get_file_name(file1)
    with open(file_1) as f1:
        with open(file_2) as f2:
            file1list = f1.read().splitlines()
            file2list = f2.read().splitlines()
            list1length = len(file1list)
            list2length = len(file2list)
            if list1length == list2length:
                with open(os.getcwd() + '\\Difference\\' + f_name_no_ext + '.txt', "w") as new_file:
                    for index in range(len(file1list)):
                        if file1list[index] != file2list[index]:
                            differences += 1
                            new_file.write("Line number { " + str(index) + " } has the below differences:\n"
                                           + file1list[index] + " --> " + file2list[index] + '\n')
                if differences == 0:
                    print("SUCCESS:\tThere are no discrepancies found in :", file_name_for_logs)
                else:
                    print("ERROR! Comparison failed due to mismatches in :", file_name_for_logs)
                new_file.close()
            else:
                print("ERROR!  Size and number of lines are different in :", file_name_for_logs)
        f1.close()
    f2.close()


def split_name():
    base_line_file_path = (os.getcwd() + '/' + 'Baseline')
    base_line_file_list_1 = os.listdir(base_line_file_path)
    base_line_file_list_1.sort()
    baseline_file_list_2 = []
    for ip in base_line_file_list_1:
        new_ip = ip[:len(base_line_file_list_1) - 6:]
        # new_ip = ip[:len(baseline_file_list_1)-6:]
        new_ip = new_ip + '_Comparison' + '.log'
        # new_ip = ip + '_Comparison' + '.log'
        baseline_file_list_2.append(new_ip)
    return baseline_file_list_2


def get_file_name(file_name):
    name = file_name.split("\\")
    return str(name[-1:])


class compare(object):

    def __init__(self, file_1, file_2, _precision_, base_dir):
        # noinspection PyGlobalUndefined
        global sheet_diff, row_rb2, row_rb1, sheet_found1
        head, tail = os.path.split(file_1)
        file_n_arr = tail.split(".")
        f_name_no_ext = ''
        for name_len in range(0, len(file_n_arr) - 1):
            if f_name_no_ext == '':
                f_name_no_ext = file_n_arr[name_len]
            else:
                f_name_no_ext = f_name_no_ext + "." + file_n_arr[name_len]

        rb1 = xlrd.open_workbook(file_1, on_demand=True)
        rb2 = xlrd.open_workbook(file_2, on_demand=True)
        diff_work_book = xlwt.Workbook()
        style = xlwt.XFStyle()
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
        style.pattern = pattern
        diff_found = 0
        file_name_for_logs = get_file_name(file1)
        for num1 in range(0, rb1.nsheets):
            row_index = 0
            sheet_index = 0
            sheet1 = rb1.sheet_by_index(num1)
            # These 2 statement will ignore the cover sheet for comparison
            if sheet1.name == 'Cover' or sheet1.name == 'Cover_Page' or sheet1.name == 'Meta Data' or sheet1.name == 'Design Document' or sheet1.name == 'Submit' or sheet1.name == 'Formula':
                continue
            sheet1name = sheet1.name
            sheet_found = 0
            sheet2 = None
            for num2 in range(0, rb2.nsheets):
                sheet2 = rb2.sheet_by_index(num2)
                # These 2 statement will ignore the cover sheet for comparison
                if sheet2.name == 'Cover' or sheet2.name == 'Cover_Page' or sheet2.name == 'Meta Data' or sheet2.name == 'Design Document' or sheet2.name == 'Submit' or sheet2.name == 'Formula':
                    continue
                sheet2name = sheet2.name
                if sheet1name == sheet2name:
                    sheet_found = 1
                    break
            if sheet_found == 1:
                sheet1rowcount = sheet1.nrows
                sheet2rowcount = sheet2.nrows
                for row_num in range(max(sheet1rowcount, sheet2rowcount)):
                    if row_index % 165536 == 0:
                        if sheet_index != 0:
                            sheet_diff = diff_work_book.add_sheet(sheet1name + str(sheet_index))
                        else:
                            sheet_diff = diff_work_book.add_sheet(sheet1name)
                        sheet_index = sheet_index + 1
                        row_index = 0
                    row = sheet_diff.row(row_index)

                    if row_num < sheet1rowcount:
                        row_rb1 = sheet1.row_values(row_num)
                        if row_num < sheet2rowcount:
                            row_rb2 = sheet2.row_values(row_num)
                            for column, (c1, c2) in enumerate(zip_longest(row_rb1, row_rb2)):
                                if str(c1) != str(c2):
                                    if isinstance(c1, float) & isinstance(c2, float):
                                        if not isclose(float(c1), float(c2), abs_tol=10 ** -_precision_):
                                            row.write(column, 'PROD:' + str(c1) + ' --> ' + 'QE:' + str(c2), style)
                                            diff_found = 1
                                    else:
                                        row.write(column, 'PROD:' + str(c1) + ' --> ' + 'QE:' + str(c2), style)
                                        diff_found = 1
                                else:
                                    row.write(column, str(c2))
                        else:
                            for col_num1, (c3, c4) in enumerate(zip_longest(row_rb1, row_rb2)):
                                row.write(col_num1, c3, style)
                                diff_found = 1
                    else:
                        row = sheet_diff.row(row_num)
                        for col_num2, (_, c6) in enumerate(zip_longest(row_rb1, row_rb2)):
                            row.write(col_num2, '', style)
                            diff_found = 1
                    row_index = row_index + 1
            else:
                sheet_diff = diff_work_book.add_sheet(sheet1name)
                for row_num1 in range(0, sheet1.nrows):
                    row_rb1 = sheet1.row_values(row_num1)
                    row_rb2 = sheet1.row_values(row_num1)
                    row = sheet_diff.row(row_num1)
                    for column, (c1, c2) in enumerate(zip_longest(row_rb1, row_rb2)):
                        row.write(column, c1, style)
                        diff_found = 1
        if diff_found == 1:
            diff_work_book.save(base_dir + '\\Difference\\' + f_name_no_ext + '.xls')
        elif diff_found == 0:
            # raise ValueError("The Data Matches")
            print('SUCCESS:\tThere are no discrepancies found in :' + file_name_for_logs)

        for num3 in range(0, rb2.nsheets):
            sheet3 = rb2.sheet_by_index(num3)
            sheet3name = sheet3.name
            sheet_found1 = 0
            for num4 in range(0, rb1.nsheets):
                sheet4 = rb1.sheet_by_index(num4)
                sheet4name = sheet4.name
                if sheet3name == sheet4name:
                    sheet_found1 = 1
                    break
            if sheet_found1 == 0:
                break

        if sheet_found1 == 0 and diff_found == 1:
            print(
                "ERROR!  Comparison failed due to mismatches in column values and Additional worksheets found in "
                "Actual file in : ", file_name_for_logs)
            # raise ValueError( "ERROR!  Comparison failed due to mismatches in column values and  Additional
            # worksheets found in Actual file")
        else:
            if sheet_found1 == 1 and diff_found == 1:
                print("ERROR!  Comparison failed due to mismatches in column values in : ", file_name_for_logs)
                # raise ValueError("ERROR!  Comparison failed due to mismatches in column values")
            else:
                if sheet_found1 == 0 and diff_found == 0:
                    print("ERROR!  Comparison failed due to Additional worksheets found in Actual file in : ",
                          file_name_for_logs)
                    # raise ValueError("ERROR!  Comparison failed due to Additional worksheets found in Actual file")


baseline_file_path = (os.getcwd() + '/' + 'Baseline')
baseline_file_list_1 = os.listdir(baseline_file_path)
baseline_file_list_1.sort()

current_file_list_path = (os.getcwd() + '/' + 'Current')
current_file_list_2 = os.listdir(current_file_list_path)
current_file_list_2.sort()

print('\tQE Comparison started\t****  {0}  ****'.format(datetime.now().strftime("%b %d %Y %H:%M:%S")))
old_stdout = sys.stdout
ip_file_names = split_name()
_precision_ = 6
log_file = open(basedir + '/Difference/' + '14A_Q_M_Report_Comparison_Log.txt', "a")
sys.stdout = log_file

with open(basedir + '/Difference/' + '14A_Q_M_Report_Comparison_Log.txt') as file:
    first_line = file.readline().rstrip()
file.close()

print('\t\tQE Comparison is performed on ****  {0}  ****'.format(datetime.now().strftime("%b %d %Y")))

for f in range(0, len(current_file_list_2)):

    file1 = os.path.join(basedir, 'Baseline', baseline_file_list_1[f])
    file2 = os.path.join(basedir, 'Current', current_file_list_2[f])

    if ".TXT" in file1 and ".TXT" in file2:
        txt_compare(file1, file2)
    else:
        xlsm_xlsx_compare_new_30(file1, file2, _precision_, basedir)

print()
sys.stdout = old_stdout
log_file.close()
print('\tQE Comparison finished\t****  {0}  ****'.format(datetime.now().strftime("%b %d %Y %H:%M:%S")))
