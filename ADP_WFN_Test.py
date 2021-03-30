import datetime
import time
import tkinter as tk
from datetime import *
from itertools import groupby
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

from PIL import Image, ImageTk
from nameparser import HumanName

import Helpers

try:
    SQLITE_FILE
except NameError:
    base_path = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\Test Databases\\'
    test_file = 'Scheduled Earnings.xlsx.db'

    SQLITE_FILE = base_path + test_file
    OUTPUT_FILE = base_path + 'Output_' + test_file + '_' + '.xls'

    cursor = Helpers.get_cursor(SQLITE_FILE)


def run():
    employees = []

    headers_dict = {'File Number': 'priorEmployeeNumber',
                    'Associate ID': 'priorEmployeeNumber',
                    'Last Name': 'lastName',
                    'First Name': 'firstName',
                    'Birth Date': 'birthDate',
                    'Payroll Company Code': 'priorCompanyCode',
                    'Tax ID (SSN)': 'ssn',
                    'Earnings Code [Additional Earnings]': 'code',
                    'Amount': 'effectiveDates amount1',
                    'Additional Earn Amount': 'effectiveDates amount1',
                    'Additonal Earn Amount': 'effectiveDates amount1',
                    # 'Additional Earnings Effective Date': '',
                    'Earnings Effective End Date': 'end_date',
                    'Additional Earnings Effective End Date': 'end_date',
                    'Status': 'status',
                    # 'Additional Earnings Code [Pay Statements]': 'code',
                    'Additional Earnings Description': 'altCodes',
                    'Earnings Code Description': 'altCodes',
                    'Position Status': 'employmentStatus',
                    'Termination Date': 'terminationDate',
                    # 'Add Additional Coverage to Amount ADP Calculates': '',
                    # 'Calculate GTL Taxable Premium Even When Not Paid': '',
                    # 'Cancel the ADP Calculation of Coverage': '',
                    # 'Employee Coverage can Exceed Annual Maximum': '',
                    # 'Employee Not Eligible for GTL (Do Not Calculate)': '',
                    # 'Override Default Benefit Factor': '',
                    # 'Override GTL Coverage Amount'
                    }

    info_headers = headers_dict.keys()

    today = datetime.now().date()
    if datetime.now().month < 3:
        effectiveDate = "01/01/" + str(datetime.now().year - 1)
    else:
        effectiveDate = "01/01/" + str(datetime.now().year)

    input_headers_dict = {}
    input_headers = cursor.execute("SELECT ElementRow, ElementText, ElementColumn FROM PageElement "
                                   "WHERE ElementRow in (SELECT ElementRow from PageElement "
                                   "WHERE ElementText = 'Payroll Company Code')").fetchall()
    for header in input_headers:
        if header[1] in info_headers:
            input_headers_dict[header[2]] = header[1]

    columns = ', '.join(map(str, input_headers_dict.keys()))

    all_data = cursor.execute(
        "SELECT ElementRow, ElementColumn, ElementText FROM PageElement WHERE ElementRow > ? "
        "and ElementColumn in (" + columns + ') '
                                             "ORDER By ElementRow", (input_headers[0][0],)).fetchall()

    for key, group in groupby(all_data, lambda row: row[0]):
        group_list = list(group)
        employee = {}
        employee['action'] = 'U'
        employee['calculate'] = 'True'
        for item in group_list:
            input_header = input_headers_dict[item[1]]
            if input_header in info_headers:
                value_header = headers_dict[input_header]
                employee[value_header] = item[2]
        emp_keys = employee.keys()

        # skip output if status of deduction is INACTIVE
        if 'status' in emp_keys:
            if employee['status'] == ('Inactive' or 'InActive'):
                continue
            else:
                del employee['status']
        if 'end_date' in emp_keys:
            check_end_date = employee['end_date']
            if check_end_date > today:
                continue
            else:
                del employee['end_date']

        # Use nameparser values if available in Full-Name, else, clean first and last name values
        if 'full_name' in emp_keys:
            full_name = employee['full_name']
            parsed_name = HumanName(full_name.encode('utf-8'))
            employee["firstName"] = parsed_name.first.title()
            employee["lastName"] = parsed_name.last.title()
            del employee['full_name']
        else:
            if 'lastName' in emp_keys:
                check_for_suffix = employee['lastName']
                if ' ' in check_for_suffix:
                    check_for_suffix = check_for_suffix.rsplit()
                    employee['lastName'] = check_for_suffix[0].title()
            if 'firstName' in emp_keys:
                check_for_middleName = employee['firstName']
                if ' ' in check_for_middleName:
                    check_for_middleName = check_for_middleName.rsplit()
                    employee['firstName'] = check_for_middleName[0].title()

        if 'effectiveDates amount1' in emp_keys:
            amount = str(employee['effectiveDates amount1'])
            if '$' in amount:
                employee['effectiveDates amount1'] = amount.replace('$', '')

        if 'effectiveDates rate1' in emp_keys:
            rate = str(employee['effectiveDates rate1'])
            if '%' in rate:
                employee['effectiveDates rate1'] = rate.replace('%', '')

        employee['effectiveDates effectiveDate1'] = effectiveDate

        # Jen Edit: changed 'code' to 'priorEmployeeNumber' since each employee should have that.
        if 'code' in employee.keys():
            if employee['code'] != '':
                employees.append(employee)

        if 'terminationDate' in employee.keys():
            if ' ' in employee['terminationDate']:
                space = employee['terminationDate'].index(' ')
                employee['terminationDate'] = str(employee['terminationDate'])[0:space]

        if 'ssn' in emp_keys:
            employee['ssn'] = Helpers.remove_non_digit(employee['ssn'])

    if employees:
        Helpers.output_excel_multitab(OUTPUT_FILE, 'Earnings', Helpers.earning_headers, employees)


run()