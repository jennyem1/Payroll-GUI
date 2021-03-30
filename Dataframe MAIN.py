from tkinter.filedialog import askopenfile

import pandas as pd

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
# from PIL import Image, ImageTk

from datetime import datetime
import os
import sys

import xlsxwriter


def main():
    # create a window for GUI app
    window = tk.Tk()
    # icon = PhotoImage('Arrow.bmp')
    # window.iconphoto(False, icon)
    window.title('Paycor Transform')
    window.geometry('350x300')

    # create frames to group widgets
    frame1 = tk.Frame(window)
    frame2 = tk.Frame(window)
    frame3 = tk.Frame(window)
    frame4 = tk.Frame(window)
    frame5 = tk.Frame(window)
    frame6 = tk.Frame(window)
    frame7 = tk.Frame(window)
    frame8 = tk.Frame(window)
    frame9 = tk.Frame(window)
    frame10 = tk.Frame(window)
    frame11 = tk.Frame(window)

    # grabbing Paycor image file and assigning it to 'logo'
    # image = Image.open('PaycorLogo.png')
    # logo = ImageTk.PhotoImage(image)

    # create logo widget
    # label = tk.Label(frame1, image=logo)
    # label.image = logo  # keep a reference!
    # label.pack()

    # create widgets for frames
    top_welcome = tk.Label(frame1, text='ADP WFN Transformations', font=('Calibri', 14, 'bold'), fg='orange')
    all_files_instruct = tk.Label(frame2, text='Transform all files to a Generic Template.', font=('Calibri', 10))
    allFiles_button = tk.Button(frame3, text='Select Folder', command=buildExcel)
    empty_space = tk.Label(frame4)
    single_files_instruct = tk.Label(frame5, text='Transform specific files to a Generic Template.', font=('Calibri', 10))
    deduction_button = tk.Button(frame6, text='Deductions', command=deductions)
    direct_deposit_button = tk.Button(frame6, text='Direct Deposit', command=directDeposit)
    earnings_button = tk.Button(frame6, text='Earnings', command=earnings)
    genEE_button = tk.Button(frame6, text='Gen EE', command=genEE)
    empty_space2 = tk.Label(frame7)
    payRates_button = tk.Button(frame8, text='Pay Rates', command=payRates)
    fedTax_button = tk.Button(frame8, text='Fed Taxes', command=fedTax)
    stateTax_button = tk.Button(frame8, text='State Taxes', command=stateTax)
    localTax_button = tk.Button(frame8, text='Local Taxes', command=localTax)
    empty_space3 = tk.Label(frame9)
    instructions_button = tk.Button(frame10, text='Notes and Instructions', command=instructWindow)
    quit_button = tk.Button(frame11, text='Quit', command=window.destroy)

    # pack everything into tkinter window
    top_welcome.pack(side='left')
    all_files_instruct.pack(side='left')
    allFiles_button.pack(side='left')
    empty_space.pack(side='left')
    single_files_instruct.pack(side='left')
    deduction_button.pack(side='left')
    direct_deposit_button.pack(side='left')
    earnings_button.pack(side='left')
    genEE_button.pack(side='left')
    empty_space2.pack(side='left')
    payRates_button.pack(side='left')
    fedTax_button.pack(side='left')
    stateTax_button.pack(side='left')
    localTax_button.pack(side='left')
    empty_space3.pack(side='left')
    instructions_button.pack(side='left')
    quit_button.pack(side='left')

    frame1.pack()
    frame2.pack()
    frame3.pack()
    frame4.pack()
    frame5.pack()
    frame6.pack()
    frame7.pack()
    frame8.pack()
    frame9.pack()
    frame10.pack()
    frame11.pack()

    tk.mainloop()


def instructWindow():
    root = tk.Tk()
    root.withdraw()

    # create a window for GUI app
    window = tk.Tk()
    # icon = PhotoImage('Arrow.bmp')
    # window.iconphoto(False, icon)
    window.title('Instructions')
    window.geometry('600x615')

    # create frames to group widgets
    frame1 = tk.Frame(window)
    frame2 = tk.Frame(window)
    frame3 = tk.Frame(window)
    frame4 = tk.Frame(window)
    frame5 = tk.Frame(window)
    frame6 = tk.Frame(window)
    frame7 = tk.Frame(window)
    frame8 = tk.Frame(window)
    frame9 = tk.Frame(window)
    frame10 = tk.Frame(window)
    frame11 = tk.Frame(window)
    frame12 = tk.Frame(window)
    frame13 = tk.Frame(window)
    frame14 = tk.Frame(window)
    frame15 = tk.Frame(window)
    frame16 = tk.Frame(window)
    frame17 = tk.Frame(window)
    frame18 = tk.Frame(window)
    frame19 = tk.Frame(window)
    frame20 = tk.Frame(window)
    frame21 = tk.Frame(window)
    frame22 = tk.Frame(window)
    frame23 = tk.Frame(window)
    frame24 = tk.Frame(window)
    frame25 = tk.Frame(window)
    frame26 = tk.Frame(window)
    frame27 = tk.Frame(window)

    # create widgets for frames
    top_welcome = tk.Label(frame1, text='ADP WFN Executable Information', font=('Calibri', 14, 'bold'), fg='orange')
    bottom_welcome1 = tk.Label(frame2,
                               text='This tool will only display the information that exists in the custom reports.',
                               font=('Calibri', 10))
    bottom_welcome2 = tk.Label(frame3,
                               text='Whatever is in the original reports will show up in the output of this tool.',
                               font=('Calibri', 10))
    details1 = tk.Label(frame4, text='This tool works from reading column headers. The column headers listed in ')
    details2 = tk.Label(frame5, text='"ADP WFN Headers.xlsx"' 'found in Teams - Paycor Data Team - General - ')
    details3 = tk.Label(frame6,
                        text='Files - _Data Desktop Tools - ADP WFN Transformation Tools. If a column is missing, ')
    details4 = tk.Label(frame7, text='there will be an error. Please check the error message for which column ')
    details5 = tk.Label(frame8, text='header that may be missing. If the error message shows something other than ')
    details6 = tk.Label(frame9, text='a missing column header, please reach out to the product owner of this tool.')
    empty_space = tk.Label(frame10)
    details7 = tk.Label(frame11, text='Every tab will contain company codes so you can filter based on company code.')
    details8 = tk.Label(frame12,
                        text='Duplicate Employees are still included on the dataframe. This issue will be addressed in the next iteration.')
    empty_space2 = tk.Label(frame13)
    instruct_title1 = tk.Label(frame14, text='Instructions for All Files Option', font=('Calibri', 12, 'bold'), fg='orange')
    instruct_2 = tk.Label(frame15, text='Place these custom excel reports to be transformed into a folder.')
    instruct_3 = tk.Label(frame16, text='Deductions | Direct Deposit | Earnings | Pay Rates | Gen EE | Fed/State/Local Tax')
    empty_space3 = tk.Label(frame17)
    instruct_title2 = tk.Label(frame18, text='Instructions for Single Files Option', font=('Calibri', 12, 'bold'), fg='orange')
    instruct_4 = tk.Label(frame19, text='Choose files from any location. Output file will be saved to the same folder.')
    empty_space4 = tk.Label(frame20)
    det_inst = tk.Label(frame21, text='NOTES', font=('Calibri', 12, 'bold'), fg='orange')
    details9 = tk.Label(frame22, text='EARNINGS: will not capture the additional codes on the payroll register.')
    details10 = tk.Label(frame23,
                         text='DIRECT DEPOSIT: accountType column will show all values other than checking or savings from client data.')
    details11 = tk.Label(frame24,
                         text='                deductionCode column may have other values than CK or SV from client data.')
    details12 = tk.Label(frame25,
                         text='                partials will be added to the DEDUCTIONS tab.')
    details13 = tk.Label(frame26,
                         text='GEN EE: middle names included in the firstName column will not be split (on account of double names).')
    details14 = tk.Label(frame27,
                         text='        tab will not identify duplicate employees. There is a column for company code on every tab.')

    top_welcome.pack(side='left')
    bottom_welcome1.pack(side='left')
    bottom_welcome2.pack(side='left')
    details1.pack(side='left')
    details2.pack(side='left')
    details3.pack(side='left')
    details4.pack(side='left')
    details5.pack(side='left')
    details6.pack(side='left')
    empty_space.pack(side='left')
    details7.pack(side='left')
    details8.pack(side='left')
    empty_space2.pack(side='left')
    instruct_title1.pack(side='left')
    instruct_2.pack(side='left')
    instruct_3.pack(side='left')
    empty_space3.pack(side='left')
    instruct_title2.pack(side='left')
    instruct_4.pack(side='left')
    empty_space4.pack(side='left')
    det_inst.pack(side='left')
    details9.pack(side='left')
    details10.pack(side='left')
    details11.pack(side='left')
    details12.pack(side='left')
    details13.pack(side='left')
    details14.pack(side='left')

    frame1.pack()
    frame2.pack()
    frame3.pack()
    frame4.pack()
    frame5.pack()
    frame6.pack()
    frame7.pack()
    frame8.pack()
    frame9.pack()
    frame10.pack()
    frame11.pack()
    frame12.pack()
    frame13.pack()
    frame14.pack()
    frame15.pack()
    frame16.pack()
    frame17.pack()
    frame18.pack()
    frame19.pack()
    frame20.pack()
    frame21.pack()
    frame22.pack()
    frame23.pack()
    frame24.pack()
    frame25.pack()
    frame26.pack()
    frame27.pack()

    tk.mainloop()


def buildExcel():
    root = tk.Tk()
    root.withdraw()

    # using tkinter filedialog to allow user to select the database
    directoryLocation = filedialog.askdirectory()
    directoryFiles = []
    for filenames in os.listdir(directoryLocation):
        if filenames.endswith('.xlsx') or filenames.endswith('.xls'):
            directoryFiles.append(os.path.join(directoryLocation, filenames))

    # breaking up the directoryFiles list into separate lists for each directory
    earnFileList = [match for match in directoryFiles if 'Earn' in match]
    genEEFileList = [match for match in directoryFiles if 'EE' in match]
    deductFileList = [match for match in directoryFiles if 'Deduct' in match]
    depositFileList = [match for match in directoryFiles if 'Deposit' in match]
    payFileList = [match for match in directoryFiles if 'Rate' in match]
    fedFileList = [match for match in directoryFiles if 'Fed' in match]
    stateFileList = [match for match in directoryFiles if 'State' in match]
    localFileList = [match for match in directoryFiles if 'Local' in match]

    # converting the lists into strings
    earnFile = ''.join(str(e) for e in earnFileList).replace('/', '\\')
    genEEFile = ''.join(str(e) for e in genEEFileList).replace('/', '\\')
    deductFile = ''.join(str(e) for e in deductFileList).replace('/', '\\')
    depositFile = ''.join(str(e) for e in depositFileList).replace('/', '\\')
    payFile = ''.join(str(e) for e in payFileList).replace('/', '\\')
    fedFile = ''.join(str(e) for e in fedFileList).replace('/', '\\')
    stateFile = ''.join(str(e) for e in stateFileList).replace('/', '\\')
    localFile = ''.join(str(e) for e in localFileList).replace('/', '\\')

    # earnFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\Scheduled Earnings.xlsx'
    # genEEFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\General EE Info.xlsx'
    # deductFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\Scheduled Deductions.xlsx'
    # depositFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\Direct Deposit.xlsx'
    # payFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\Pay Rates.xlsx'
    # fedFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\Federal Tax Info.xlsx'
    # stateFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\State Tax Info.xlsx'
    # localFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127163\\Local Tax Info.xlsx'

    try:
        ###############  DIRECT DEPOSIT  ###############
        # keeping leading zeros for file number when imported
        df_directDep = pd.read_excel(depositFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_directDep['Tax ID (SSN)'] = df_directDep['Tax ID (SSN)'].str.replace('-', '')

        df_directDep_NEW = df_directDep[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date', 'Full Net',
             'Routing Number', 'Account Number', 'Deduction Code [Direct Deposit]', 'Deduction Description',
             'Deduction Percent', 'Amount', 'Status', 'Effective Date']]

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_directDep_NEW['Effective Date'] = pd.to_datetime(df_directDep_NEW['Effective Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_directDep_NEW['Effective Date'] = pd.to_datetime(df_directDep_NEW['Effective Date']).dt.date

        # changes the date to the correct format
        df_directDep_NEW['Effective Date'] = pd.to_datetime(df_directDep_NEW['Effective Date'])
        df_directDep_NEW['Effective Date'] = df_directDep_NEW['Effective Date'].dt.strftime('%m/%d/%Y')

        df_directDep_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                         'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                         'Birth Date': 'birthDate', 'Routing Number': 'routingNumber',
                                         'Account Number': 'accountNumber',
                                         'Deduction Code [Direct Deposit]': 'deductionCode',
                                         'Deduction Description': 'accountDescription', 'Deduction Percent': 'rate',
                                         'Amount': 'amount'}, inplace=True)

        # adding new columns to match the Generic Template
        df_directDep_NEW['action'] = ''
        df_directDep_NEW['clientId'] = ''
        df_directDep_NEW['calculate'] = ''
        df_directDep_NEW['directDepositType'] = ''
        df_directDep_NEW['frequency'] = ''
        df_directDep_NEW['employmentStatus'] = ''
        df_directDep_NEW['terminationDate'] = ''
        df_directDep_NEW['accountType'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'deductionCode', 'routingNumber', 'accountNumber', 'accountType', 'accountDescription',
                        'rate', 'amount', 'calculate', 'directDepositType', 'frequency', 'Full Net', 'Status',
                        'Effective Date']

        df_directDep_NEW = df_directDep_NEW.reindex(columns=column_names)

        # copying the contents of Full Net column to directDeposityType column
        df_directDep_NEW['directDepositType'] = df_directDep_NEW['Full Net'].copy()

        # Changing the values to GT standard values
        df_directDep_NEW['directDepositType'] = df_directDep_NEW['directDepositType'].replace('Yes', 'Net')
        df_directDep_NEW['directDepositType'] = df_directDep_NEW['directDepositType'].replace('No', 'Partial')

        # adding 'Every pay period' to frequency column
        df_directDep_NEW['frequency'] = 'Every pay period'

        # Taking partials and calling them out in accountType col - per Kim's advice.
        # Since not all files will have "Checking" or "Savings" in the deduction description. Humans will have to do logic
        # to fill in partials and place them on the Deductions tab for now.
        df_directDep_NEW['accountType'] = df_directDep_NEW['directDepositType'].copy()

        # deleting Inactive direct deposits
        df_directDep_NEW = df_directDep_NEW[~df_directDep_NEW['Status'].isin(['Inactive'])]

        # deleting extra columns to match GT
        df_directDep_NEW.drop(columns=['directDepositType', 'Full Net', 'Status', 'accountDescription'])

        # filling in the NaN values with blank cells
        df_directDep_NEW.fillna('', inplace=True)

        ##########  DIRECT DEPOSIT PARTIALS TO DEDUCTIONS  ##########
        # making a dataframe for partial Direct Deposits to add to Deductions
        df_directDep_Partial = df_directDep_NEW.copy()

        partials = df_directDep_Partial[df_directDep_Partial['accountType'] == 'Net'].index
        df_directDep_Partial.drop(partials, inplace=True)

        # Creating a new dataframe with only the columns needed
        df_directDep_Partial_NEW = df_directDep_Partial[['priorCompanyCode', 'priorEmployeeNumber', 'lastName', 'firstName',
                                                         'ssn', 'birthDate', 'Effective Date', 'deductionCode', 'rate',
                                                         'amount']]

        df_directDep_Partial_NEW.rename(columns={'Effective Date': 'effectiveDates effectiveDate1', 'deductionCode': 'code',
                                                 'rate': 'effectiveDates rate', 'amount': 'effectiveDates amount'},
                                        inplace=True)

        # adding new columns to match the Generic Template
        df_directDep_Partial_NEW['calculate'] = 'True'
        df_directDep_Partial_NEW['limits maxAmount1'] = ''
        df_directDep_Partial_NEW['frequency'] = ''
        df_directDep_Partial_NEW['deductionType'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName', 'firstName',
                        'effectiveDates effectiveDate1', 'code', 'effectiveDates rate', 'effectiveDates amount',
                        'calculate', 'limits maxAmount1', 'frequency', 'deductionType']

        df_directDep_Partial_NEW = df_directDep_Partial_NEW.reindex(columns=column_names)




        ###############  DEDUCTIONS  ###############
        # keeping leading zeros for file number when imported
        df_deduct = pd.read_excel(deductFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_deduct['Tax ID (SSN)'] = df_deduct['Tax ID (SSN)'].str.replace('-', '')

        # deleting unused columns
        df_deduct_NEW = df_deduct[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Deduction Code [Deductions]', 'Percent', 'Amount', 'Effective Date', 'Accrued', 'Balance', 'Limit']]

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_deduct_NEW['Effective Date'] = pd.to_datetime(df_deduct_NEW['Effective Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_deduct_NEW['Effective Date'] = pd.to_datetime(df_deduct_NEW['Effective Date']).dt.date

        # changes the date to the correct format
        df_deduct_NEW['Effective Date'] = pd.to_datetime(df_deduct_NEW['Effective Date'])
        df_deduct_NEW['Effective Date'] = df_deduct_NEW['Effective Date'].dt.strftime('%m/%d/%Y')

        # renaming the columns to match the GT
        df_deduct_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                      'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                      'Birth Date': 'birthDate', 'Deduction Code [Deductions]': 'code',
                                      'Percent': 'effectiveDates rate1', 'Amount': 'effectiveDates amount1',
                                      'Effective Date': 'effectiveDates effectiveDate1',
                                      'Limit': 'limits maxAmount1'}, inplace=True)

        # adding new columns to match the Generic Template
        df_deduct_NEW['action'] = ''
        df_deduct_NEW['clientId'] = ''
        df_deduct_NEW['frequency'] = ''
        df_deduct_NEW['deductionType'] = ''
        df_deduct_NEW['calculate'] = 'True'

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'effectiveDates effectiveDate1', 'code', 'effectiveDates rate1',
                        'effectiveDates amount1', 'calculate', 'limits maxAmount1', 'frequency', 'deductionType']

        df_deduct_NEW = df_deduct_NEW.reindex(columns=column_names)

        # changing rate and amount columns from strings to floats
        df_deduct_NEW["effectiveDates rate1"] = pd.to_numeric(df_deduct_NEW["effectiveDates rate1"], downcast="float")
        df_deduct_NEW["effectiveDates amount1"] = pd.to_numeric(df_deduct_NEW["effectiveDates amount1"],
                                                                downcast="float")

        # Display rate and amount to two decimals
        pd.options.display.float_format = "{:,.2f}".format

        # filling in the NaN values with blank cells
        df_deduct_NEW.fillna('', inplace=True)

        # adding partials from Direct Deposit dataframe
        df_deduct_NEW = df_deduct_NEW.append(df_directDep_Partial_NEW, ignore_index=True)




        ###############  EARNINGS  ###############
        # keeping leading zeros for file number when imported
        df_earn = pd.read_excel(earnFile, converters={'File Number': lambda x: str(x)})

        # removing dashes from SSN numbers in Tax ID column
        df_earn['Tax ID (SSN)'] = df_earn['Tax ID (SSN)'].str.replace('-', '')

        # Creating a new dataframe with only the columns needed
        df_earn_NEW = df_earn[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Earnings Code [Additional Earnings]', 'Amount', 'Additional Earnings Effective End Date']]

        df_earn_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                    'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                    'Birth Date': 'birthDate', 'Earnings Code [Additional Earnings]': 'code',
                                    'Amount': 'effectiveDates amount1'}, inplace=True)

        # adding new columns to match the Generic Template
        df_earn_NEW['action'] = ''
        df_earn_NEW['clientId'] = ''
        df_earn_NEW['effectiveDates effectiveDate1'] = ''
        df_earn_NEW['effectiveDates rate1'] = ''
        df_earn_NEW['calculate'] = 'True'
        df_earn_NEW['policyAmount'] = ''
        df_earn_NEW['hours'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'effectiveDates effectiveDate1', 'code',
                        'effectiveDates amount1', 'effectiveDates rate1', 'calculate', 'policyAmount', 'hours']

        df_earn_NEW = df_earn_NEW.reindex(columns=column_names)

        # In order to stack the employees on top of one another with REG and OT, making two more dataframes with those values to append to df_earn_NEW
        df_earn_REG = df_earn_NEW.copy()

        # filling the code column with REG and OT respectively
        df_earn_REG['code'] = 'REG'

        # clearing the rates and amounts from the REG dataframe
        df_earn_REG['effectiveDates rate1'] = ''
        df_earn_REG['effectiveDates amount1'] = ''

        # adding the REG dataframe to the bottom of df_earn_NEW dataframe
        # ignore_index will result in a continuous index value and append dataframes of different sizes
        df_earn_NEW = df_earn_NEW.append(df_earn_REG, ignore_index=True)

        # In order to stack the employees on top of one another with REG and OT, making two more dataframes with those values to append to df_earn_NEW
        df_earn_OT = df_earn_NEW.copy()

        # filling the code column with REG and OT respectively
        df_earn_OT['code'] = 'OT'

        # clearing the rates and amounts from the REG dataframe
        df_earn_OT['effectiveDates rate1'] = ''
        df_earn_OT['effectiveDates amount1'] = ''

        # adding the OT dataframe to the bottom of df_earn_NEW dataframe
        # ignore_index will result in a continuous index value and append dataframes of different sizes
        df_earn_NEW = df_earn_NEW.append(df_earn_OT, ignore_index=True)

        # dropping NaN values from the code column
        df_earn_NEW.dropna(subset=['code'], inplace=True)

        # filling in the NaN values with blank cells
        df_earn_NEW.fillna('', inplace=True)

        # adding dates to the effectiveDates column
        if datetime.now().month < 3:
            df_earn_NEW['effectiveDates effectiveDate1'] = '01/01/' + str(datetime.now().year - 1)
        else:
            df_earn_NEW['effectiveDates effectiveDate1'] = '01/01/' + str(datetime.now().year)




        ###############  GENERAL EE  ###############
        # keeping leading zeros for file number when imported
        df_ee = pd.read_excel(genEEFile, converters={'File Number': lambda x: str(x)})

        # removing dashes from SSN numbers in Tax ID column
        df_ee['Tax ID (SSN)'] = df_ee['Tax ID (SSN)'].str.replace('-', '')

        # #removing decimals from the Home Department Code - USE THIS CODE AFTER REMOVING ALL NAN VALUES FROM THE DATAFRAME
        # df_ee['Home Department Code'] = df_ee['Home Department Code'].astype(int)
        df_ee.fillna('', inplace=True)

        pd.set_option('display.max_columns', None)
        # pd.set_option('display.max_rows', None)

        # Creating a new dataframe with only the columns needed
        df_ee_NEW = df_ee[
            ['Payroll Company Code', 'File Number', 'Tax ID (SSN)', 'Birth Date', 'Last Name', 'First Name',
             'Middle Name',
             'Generation Suffix Code', 'Position Status', 'Termination Date', 'Termination Reason Description',
             'Home Department Code', 'Home Department Description', 'Associate ID', 'Hire Date',
             'Worker Category Description', 'FLSA Description', 'Job Title Description', 'Reports To Associate ID',
             'Reports To Name', 'Rehire Date', 'Legal / Preferred Address: Address Line 1',
             'Legal / Preferred Address: Address Line 2', 'Legal / Preferred Address: Address Line 3',
             'Legal / Preferred Address: City', 'Legal / Preferred Address: State / Territory Code',
             'Legal / Preferred Address: Zip / Postal Code', 'Personal Contact: Home Phone',
             'Personal Contact: Personal Email', 'Personal Contact: Personal Mobile', 'Work Contact: Work Email',
             'Work Contact: Work Phone', 'Gender', 'Race Description']]

        # # splitting firstName column in case some middle names shifted
        # newName = df_ee_NEW['First Name'].str.split(' ', n=1, expand=True)
        # # making separate first/middle name column from new data frame
        # df_ee_NEW['first name'] = newName[0]
        # df_ee_NEW['middleInitial'] = newName[1]
        #
        # df_ee_NEW.fillna('', inplace=True)
        # df_ee_NEW['middleNames'] = df_ee_NEW['Middle Name'] + df_ee_NEW['middleInitial']
        #
        # df_ee_NEW.rename(columns={'middleNames': 'middleName', 'first name': 'firstName'}, inplace=True)

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date'].astype(str), errors='coerce')
        df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'].astype(str), errors='coerce')
        df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Rehire Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date']).dt.date
        df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date']).dt.date
        df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Rehire Date']).dt.date

        # changes the date to the correct format
        df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date'])
        df_ee_NEW['Termination Date'] = df_ee_NEW['Termination Date'].dt.strftime('%m/%d/%Y')

        df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'])
        df_ee_NEW['Hire Date'] = df_ee_NEW['Hire Date'].dt.strftime('%m/%d/%Y')

        df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'])
        df_ee_NEW['Rehire Date'] = df_ee_NEW['Rehire Date'].dt.strftime('%m/%d/%Y')

        # formatting the phone number columns
        df_ee_NEW['Personal Contact: Home Phone'] = df_ee_NEW['Personal Contact: Home Phone'].str.replace('[(,),-, ]',
                                                                                                          '')
        df_ee_NEW['Personal Contact: Personal Mobile'] = df_ee_NEW['Personal Contact: Personal Mobile'].str.replace(
            '[(,),-, ]', '')
        df_ee_NEW['Work Contact: Work Phone'] = df_ee_NEW['Work Contact: Work Phone'].str.replace('[(,),-, ]', '')

        df_ee_NEW['Personal Contact: Personal Mobile'] = df_ee_NEW['Personal Contact: Personal Mobile'].str.replace(
            '[-, ]',
            '')
        df_ee_NEW['Personal Contact: Home Phone'] = df_ee_NEW['Personal Contact: Home Phone'].str.replace('[-, ]', '')
        df_ee_NEW['Work Contact: Work Phone'] = df_ee_NEW['Work Contact: Work Phone'].str.replace('[-, ]', '')

        # #getting rid of NaN values in the dataframe
        # df_ee_NEW.fillna('', inplace = True)

        # concatenating the Address Line 2 & 3
        df_ee_NEW['addressLine2'] = df_ee_NEW['Legal / Preferred Address: Address Line 2'] + df_ee_NEW[
            'Legal / Preferred Address: Address Line 3']

        # changing the column names to fit the Generic Template names
        df_ee_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                  'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate', 'Last Name': 'lastName',
                                  'Generation Suffix Code': 'suffix', 'Middle Name': 'middleName',
                                  'Position Status': 'employmentStatus', 'Termination Date': 'terminationDate',
                                  'Termination Reason Description': 'terminationReason',
                                  'Home Department Code': 'departmentCode',
                                  'Home Department Description': 'departmentDescription',
                                  'Associate ID': 'employeeNumber',
                                  'Hire Date': 'hireDate', 'Worker Category Description': 'statusType',
                                  'FLSA Description': 'flsa',
                                  'Job Title Description': 'jobTitle',
                                  'Reports To Associate ID': 'managerPriorEmployeeNumber',
                                  'Reports To Name': 'reportsToName', 'Rehire Date': 'rehireDate',
                                  'Legal / Preferred Address: Address Line 1': 'addressLine1',
                                  'Legal / Preferred Address: City': 'city',
                                  'Legal / Preferred Address: State / Territory Code': 'state',
                                  'Legal / Preferred Address: Zip / Postal Code': 'zip',
                                  'Personal Contact: Home Phone': 'homePhone',
                                  'Personal Contact: Personal Email': 'homeEmail',
                                  'Personal Contact: Personal Mobile': 'mobilePhone',
                                  'Work Contact: Work Email': 'workEmail', 'Work Contact: Work Phone': 'workPhone',
                                  'Gender': 'gender',
                                  'Race Description': 'ethnicity'}, inplace=True)

        # splitting reports to manager name column
        df_ee_NEW[['managerLastName', 'managerFirstName']] = df_ee_NEW.reportsToName.str.split(',', expand=True)

        # adding new columns to match the Generic Template
        df_ee_NEW['action'] = ''
        df_ee_NEW['clientId'] = ''
        df_ee_NEW['prefix'] = ''
        df_ee_NEW['accredited'] = ''
        df_ee_NEW['payrollCode'] = ''
        df_ee_NEW['paygroupDescription'] = ''
        df_ee_NEW['employeeType'] = 'Regular'
        df_ee_NEW['maritalStatus'] = ''
        df_ee_NEW['workPhoneNumberExtension'] = ''
        df_ee_NEW['annualHours'] = ''
        df_ee_NEW['ownerOfficer'] = ''
        df_ee_NEW['baseShift'] = ''
        df_ee_NEW['managerClientId'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'middleName', 'prefix', 'suffix', 'accredited', 'employeeNumber', 'addressLine1',
                        'addressLine2', 'city', 'state', 'zip', 'departmentCode', 'departmentDescription',
                        'payrollCode',
                        'paygroupDescription', 'employmentStatus', 'terminationDate', 'terminationReason', 'reHireDate',
                        'hireDate', 'flsa', 'statusType', 'employeeType', 'maritalStatus', 'gender', 'ethnicity',
                        'jobTitle', 'workPhone', 'workPhoneNumberExtension', 'workEmail', 'mobilePhone', 'homePhone',
                        'homeEmail', 'annualHours', 'ownerOfficer', 'baseShift', 'managerPriorEmployeeNumber',
                        'managerFirstName', 'managerLastName', 'managerClientId']

        df_ee_NEW = df_ee_NEW.reindex(columns=column_names)




        ###############  PAY RATES  ###############
        # keeping leading zeros for file number when imported
        df_payRates = pd.read_excel(payFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_payRates['Tax ID (SSN)'] = df_payRates['Tax ID (SSN)'].str.replace('-', '')

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_payRates['Regular Pay Effective Date'] = pd.to_datetime(
            df_payRates['Regular Pay Effective Date'].astype(str),
            errors='coerce')
        df_payRates['Additional Rates Effective Date'] = pd.to_datetime(
            df_payRates['Additional Rates Effective Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_payRates['Regular Pay Effective Date'] = pd.to_datetime(df_payRates['Regular Pay Effective Date']).dt.date
        df_payRates['Additional Rates Effective Date'] = pd.to_datetime(
            df_payRates['Additional Rates Effective Date']).dt.date

        # changes the date to the correct format
        df_payRates['Regular Pay Effective Date'] = pd.to_datetime(df_payRates['Regular Pay Effective Date'])
        df_payRates['Regular Pay Effective Date'] = df_payRates['Regular Pay Effective Date'].dt.strftime('%m/%d/%Y')

        df_payRates['Additional Rates Effective Date'] = pd.to_datetime(df_payRates['Additional Rates Effective Date'])
        df_payRates['Additional Rates Effective Date'] = df_payRates['Additional Rates Effective Date'].dt.strftime(
            '%m/%d/%Y')

        # deleting unused columns
        df_payRates_NEW = df_payRates[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Pay Frequency', 'Regular Pay Effective Date',
             'Regular Pay Rate Description', 'Regular Pay Rate Amount', 'Additional Rates Effective Date', 'Rate 2',
             'Rate 3', 'Rate 4', 'Rate 5', 'Rate 6', 'Rate 7', 'Rate 8', 'Rate 9']]

        df_payRates_NEW.rename(
            columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                     'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                     'Birth Date': 'birthDate', 'Regular Pay Effective Date': 'effectiveDate'},
            inplace=True)

        # adding new columns to match the Generic Template
        df_payRates_NEW['action'] = ''
        df_payRates_NEW['clientId'] = ''
        df_payRates_NEW['sequence'] = ''
        df_payRates_NEW['payType'] = ''
        df_payRates_NEW['payRate'] = ''
        df_payRates_NEW['description'] = ''
        df_payRates_NEW['reason'] = ''
        df_payRates_NEW['employmentStatus'] = ''
        df_payRates_NEW['terminationDate'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'effectiveDate', 'sequence', 'payType', 'payRate',
                        'description', 'reason', 'employmentStatus', 'terminationDate', 'Pay Frequency',
                        'Regular Pay Rate Description', 'Regular Pay Rate Amount', 'Additional Rates Effective Date',
                        'Rate 2', 'Rate 3', 'Rate 4', 'Rate 5', 'Rate 6', 'Rate 7', 'Rate 8', 'Rate 9']

        df_payRates_NEW = df_payRates_NEW.reindex(columns=column_names)

        # setting sequence col to '1' for primary pay rate/amount and description col to 'Rate 1'
        df_payRates_NEW['sequence'] = '1'
        df_payRates_NEW['description'] = 'Rate 1'

        # moving data to payRate col and payType col
        df_payRates_NEW['payType'] = df_payRates_NEW['Regular Pay Rate Description'].copy()
        df_payRates_NEW['payRate'] = df_payRates_NEW['Regular Pay Rate Amount'].copy()

        # changing type of payRate col
        df_payRates_NEW['payRate'] = df_payRates_NEW['payRate'].astype(float)

        # modifying the payRate decimals
        if 'Hourly' in df_payRates_NEW['payType']:
            df_payRates_NEW['payRate'].map('{payRates:,.4f}'.format(4))
        else:
            df_payRates_NEW['payRate'].round(2)

        # deleting the unused columns
        df_payRates_NEW.drop(['employmentStatus', 'terminationDate', 'Pay Frequency', 'Regular Pay Rate Description',
                              'Regular Pay Rate Amount', 'Additional Rates Effective Date'], axis=1,
                             inplace=True)

        # drop rows with NaN values in payRate
        df_payRates_NEW.dropna(subset=['payRate'], inplace=True)

        # filling in the NaN values with blank cells
        df_payRates_NEW.fillna('', inplace=True)


        ###############  FEDERAL TAXES  ###############
        # keeping leading zeros for file number when imported
        df_taxFed = pd.read_excel(fedFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_taxFed['Tax ID (SSN)'] = df_taxFed['Tax ID (SSN)'].str.replace('-', '')

        # creating a MED dataframe - this will be the main dataframe to work from when creating all other tax code dataframes
        df_taxFed_MED = df_taxFed[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date']]

        # renaming the columns
        df_taxFed_MED.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                      'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                      'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxFed_MED['action'] = ''
        df_taxFed_MED['clientId'] = ''
        df_taxFed_MED['withholdingEffectiveStartDate'] = ''
        df_taxFed_MED['taxCode'] = 'MED'
        df_taxFed_MED['filingStatus'] = ''
        df_taxFed_MED['exemptions'] = ''
        df_taxFed_MED['adjustWithHolding'] = ''
        df_taxFed_MED['percentage'] = ''
        df_taxFed_MED['amount'] = ''
        df_taxFed_MED['reciprocity'] = ''
        df_taxFed_MED['blockDate'] = ''
        df_taxFed_MED['calculate'] = 'True'
        df_taxFed_MED['spouseWork'] = ''
        df_taxFed_MED['additionalStateExemptions'] = ''
        df_taxFed_MED['nonResidentAlienAdditionalFithwh'] = ''
        df_taxFed_MED['ncciCode'] = ''
        df_taxFed_MED['psdCode'] = ''
        df_taxFed_MED['psdRate'] = ''
        df_taxFed_MED['blockIndicator'] = ''
        df_taxFed_MED['employmentStatus'] = ''

        df_taxFed_MEDER = df_taxFed_MED.copy()
        df_taxFed_MEDER['taxCode'] = 'MEDER'

        df_taxFed_SOC = df_taxFed_MED.copy()
        df_taxFed_SOC['taxCode'] = 'SOC'

        df_taxFed_SS = df_taxFed_MED.copy()
        df_taxFed_SS['taxCode'] = 'SOCER'

        df_taxFed_FUI = df_taxFed_MED.copy()
        df_taxFed_FUI['taxCode'] = 'FUI'

        # creating a new master dataframe
        df_taxFed_NEW = df_taxFed_MED.copy()

        # adding the new dataframes to the master dataframe
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_MEDER, ignore_index=True)
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_SOC, ignore_index=True)
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_SS, ignore_index=True)
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_FUI, ignore_index=True)

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus',
                        'exemptions', 'adjustWithHolding', 'percentage', 'amount', 'reciprocity', 'blockDate',
                        'calculate',
                        'spouseWork', 'additionalStateExemptions', 'nonResidentAlienAdditionalFitwh', 'ncciCode',
                        'psdCode',
                        'psdRate', 'blockIndicator', 'employmentStatus', 'terminationDate']

        df_taxFed_NEW = df_taxFed_NEW.reindex(columns=column_names)

        # # changing the date to string and forcing anything else in the column to do so as well with coerce
        # df_taxFed_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_NEW['withholdingEffectiveStartDate'].astype(str), errors='coerce')
        #
        # # changing the date to a date format
        # df_taxFed_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_NEW['withholdingEffectiveStartDate']).dt.date
        #
        # # changes the date to the correct format
        # df_taxFed_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_NEW['withholdingEffectiveStartDate'])
        # df_taxFed_NEW['withholdingEffectiveStartDate'] = df_taxFed_NEW['withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')

        # pulling in original dataframe to calculate FIT tax - additional columns needed
        df_taxFed_FIT = df_taxFed[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Do Not Calculate Federal Income Tax', 'Federal/W4 Effective Date',
             'Federal/W4 Marital Status Description', 'Federal/W4 Exemptions', 'Federal Additional Tax Amount',
             'Federal Additional Tax Amount Percentage']]

        # adding new columns to match the Generic Template
        df_taxFed_FIT['action'] = ''
        df_taxFed_FIT['clientId'] = ''
        df_taxFed_FIT['withholdingEffectiveStartDate'] = ''
        df_taxFed_FIT['taxCode'] = 'FIT'
        df_taxFed_FIT['filingStatus'] = ''
        df_taxFed_FIT['exemptions'] = ''
        df_taxFed_FIT['adjustWithHolding'] = ''
        df_taxFed_FIT['percentage'] = ''
        df_taxFed_FIT['amount'] = ''
        df_taxFed_FIT['reciprocity'] = ''
        df_taxFed_FIT['blockDate'] = ''
        df_taxFed_FIT['calculate'] = 'True'
        df_taxFed_FIT['spouseWork'] = ''
        df_taxFed_FIT['additionalStateExemptions'] = ''
        df_taxFed_FIT['nonResidentAlienAdditionalFithwh'] = ''
        df_taxFed_FIT['ncciCode'] = ''
        df_taxFed_FIT['psdCode'] = ''
        df_taxFed_FIT['psdRate'] = ''
        df_taxFed_FIT['blockIndicator'] = ''
        df_taxFed_FIT['employmentStatus'] = ''

        # adding data to GT columns
        df_taxFed_FIT['filingStatus'] = df_taxFed_FIT['Federal/W4 Marital Status Description']
        df_taxFed_FIT['exemptions'] = df_taxFed_FIT['Federal/W4 Exemptions']
        df_taxFed_FIT['amount'] = df_taxFed_FIT['Federal Additional Tax Amount']
        df_taxFed_FIT['withholdingEffectiveStartDate'] = df_taxFed_FIT['Federal/W4 Effective Date']

        # removing NaN values from the amount and percentage columns
        df_taxFed_FIT['Federal Additional Tax Amount'].fillna(0, inplace=True)
        df_taxFed_FIT['Federal Additional Tax Amount Percentage'].fillna(0, inplace=True)

        # changing data type from str to in for Amount and Percentage columns
        df_taxFed_FIT['Federal Additional Tax Amount'] = df_taxFed_FIT['Federal Additional Tax Amount'].astype(float)
        df_taxFed_FIT['Federal Additional Tax Amount Percentage'] = df_taxFed_FIT[
            'Federal Additional Tax Amount Percentage'].astype(float)

        # conditional statements for adjustWithHolding
        df_taxFed_FIT.loc[
            df_taxFed_FIT['Federal Additional Tax Amount'] > 0, 'adjustWithHolding'] = 'Add to withholding'
        df_taxFed_FIT.loc[
            df_taxFed_FIT['Federal Additional Tax Amount Percentage'] > 0, 'adjustWithHolding'] = 'Add to withholding'

        # renaming the columns
        df_taxFed_FIT.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                      'Last Name': 'lastName', 'First Name': 'firstName',
                                      'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate'}, inplace=True)

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus',
                        'exemptions', 'adjustWithHolding', 'percentage', 'amount', 'reciprocity', 'blockDate',
                        'calculate',
                        'spouseWork', 'additionalStateExemptions', 'nonResidentAlienAdditionalFitwh',
                        'ncciCode', 'psdCode', 'psdRate', 'blockIndicator', 'employmentStatus', 'terminationDate']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxFed_FIT['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxFed_FIT['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_FIT['withholdingEffectiveStartDate'])
        df_taxFed_FIT['withholdingEffectiveStartDate'] = df_taxFed_FIT['withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxFed_FIT = df_taxFed_FIT.reindex(columns=column_names)
        # adding the new dataframes to the master dataframe
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_FIT, ignore_index=True)

        # filling in the NaN values with blank cells
        df_taxFed_NEW.fillna('', inplace=True)


        ###############  STATE TAXES  ###############
        # keeping leading zeros for file number when imported
        df_taxState = pd.read_excel(stateFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_taxState['Tax ID (SSN)'] = df_taxState['Tax ID (SSN)'].str.replace('-', '')
        # creating a SUI dataframe
        df_taxState_State = df_taxState[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'State Additional Tax Amount', 'State Additional Tax Amount Percentage', 'State Effective Date',
             'State Exemptions/Allowances', 'State Marital Status Description', 'State Tax Code',
             'State Tax Description']]

        # renaming the columns
        df_taxState_State.rename(
            columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                     'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                     'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxState_State['action'] = ''
        df_taxState_State['clientId'] = ''
        df_taxState_State['withholdingEffectiveStartDate'] = ''
        df_taxState_State['taxCode'] = ''
        df_taxState_State['filingStatus'] = ''
        df_taxState_State['exemptions'] = ''
        df_taxState_State['adjustWithHolding'] = ''
        df_taxState_State['percentage'] = ''
        df_taxState_State['amount'] = ''
        df_taxState_State['reciprocity'] = ''
        df_taxState_State['blockDate'] = ''
        df_taxState_State['calculate'] = 'True'
        df_taxState_State['spouseWork'] = ''
        df_taxState_State['additionalStateExemptions'] = ''
        df_taxState_State['nonResidentAlienAdditionalFithwh'] = ''
        df_taxState_State['ncciCode'] = ''
        df_taxState_State['psdCode'] = ''
        df_taxState_State['psdRate'] = ''
        df_taxState_State['blockIndicator'] = ''

        # adding data to GT columns
        df_taxState_State['taxCode'] = df_taxState_State['State Tax Code']
        df_taxState_State['calculate'] = 'True'

        df_taxState_State['amount'] = df_taxState_State['State Additional Tax Amount']
        df_taxState_State['percentage'] = df_taxState_State['State Additional Tax Amount Percentage']

        df_taxState_State['withholdingEffectiveStartDate'] = df_taxState_State['State Effective Date']

        # 201223 Adding the `filingStatus` and `exemptions` fields to state taxes.
        # Also added on line 2005.
        df_taxState_State['filingStatus'] = df_taxState_State['State Marital Status Description']
        df_taxState_State['exemptions'] = df_taxState_State['State Exemptions/Allowances']

        # removing NaN values from the amount and percentage columns
        df_taxState_State['State Additional Tax Amount'].fillna(0, inplace=True)
        df_taxState_State['State Additional Tax Amount Percentage'].fillna(0, inplace=True)

        # changing data type from str to in for Amount and Percentage columns
        df_taxState_State['State Additional Tax Amount'] = df_taxState_State['State Additional Tax Amount'].astype(
            float)
        df_taxState_State['State Additional Tax Amount Percentage'] = df_taxState_State[
            'State Additional Tax Amount Percentage'].astype(float)

        # conditional statements for adjustWithHolding
        df_taxState_State.loc[
            df_taxState_State['State Additional Tax Amount'] > 0, 'adjustWithHolding'] = 'Add to withholding'
        df_taxState_State.loc[
            df_taxState_State['State Additional Tax Amount Percentage'] > 0, 'adjustWithHolding'] = 'Add to withholding'

        # dropping columns from the dataframe
        df_taxState_State = df_taxState_State.drop(
            columns=['State Additional Tax Amount', 'State Additional Tax Amount Percentage', 'State Effective Date',
                     'State Exemptions/Allowances', 'State Marital Status Description', 'State Tax Code',
                     'State Tax Description'])

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_State['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_State['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_State['withholdingEffectiveStartDate'])
        df_taxState_State['withholdingEffectiveStartDate'] = df_taxState_State[
            'withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxState_State = df_taxState_State.reindex(columns=column_names)

        # filling in the NaN values with blank cells
        df_taxState_State.fillna('', inplace=True)

        # creating a new state UNE dataframe
        df_taxState_UNE = df_taxState[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'SUI/SDI Effective Date', 'SUI/SDI Tax Code', 'SUI/SDI Tax Code Description']]

        # renaming the columns
        df_taxState_UNE.rename(
            columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                     'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                     'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxState_UNE['action'] = ''
        df_taxState_UNE['clientId'] = ''
        df_taxState_State['withholdingEffectiveStartDate'] = ''
        df_taxState_UNE['taxCode'] = ''
        df_taxState_UNE['filingStatus'] = ''
        df_taxState_UNE['exemptions'] = ''
        df_taxState_UNE['adjustWithHolding'] = ''
        df_taxState_UNE['percentage'] = ''
        df_taxState_UNE['amount'] = ''
        df_taxState_UNE['reciprocity'] = ''
        df_taxState_UNE['blockDate'] = ''
        df_taxState_UNE['calculate'] = 'True'
        df_taxState_UNE['spouseWork'] = ''
        df_taxState_UNE['additionalStateExemptions'] = ''
        df_taxState_UNE['nonResidentAlienAdditionalFithwh'] = ''
        df_taxState_UNE['ncciCode'] = ''
        df_taxState_UNE['psdCode'] = ''
        df_taxState_UNE['psdRate'] = ''
        df_taxState_UNE['blockIndicator'] = ''

        # adding data to GT columns
        df_taxState_UNE['taxCode'] = df_taxState_UNE['SUI/SDI Tax Code']

        df_taxState_UNE['withholdingEffectiveStartDate'] = df_taxState_UNE['SUI/SDI Effective Date']

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_UNE['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_UNE['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_UNE['withholdingEffectiveStartDate'])
        df_taxState_UNE['withholdingEffectiveStartDate'] = df_taxState_UNE['withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxState_UNE = df_taxState_UNE.reindex(columns=column_names)

        # creating a new master dataframe
        df_taxState_NEW = df_taxState_State.copy()

        # adding the new dataframes to the master dataframe
        df_taxState_NEW = df_taxState_NEW.append(df_taxState_UNE, ignore_index=True)

        # filling in the NaN values with blank cells
        df_taxState_NEW.fillna('', inplace=True)


        ###############  LOCAL TAXES  ###############
        # keeping leading zeros for file number when imported
        df_taxLocal = pd.read_excel(localFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_taxLocal['Tax ID (SSN)'] = df_taxLocal['Tax ID (SSN)'].str.replace('-', '')

        # getting rid of the unneeded column headers
        df_taxLocal_NEW = df_taxLocal[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'School Tax Code',
             'School Tax Description', 'Local Tax Code', 'Local Tax Description', 'Local Effective Date']]

        # renaming the columns
        df_taxLocal_NEW.rename(
            columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                     'Last Name': 'lastName',
                     'First Name': 'firstName', 'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxLocal_NEW['action'] = ''
        df_taxLocal_NEW['clientId'] = ''
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = ''
        df_taxLocal_NEW['taxCode'] = ''
        df_taxLocal_NEW['filingStatus'] = ''
        df_taxLocal_NEW['exemptions'] = ''
        df_taxLocal_NEW['adjustWithHolding'] = ''
        df_taxLocal_NEW['percentage'] = ''
        df_taxLocal_NEW['amount'] = ''
        df_taxLocal_NEW['reciprocity'] = ''
        df_taxLocal_NEW['blockDate'] = ''
        df_taxLocal_NEW['calculate'] = 'True'
        df_taxLocal_NEW['spouseWork'] = ''
        df_taxLocal_NEW['additionalStateExemptions'] = ''
        df_taxLocal_NEW['nonResidentAlienAdditionalFithwh'] = ''
        df_taxLocal_NEW['ncciCode'] = ''
        df_taxLocal_NEW['psdCode'] = ''
        df_taxLocal_NEW['psdRate'] = ''
        df_taxLocal_NEW['blockIndicator'] = ''

        # adding data to GT columns
        df_taxLocal_NEW['taxCode'] = df_taxLocal_NEW['Local Tax Code'] + ' ' + df_taxLocal_NEW['Local Tax Description']
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = df_taxLocal_NEW['Local Effective Date']

        # dropping columns from the dataframe
        df_taxLocal_NEW = df_taxLocal_NEW.drop(
            columns=['School Tax Code', 'School Tax Description', 'Local Tax Code', 'Local Tax Description',
                     'Local Effective Date'])

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_NEW['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_NEW['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_NEW['withholdingEffectiveStartDate'])
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = df_taxLocal_NEW['withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxLocal_NEW = df_taxLocal_NEW.reindex(columns=column_names)

        # creating a dataframe for school tax if present
        # getting rid of the unneeded column headers
        df_taxLocal_School = df_taxLocal[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'School Tax Code',
             'School Tax Description', 'School District Effective Date']]

        # renaming the columns
        df_taxLocal_School.rename(
            columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                     'Last Name': 'lastName',
                     'First Name': 'firstName', 'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxLocal_School['action'] = ''
        df_taxLocal_School['clientId'] = ''
        df_taxLocal_School['withholdingEffectiveStartDate'] = ''
        df_taxLocal_School['taxCode'] = ''
        df_taxLocal_School['filingStatus'] = ''
        df_taxLocal_School['exemptions'] = ''
        df_taxLocal_School['adjustWithHolding'] = ''
        df_taxLocal_School['percentage'] = ''
        df_taxLocal_School['amount'] = ''
        df_taxLocal_School['reciprocity'] = ''
        df_taxLocal_School['blockDate'] = ''
        df_taxLocal_School['calculate'] = 'True'
        df_taxLocal_School['spouseWork'] = ''
        df_taxLocal_School['additionalStateExemptions'] = ''
        df_taxLocal_School['nonResidentAlienAdditionalFithwh'] = ''
        df_taxLocal_School['ncciCode'] = ''
        df_taxLocal_School['psdCode'] = ''
        df_taxLocal_School['psdRate'] = ''
        df_taxLocal_School['blockIndicator'] = ''

        # adding data to GT columns
        df_taxLocal_School['taxCode'] = df_taxLocal_School['School Tax Code'] + ' ' + df_taxLocal_School[
            'School Tax Description']
        df_taxLocal_School['withholdingEffectiveStartDate'] = df_taxLocal_School['School District Effective Date']

        # dropping columns from the dataframe
        df_taxLocal_School = df_taxLocal_School.drop(
            columns=['School Tax Code', 'School Tax Description', 'School District Effective Date'])

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_School['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_School['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_School['withholdingEffectiveStartDate'])
        df_taxLocal_School['withholdingEffectiveStartDate'] = df_taxLocal_School[
            'withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')

        df_taxLocal_School = df_taxLocal_School.reindex(columns=column_names)

        # adding the new dataframes to the master dataframe
        df_taxLocal_NEW = df_taxLocal_NEW.append(df_taxLocal_School, ignore_index=True)

        # deleting rows with NaN values in taxCode column
        df_taxLocal_NEW = df_taxLocal_NEW.dropna(subset=['taxCode'])

        # filling in the NaN values with blank cells
        df_taxLocal_NEW.fillna('', inplace=True)


        #############  COMBINING TAXES  #############
        # adding the new dataframes to the master dataframe
        df_taxes = df_taxFed_NEW.copy()
        df_taxes = df_taxes.append(df_taxState_NEW, ignore_index=True)
        df_taxFed_NEW = df_taxes.append(df_taxLocal_NEW, ignore_index=True)

        df_taxes.loc[
            df_taxes['filingStatus'] == 'Married, but withhold at higher single rate', 'filingStatus'] = 'Single'
        df_taxes.loc[
            df_taxes['filingStatus'] == 'Married filing jointly (or Qualifying widow(er))', 'filingStatus'] = 'Married'
        df_taxes.loc[df_taxes['filingStatus'] == 'Single or Married Filing Separately', 'filingStatus'] = 'Single'


        #############  OUTPUT EXCEL GT  #############
        # create a pandas excel writer using xlsxwriter as the engine.

        # using tkinter filedialog to allow user to select the database
        saveLocation = directoryLocation
        saveDir = ''.join(str(e) for e in saveLocation).replace('/', '\\')
        writer = pd.ExcelWriter(saveDir+'\\Generic Template OUTPUT.xlsx', engine='xlsxwriter',
                                date_format='dd/mm/yy')  # pylint: disable=abstract-class-instantiated

        # combining all of the excel files into one GT (deleting the individual files from directory) using xlsxwriter
        df_ee_NEW.to_excel(writer, sheet_name='Employee', index=False)
        df_taxes.to_excel(writer, sheet_name='Employee Taxes', index=False)
        df_earn_NEW.to_excel(writer, sheet_name='Employee Earning', index=False)
        df_deduct_NEW.to_excel(writer, sheet_name='Employee Deductions', index=False)
        df_payRates_NEW.to_excel(writer, sheet_name='Employee Pay Rates', index=False)
        df_directDep_NEW.to_excel(writer, sheet_name='Employee Direct Deposit', index=False)

        writer.save()

        messagebox.showinfo('File', 'File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def deductions():
    root = tk.Tk()
    root.withdraw()

    deductFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_deduct = pd.read_excel(deductFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_deduct['Tax ID (SSN)'] = df_deduct['Tax ID (SSN)'].str.replace('-', '')

        # deleting unused columns
        df_deduct_NEW = df_deduct[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Deduction Code [Deductions]', 'Percent', 'Amount', 'Effective Date', 'Accrued', 'Balance', 'Limit']]

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_deduct_NEW['Effective Date'] = pd.to_datetime(df_deduct_NEW['Effective Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_deduct_NEW['Effective Date'] = pd.to_datetime(df_deduct_NEW['Effective Date']).dt.date

        # changes the date to the correct format
        df_deduct_NEW['Effective Date'] = pd.to_datetime(df_deduct_NEW['Effective Date'])
        df_deduct_NEW['Effective Date'] = df_deduct_NEW['Effective Date'].dt.strftime('%m/%d/%Y')

        # renaming the columns to match the GT
        df_deduct_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                      'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                      'Birth Date': 'birthDate', 'Deduction Code [Deductions]': 'code',
                                      'Percent': 'effectiveDates rate1', 'Amount': 'effectiveDates amount1',
                                      'Effective Date': 'effectiveDates effectiveDate1',
                                      'Limit': 'limits maxAmount1'}, inplace=True)

        # adding new columns to match the Generic Template
        df_deduct_NEW['action'] = ''
        df_deduct_NEW['clientId'] = ''
        df_deduct_NEW['frequency'] = ''
        df_deduct_NEW['deductionType'] = ''
        df_deduct_NEW['calculate'] = 'True'

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName',
                        'effectiveDates effectiveDate1', 'code', 'effectiveDates rate1', 'effectiveDates amount1',
                        'calculate', 'limits maxAmount1', 'frequency', 'deductionType']

        df_deduct_NEW = df_deduct_NEW.reindex(columns=column_names)

        # changing rate and amount columns from strings to floats
        df_deduct_NEW["effectiveDates rate1"] = pd.to_numeric(df_deduct_NEW["effectiveDates rate1"], downcast="float")
        df_deduct_NEW["effectiveDates amount1"] = pd.to_numeric(df_deduct_NEW["effectiveDates amount1"], downcast="float")

        # Display rate and amount to two decimals
        pd.options.display.float_format = "{:,.2f}".format

        # filling in the NaN values with blank cells
        df_deduct_NEW.fillna('', inplace=True)

        # saving deductions file
        deductFilePath = os.path.dirname(deductFile)
        # deductFilePath = tk.filedialog.asksaveasfilename()
        df_deduct_NEW.to_excel(deductFilePath+'\\Deductions OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'Deductions File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def directDeposit():
    root = tk.Tk()
    root.withdraw()

    depositFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_directDep = pd.read_excel(depositFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_directDep['Tax ID (SSN)'] = df_directDep['Tax ID (SSN)'].str.replace('-', '')

        df_directDep_NEW = df_directDep[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date', 'Full Net',
             'Routing Number', 'Account Number', 'Deduction Code [Direct Deposit]', 'Deduction Description',
             'Deduction Percent', 'Amount', 'Status', 'Effective Date']]

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_directDep_NEW['Effective Date'] = pd.to_datetime(df_directDep_NEW['Effective Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_directDep_NEW['Effective Date'] = pd.to_datetime(df_directDep_NEW['Effective Date']).dt.date

        # changes the date to the correct format
        df_directDep_NEW['Effective Date'] = pd.to_datetime(df_directDep_NEW['Effective Date'])
        df_directDep_NEW['Effective Date'] = df_directDep_NEW['Effective Date'].dt.strftime('%m/%d/%Y')

        df_directDep_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                         'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                         'Birth Date': 'birthDate', 'Routing Number': 'routingNumber',
                                         'Account Number': 'accountNumber',
                                         'Deduction Code [Direct Deposit]': 'deductionCode',
                                         'Deduction Description': 'accountDescription', 'Deduction Percent': 'rate',
                                         'Amount': 'amount'}, inplace=True)

        # adding new columns to match the Generic Template
        df_directDep_NEW['action'] = ''
        df_directDep_NEW['clientId'] = ''
        df_directDep_NEW['calculate'] = ''
        df_directDep_NEW['directDepositType'] = ''
        df_directDep_NEW['frequency'] = ''
        df_directDep_NEW['employmentStatus'] = ''
        df_directDep_NEW['terminationDate'] = ''
        df_directDep_NEW['accountType'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'deductionCode', 'routingNumber', 'accountNumber', 'accountType', 'accountDescription',
                        'rate', 'amount', 'calculate', 'directDepositType', 'frequency', 'Full Net', 'Status',
                        'Effective Date']

        df_directDep_NEW = df_directDep_NEW.reindex(columns=column_names)

        # copying the contents of Full Net column to directDeposityType column
        df_directDep_NEW['directDepositType'] = df_directDep_NEW['Full Net'].copy()

        # Changing the values to GT standard values
        df_directDep_NEW['directDepositType'] = df_directDep_NEW['directDepositType'].replace('Yes', 'Net')
        df_directDep_NEW['directDepositType'] = df_directDep_NEW['directDepositType'].replace('No', 'Partial')

        # adding 'Every pay period' to frequency column
        df_directDep_NEW['frequency'] = 'Every pay period'

        # Taking partials and calling them out in accountType col - per Kim's advice.
        # Since not all files will have "Checking" or "Savings" in the deduction description. Humans will have to do logic
        # to fill in partials and place them on the Deductions tab for now.
        df_directDep_NEW['accountType'] = df_directDep_NEW['directDepositType'].copy()

        # deleting Inactive direct deposits
        df_directDep_NEW = df_directDep_NEW[~df_directDep_NEW['Status'].isin(['Inactive'])]

        # deleting extra columns to match GT
        df_directDep_NEW.drop(columns=['directDepositType', 'Full Net', 'Status', 'accountDescription'])

        # filling in the NaN values with blank cells
        df_directDep_NEW.fillna('', inplace=True)

        # saving deductions file
        dirDepFilePath = os.path.dirname(depositFile)
        df_directDep_NEW.to_excel(dirDepFilePath+'\\Direct Deposit OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'Direct Deposit File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def earnings():
    root = tk.Tk()
    root.withdraw()

    earnFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_earn = pd.read_excel(earnFile, converters={'File Number': lambda x: str(x)})

        # removing dashes from SSN numbers in Tax ID column
        df_earn['Tax ID (SSN)'] = df_earn['Tax ID (SSN)'].str.replace('-', '')

        # Creating a new dataframe with only the columns needed
        df_earn_NEW = df_earn[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Earnings Code [Additional Earnings]', 'Amount', 'Additional Earnings Effective End Date']]

        df_earn_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                    'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                    'Birth Date': 'birthDate', 'Earnings Code [Additional Earnings]': 'code',
                                    'Amount': 'effectiveDates amount1'}, inplace=True)

        # adding new columns to match the Generic Template
        df_earn_NEW['action'] = ''
        df_earn_NEW['clientId'] = ''
        df_earn_NEW['effectiveDates effectiveDate1'] = ''
        df_earn_NEW['effectiveDates rate1'] = ''
        df_earn_NEW['calculate'] = 'True'
        df_earn_NEW['policyAmount'] = ''
        df_earn_NEW['hours'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'effectiveDates effectiveDate1', 'code', 'effectiveDates amount1',
                        'effectiveDates rate1',
                        'calculate', 'policyAmount', 'hours']

        df_earn_NEW = df_earn_NEW.reindex(columns=column_names)

        # In order to stack the employees on top of one another with REG and OT, making two more dataframes with those values to append to df_earn_NEW
        df_earn_REG = df_earn_NEW.copy()

        # filling the code column with REG and OT respectively
        df_earn_REG['code'] = 'REG'

        # clearing the rates and amounts from the REG dataframe
        df_earn_REG['effectiveDates rate1'] = ''
        df_earn_REG['effectiveDates amount1'] = ''

        # adding the REG dataframe to the bottom of df_earn_NEW dataframe
        # ignore_index will result in a continuous index value and append dataframes of different sizes
        df_earn_NEW = df_earn_NEW.append(df_earn_REG, ignore_index=True)

        # In order to stack the employees on top of one another with REG and OT, making two more dataframes with those values to append to df_earn_NEW
        df_earn_OT = df_earn_NEW.copy()

        # filling the code column with REG and OT respectively
        df_earn_OT['code'] = 'OT'

        # clearing the rates and amounts from the REG dataframe
        df_earn_OT['effectiveDates rate1'] = ''
        df_earn_OT['effectiveDates amount1'] = ''

        # adding the OT dataframe to the bottom of df_earn_NEW dataframe
        # ignore_index will result in a continuous index value and append dataframes of different sizes
        df_earn_NEW = df_earn_NEW.append(df_earn_OT, ignore_index=True)

        # dropping NaN values from the code column
        df_earn_NEW.dropna(subset=['code'], inplace=True)

        # filling in the NaN values with blank cells
        df_earn_NEW.fillna('', inplace=True)

        # adding dates to the effectiveDates column
        if datetime.now().month < 3:
            df_earn_NEW['effectiveDates effectiveDate1'] = '01/01/' + str(datetime.now().year - 1)
        else:
            df_earn_NEW['effectiveDates effectiveDate1'] = '01/01/' + str(datetime.now().year)

        # saving earnings file
        earnFilePath = os.path.dirname(earnFile)
        df_earn_NEW.to_excel(earnFilePath+'\\Earnings OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'Earnings File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def genEE():
    root = tk.Tk()
    root.withdraw()

    genEEFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_ee = pd.read_excel(genEEFile, converters={'File Number': lambda x: str(x)})

        # removing dashes from SSN numbers in Tax ID column
        df_ee['Tax ID (SSN)'] = df_ee['Tax ID (SSN)'].str.replace('-', '')

        # #removing decimals from the Home Department Code - USE THIS CODE AFTER REMOVING ALL NAN VALUES FROM THE DATAFRAME
        # df_ee['Home Department Code'] = df_ee['Home Department Code'].astype(int)
        df_ee.fillna('', inplace=True)

        pd.set_option('display.max_columns', None)
        # pd.set_option('display.max_rows', None)

        # Creating a new dataframe with only the columns needed
        df_ee_NEW = df_ee[
            ['Payroll Company Code', 'File Number', 'Tax ID (SSN)', 'Birth Date', 'Last Name', 'First Name', 'Middle Name',
             'Generation Suffix Code', 'Position Status', 'Termination Date', 'Termination Reason Description',
             'Home Department Code', 'Home Department Description', 'Associate ID', 'Hire Date',
             'Worker Category Description', 'FLSA Description', 'Job Title Description', 'Reports To Associate ID',
             'Reports To Name', 'Rehire Date', 'Legal / Preferred Address: Address Line 1',
             'Legal / Preferred Address: Address Line 2', 'Legal / Preferred Address: Address Line 3',
             'Legal / Preferred Address: City', 'Legal / Preferred Address: State / Territory Code',
             'Legal / Preferred Address: Zip / Postal Code', 'Personal Contact: Home Phone',
             'Personal Contact: Personal Email', 'Personal Contact: Personal Mobile', 'Work Contact: Work Email',
             'Work Contact: Work Phone', 'Gender', 'Race Description']]

        # splitting firstName column in case some middle names shifted
        newName = df_ee_NEW['First Name'].str.split(' ', n=1, expand=True)
        # making separate first/middle name column from new data frame
        df_ee_NEW['first name'] = newName[0]
        df_ee_NEW['middleInitial'] = newName[1]

        df_ee_NEW.fillna('', inplace=True)
        df_ee_NEW['middleNames'] = df_ee_NEW['Middle Name'] + df_ee_NEW['middleInitial']

        df_ee_NEW.rename(columns={'middleNames': 'middleName', 'first name': 'firstName'}, inplace=True)

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date'].astype(str), errors='coerce')
        df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'].astype(str), errors='coerce')
        df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Rehire Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date']).dt.date
        df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date']).dt.date
        df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Rehire Date']).dt.date

        # changes the date to the correct format
        df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date'])
        df_ee_NEW['Termination Date'] = df_ee_NEW['Termination Date'].dt.strftime('%m/%d/%Y')

        df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'])
        df_ee_NEW['Hire Date'] = df_ee_NEW['Hire Date'].dt.strftime('%m/%d/%Y')

        df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'])
        df_ee_NEW['Rehire Date'] = df_ee_NEW['Rehire Date'].dt.strftime('%m/%d/%Y')

        # formatting the phone number columns
        df_ee_NEW['Personal Contact: Home Phone'] = df_ee_NEW['Personal Contact: Home Phone'].str.replace('[(,),-, ]', '')
        df_ee_NEW['Personal Contact: Personal Mobile'] = df_ee_NEW['Personal Contact: Personal Mobile'].str.replace(
            '[(,),-, ]', '')
        df_ee_NEW['Work Contact: Work Phone'] = df_ee_NEW['Work Contact: Work Phone'].str.replace('[(,),-, ]', '')

        df_ee_NEW['Personal Contact: Personal Mobile'] = df_ee_NEW['Personal Contact: Personal Mobile'].str.replace('[-, ]',
                                                                                                                    '')
        df_ee_NEW['Personal Contact: Home Phone'] = df_ee_NEW['Personal Contact: Home Phone'].str.replace('[-, ]', '')
        df_ee_NEW['Work Contact: Work Phone'] = df_ee_NEW['Work Contact: Work Phone'].str.replace('[-, ]', '')

        # #getting rid of NaN values in the dataframe
        # df_ee_NEW.fillna('', inplace = True)

        # concatenating the Address Line 2 & 3
        df_ee_NEW['addressLine2'] = df_ee_NEW['Legal / Preferred Address: Address Line 2'] + df_ee_NEW[
            'Legal / Preferred Address: Address Line 3']

        # changing the column names to fit the Generic Template names
        df_ee_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                  'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate', 'Last Name': 'lastName',
                                  'Generation Suffix Code': 'suffix',
                                  'Position Status': 'employmentStatus', 'Termination Date': 'terminationDate',
                                  'Termination Reason Description': 'terminationReason',
                                  'Home Department Code': 'departmentCode',
                                  'Home Department Description': 'departmentDescription', 'Associate ID': 'employeeNumber',
                                  'Hire Date': 'hireDate', 'Worker Category Description': 'statusType',
                                  'FLSA Description': 'flsa',
                                  'Job Title Description': 'jobTitle',
                                  'Reports To Associate ID': 'managerPriorEmployeeNumber',
                                  'Reports To Name': 'reportsToName', 'Rehire Date': 'rehireDate',
                                  'Legal / Preferred Address: Address Line 1': 'addressLine1',
                                  'Legal / Preferred Address: City': 'city',
                                  'Legal / Preferred Address: State / Territory Code': 'state',
                                  'Legal / Preferred Address: Zip / Postal Code': 'zip',
                                  'Personal Contact: Home Phone': 'homePhone',
                                  'Personal Contact: Personal Email': 'homeEmail',
                                  'Personal Contact: Personal Mobile': 'mobilePhone',
                                  'Work Contact: Work Email': 'workEmail', 'Work Contact: Work Phone': 'workPhone',
                                  'Gender': 'gender',
                                  'Race Description': 'ethnicity'}, inplace=True)

        # splitting reports to manager name column
        df_ee_NEW[['managerLastName', 'managerFirstName']] = df_ee_NEW.reportsToName.str.split(',', expand=True)

        # adding new columns to match the Generic Template
        df_ee_NEW['action'] = ''
        df_ee_NEW['clientId'] = ''
        df_ee_NEW['prefix'] = ''
        df_ee_NEW['accredited'] = ''
        df_ee_NEW['payrollCode'] = ''
        df_ee_NEW['paygroupDescription'] = ''
        df_ee_NEW['employeeType'] = 'Regular'
        df_ee_NEW['maritalStatus'] = ''
        df_ee_NEW['workPhoneNumberExtension'] = ''
        df_ee_NEW['annualHours'] = ''
        df_ee_NEW['ownerOfficer'] = ''
        df_ee_NEW['baseShift'] = ''
        df_ee_NEW['managerClientId'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'middleName', 'prefix', 'suffix', 'accredited', 'employeeNumber', 'addressLine1',
                        'addressLine2', 'city', 'state', 'zip', 'departmentCode', 'departmentDescription', 'payrollCode',
                        'paygroupDescription', 'employmentStatus', 'terminationDate', 'terminationReason', 'reHireDate',
                        'hireDate', 'flsa', 'statusType', 'employeeType', 'maritalStatus', 'gender', 'ethnicity',
                        'jobTitle',
                        'workPhone', 'workPhoneNumberExtension', 'workEmail', 'mobilePhone', 'homePhone', 'homeEmail',
                        'annualHours', 'ownerOfficer', 'baseShift', 'managerPriorEmployeeNumber', 'managerFirstName',
                        'managerLastName', 'managerClientId']

        df_ee_NEW = df_ee_NEW.reindex(columns=column_names)

        # saving gen EE file
        genEEFilePath = os.path.dirname(genEEFile)
        df_ee_NEW.to_excel(genEEFilePath+'\\Gen EE OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'Gen EE File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def payRates():
    root = tk.Tk()
    root.withdraw()

    payFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_payRates = pd.read_excel(payFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_payRates['Tax ID (SSN)'] = df_payRates['Tax ID (SSN)'].str.replace('-', '')

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_payRates['Regular Pay Effective Date'] = pd.to_datetime(df_payRates['Regular Pay Effective Date'].astype(str),
                                                                   errors='coerce')
        df_payRates['Additional Rates Effective Date'] = pd.to_datetime(
            df_payRates['Additional Rates Effective Date'].astype(str), errors='coerce')

        # changing the date to a date format
        df_payRates['Regular Pay Effective Date'] = pd.to_datetime(df_payRates['Regular Pay Effective Date']).dt.date
        df_payRates['Additional Rates Effective Date'] = pd.to_datetime(
            df_payRates['Additional Rates Effective Date']).dt.date

        # changes the date to the correct format
        df_payRates['Regular Pay Effective Date'] = pd.to_datetime(df_payRates['Regular Pay Effective Date'])
        df_payRates['Regular Pay Effective Date'] = df_payRates['Regular Pay Effective Date'].dt.strftime('%m/%d/%Y')

        df_payRates['Additional Rates Effective Date'] = pd.to_datetime(df_payRates['Additional Rates Effective Date'])
        df_payRates['Additional Rates Effective Date'] = df_payRates['Additional Rates Effective Date'].dt.strftime(
            '%m/%d/%Y')

        # deleting unused columns
        df_payRates_NEW = df_payRates[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Pay Frequency', 'Regular Pay Effective Date',
             'Regular Pay Rate Description', 'Regular Pay Rate Amount', 'Additional Rates Effective Date', 'Rate 2',
             'Rate 3', 'Rate 4', 'Rate 5', 'Rate 6', 'Rate 7', 'Rate 8', 'Rate 9']]

        df_payRates_NEW.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                        'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                        'Birth Date': 'birthDate', 'Regular Pay Effective Date': 'effectiveDate'},
                               inplace=True)

        # adding new columns to match the Generic Template
        df_payRates_NEW['action'] = ''
        df_payRates_NEW['clientId'] = ''
        df_payRates_NEW['sequence'] = ''
        df_payRates_NEW['payType'] = ''
        df_payRates_NEW['payRate'] = ''
        df_payRates_NEW['description'] = ''
        df_payRates_NEW['reason'] = ''
        df_payRates_NEW['employmentStatus'] = ''
        df_payRates_NEW['terminationDate'] = ''

        # Reordering the columns in the dataframe to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'effectiveDate', 'sequence', 'payType', 'payRate',
                        'description', 'reason', 'employmentStatus', 'terminationDate', 'Pay Frequency',
                        'Regular Pay Rate Description', 'Regular Pay Rate Amount', 'Additional Rates Effective Date',
                        'Rate 2', 'Rate 3', 'Rate 4', 'Rate 5', 'Rate 6', 'Rate 7', 'Rate 8', 'Rate 9']

        df_payRates_NEW = df_payRates_NEW.reindex(columns=column_names)

        # setting sequence col to '1' for primary pay rate/amount and description col to 'Rate 1'
        df_payRates_NEW['sequence'] = '1'
        df_payRates_NEW['description'] = 'Rate 1'

        # moving data to payRate col and payType col
        df_payRates_NEW['payType'] = df_payRates_NEW['Regular Pay Rate Description'].copy()
        df_payRates_NEW['payRate'] = df_payRates_NEW['Regular Pay Rate Amount'].copy()

        # changing type of payRate col
        df_payRates_NEW['payRate'] = df_payRates_NEW['payRate'].astype(float)

        # modifying the payRate decimals
        if 'Hourly' in df_payRates_NEW['payType']:
            df_payRates_NEW['payRate'].map('{payRates:,.4f}'.format(4))
        else:
            df_payRates_NEW['payRate'].round(2)

        # deleting the unused columns
        df_payRates_NEW.drop(['employmentStatus', 'terminationDate', 'Pay Frequency', 'Regular Pay Rate Description',
                              'Regular Pay Rate Amount', 'Additional Rates Effective Date'], axis=1,
                             inplace=True)

        # drop rows with NaN values in payRate
        df_payRates_NEW.dropna(subset=['payRate'], inplace=True)

        # filling in the NaN values with blank cells
        df_payRates_NEW.fillna('', inplace=True)

        # saving Pay Rates file
        payRatesFilePath = os.path.dirname(payFile)
        df_payRates_NEW.to_excel(payRatesFilePath+'\\Pay Rates OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'Pay Rates File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def fedTax():
    root = tk.Tk()
    root.withdraw()

    fedFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_taxFed = pd.read_excel(fedFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_taxFed['Tax ID (SSN)'] = df_taxFed['Tax ID (SSN)'].str.replace('-', '')

        # creating a MED dataframe - this will be the main dataframe to work from when creating all other tax code dataframes
        df_taxFed_MED = df_taxFed[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date']]

        # renaming the columns
        df_taxFed_MED.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                      'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                      'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxFed_MED['action'] = ''
        df_taxFed_MED['clientId'] = ''
        df_taxFed_MED['withholdingEffectiveStartDate'] = ''
        df_taxFed_MED['taxCode'] = 'MED'
        df_taxFed_MED['altCodes'] = ''
        df_taxFed_MED['filingStatus'] = ''
        df_taxFed_MED['exemptions'] = ''
        df_taxFed_MED['adjustWithHolding'] = ''
        df_taxFed_MED['percentage'] = ''
        df_taxFed_MED['amount'] = ''
        df_taxFed_MED['reciprocity'] = ''
        df_taxFed_MED['blockDate'] = ''
        df_taxFed_MED['calculate'] = 'True'
        df_taxFed_MED['spouseWork'] = ''
        df_taxFed_MED['additionalStateExemptions'] = ''
        df_taxFed_MED['nonResidentAlienAdditionalFithwh'] = ''
        df_taxFed_MED['ncciCode'] = ''
        df_taxFed_MED['psdCode'] = ''
        df_taxFed_MED['psdRate'] = ''
        df_taxFed_MED['blockIndicator'] = ''
        df_taxFed_MED['employmentStatus'] = ''

        df_taxFed_MEDER = df_taxFed_MED.copy()
        df_taxFed_MEDER['taxCode'] = 'MEDER'

        df_taxFed_SOC = df_taxFed_MED.copy()
        df_taxFed_SOC['taxCode'] = 'SOC'

        df_taxFed_SS = df_taxFed_MED.copy()
        df_taxFed_SS['taxCode'] = 'SOCER'

        df_taxFed_FUI = df_taxFed_MED.copy()
        df_taxFed_FUI['taxCode'] = 'FUI'

        # creating a new master dataframe
        df_taxFed_NEW = df_taxFed_MED.copy()

        # adding the new dataframes to the master dataframe
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_MEDER, ignore_index=True)
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_SOC, ignore_index=True)
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_SS, ignore_index=True)
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_FUI, ignore_index=True)

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'altCodes', 'filingStatus',
                        'exemptions', 'adjustWithHolding', 'percentage', 'amount', 'reciprocity', 'blockDate', 'calculate',
                        'spouseWork', 'additionalStateExemptions', 'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode',
                        'psdRate', 'blockIndicator', 'employmentStatus', 'terminationDate']

        df_taxFed_NEW = df_taxFed_NEW.reindex(columns=column_names)

        # pulling in original dataframe to calculate FIT tax - additional columns needed
        df_taxFed_FIT = df_taxFed[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'Do Not Calculate Federal Income Tax', 'Federal/W4 Effective Date',
             'Federal/W4 Marital Status Description', 'Federal/W4 Exemptions', 'Federal Additional Tax Amount',
             'Federal Additional Tax Amount Percentage']]

        # adding new columns to match the Generic Template
        df_taxFed_FIT['action'] = ''
        df_taxFed_FIT['clientId'] = ''
        df_taxFed_FIT['withholdingEffectiveStartDate'] = ''
        df_taxFed_FIT['taxCode'] = 'FIT'
        df_taxFed_FIT['altCodes'] = ''
        df_taxFed_FIT['filingStatus'] = ''
        df_taxFed_FIT['exemptions'] = ''
        df_taxFed_FIT['adjustWithHolding'] = ''
        df_taxFed_FIT['percentage'] = ''
        df_taxFed_FIT['amount'] = ''
        df_taxFed_FIT['reciprocity'] = ''
        df_taxFed_FIT['blockDate'] = ''
        df_taxFed_FIT['calculate'] = 'True'
        df_taxFed_FIT['spouseWork'] = ''
        df_taxFed_FIT['additionalStateExemptions'] = ''
        df_taxFed_FIT['nonResidentAlienAdditionalFithwh'] = ''
        df_taxFed_FIT['ncciCode'] = ''
        df_taxFed_FIT['psdCode'] = ''
        df_taxFed_FIT['psdRate'] = ''
        df_taxFed_FIT['blockIndicator'] = ''
        df_taxFed_FIT['employmentStatus'] = ''

        # adding data to GT columns
        df_taxFed_FIT['filingStatus'] = df_taxFed_FIT['Federal/W4 Marital Status Description']
        df_taxFed_FIT['exemptions'] = df_taxFed_FIT['Federal/W4 Exemptions']
        df_taxFed_FIT['amount'] = df_taxFed_FIT['Federal Additional Tax Amount']
        df_taxFed_FIT['withholdingEffectiveStartDate'] = df_taxFed_FIT['Federal/W4 Effective Date']

        # removing NaN values from the amount and percentage columns
        df_taxFed_FIT['Federal Additional Tax Amount'].fillna(0, inplace=True)
        df_taxFed_FIT['Federal Additional Tax Amount Percentage'].fillna(0, inplace=True)

        # changing data type from str to in for Amount and Percentage columns
        df_taxFed_FIT['Federal Additional Tax Amount'] = df_taxFed_FIT['Federal Additional Tax Amount'].astype(float)
        df_taxFed_FIT['Federal Additional Tax Amount Percentage'] = df_taxFed_FIT[
            'Federal Additional Tax Amount Percentage'].astype(float)

        # conditional statements for adjustWithHolding
        df_taxFed_FIT.loc[df_taxFed_FIT['Federal Additional Tax Amount'] > 0, 'adjustWithHolding'] = 'Add to withholding'
        df_taxFed_FIT.loc[
            df_taxFed_FIT['Federal Additional Tax Amount Percentage'] > 0, 'adjustWithHolding'] = 'Add to withholding'

        # renaming the columns
        df_taxFed_FIT.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                      'Last Name': 'lastName', 'First Name': 'firstName',
                                      'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate'}, inplace=True)

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'altCodes', 'filingStatus',
                        'exemptions', 'adjustWithHolding', 'percentage', 'amount', 'reciprocity', 'blockDate', 'calculate',
                        'spouseWork', 'additionalStateExemptions', 'nonResidentAlienAdditionalFitwh',
                        'ncciCode', 'psdCode', 'psdRate', 'blockIndicator', 'employmentStatus', 'terminationDate']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxFed_FIT['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxFed_FIT['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_FIT['withholdingEffectiveStartDate'])
        df_taxFed_FIT['withholdingEffectiveStartDate'] = df_taxFed_FIT['withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxFed_FIT = df_taxFed_FIT.reindex(columns=column_names)
        # adding the new dataframes to the master dataframe
        df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_FIT, ignore_index=True)

        # filling in the NaN values with blank cells
        df_taxFed_NEW.fillna('', inplace=True)

        # saving Fed Taxes file
        fedTaxFilePath = os.path.dirname(fedFile)
        df_taxFed_NEW.to_excel(fedTaxFilePath+'\\Federal Taxes OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'Fed Tax File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def stateTax():
    root = tk.Tk()
    root.withdraw()

    stateFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_taxState = pd.read_excel(stateFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_taxState['Tax ID (SSN)'] = df_taxState['Tax ID (SSN)'].str.replace('-', '')
        # creating a SUI dataframe
        df_taxState_State = df_taxState[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'State Additional Tax Amount', 'State Additional Tax Amount Percentage', 'State Effective Date',
             'State Exemptions/Allowances', 'State Marital Status Description', 'State Tax Code', 'State Tax Description']]

        # renaming the columns
        df_taxState_State.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                          'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                          'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxState_State['action'] = ''
        df_taxState_State['clientId'] = ''
        df_taxState_State['withholdingEffectiveStartDate'] = ''
        df_taxState_State['taxCode'] = ''
        df_taxState_State['filingStatus'] = ''
        df_taxState_State['exemptions'] = ''
        df_taxState_State['adjustWithHolding'] = ''
        df_taxState_State['percentage'] = ''
        df_taxState_State['amount'] = ''
        df_taxState_State['reciprocity'] = ''
        df_taxState_State['blockDate'] = ''
        df_taxState_State['calculate'] = 'True'
        df_taxState_State['spouseWork'] = ''
        df_taxState_State['additionalStateExemptions'] = ''
        df_taxState_State['nonResidentAlienAdditionalFithwh'] = ''
        df_taxState_State['ncciCode'] = ''
        df_taxState_State['psdCode'] = ''
        df_taxState_State['psdRate'] = ''
        df_taxState_State['blockIndicator'] = ''

        # adding data to GT columns
        df_taxState_State['taxCode'] = df_taxState_State['State Tax Code']
        df_taxState_State['calculate'] = 'True'

        df_taxState_State['amount'] = df_taxState_State['State Additional Tax Amount']
        df_taxState_State['percentage'] = df_taxState_State['State Additional Tax Amount Percentage']

        df_taxState_State['withholdingEffectiveStartDate'] = df_taxState_State['State Effective Date']

        # removing NaN values from the amount and percentage columns
        df_taxState_State['State Additional Tax Amount'].fillna(0, inplace=True)
        df_taxState_State['State Additional Tax Amount Percentage'].fillna(0, inplace=True)

        # changing data type from str to in for Amount and Percentage columns
        df_taxState_State['State Additional Tax Amount'] = df_taxState_State['State Additional Tax Amount'].astype(float)
        df_taxState_State['State Additional Tax Amount Percentage'] = df_taxState_State[
            'State Additional Tax Amount Percentage'].astype(float)

        # 201223 Adding the `filingStatus` and `exemptions` fields to state taxes.
        # Also added on line 966.
        df_taxState_State['filingStatus'] = df_taxState_State['State Marital Status Description']
        df_taxState_State['exemptions'] = df_taxState_State['State Exemptions/Allowances']

        # conditional statements for adjustWithHolding
        df_taxState_State.loc[
            df_taxState_State['State Additional Tax Amount'] > 0, 'adjustWithHolding'] = 'Add to withholding'
        df_taxState_State.loc[
            df_taxState_State['State Additional Tax Amount Percentage'] > 0, 'adjustWithHolding'] = 'Add to withholding'

        # dropping columns from the dataframe
        df_taxState_State = df_taxState_State.drop(
            columns=['State Additional Tax Amount', 'State Additional Tax Amount Percentage', 'State Effective Date',
                     'State Exemptions/Allowances', 'State Marital Status Description', 'State Tax Code',
                     'State Tax Description'])

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_State['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_State['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_State['withholdingEffectiveStartDate'])
        df_taxState_State['withholdingEffectiveStartDate'] = df_taxState_State['withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxState_State = df_taxState_State.reindex(columns=column_names)

        # filling in the NaN values with blank cells
        df_taxState_State.fillna('', inplace=True)

        # creating a new state UNE dataframe
        df_taxState_UNE = df_taxState[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'SUI/SDI Effective Date', 'SUI/SDI Tax Code', 'SUI/SDI Tax Code Description']]

        # renaming the columns
        df_taxState_UNE.rename(columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                                        'Last Name': 'lastName', 'First Name': 'firstName', 'Tax ID (SSN)': 'ssn',
                                        'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxState_UNE['action'] = ''
        df_taxState_UNE['clientId'] = ''
        df_taxState_State['withholdingEffectiveStartDate'] = ''
        df_taxState_UNE['taxCode'] = ''
        df_taxState_UNE['filingStatus'] = ''
        df_taxState_UNE['exemptions'] = ''
        df_taxState_UNE['adjustWithHolding'] = ''
        df_taxState_UNE['percentage'] = ''
        df_taxState_UNE['amount'] = ''
        df_taxState_UNE['reciprocity'] = ''
        df_taxState_UNE['blockDate'] = ''
        df_taxState_UNE['calculate'] = 'True'
        df_taxState_UNE['spouseWork'] = ''
        df_taxState_UNE['additionalStateExemptions'] = ''
        df_taxState_UNE['nonResidentAlienAdditionalFithwh'] = ''
        df_taxState_UNE['ncciCode'] = ''
        df_taxState_UNE['psdCode'] = ''
        df_taxState_UNE['psdRate'] = ''
        df_taxState_UNE['blockIndicator'] = ''

        # adding data to GT columns
        df_taxState_UNE['taxCode'] = df_taxState_UNE['SUI/SDI Tax Code']

        df_taxState_UNE['withholdingEffectiveStartDate'] = df_taxState_UNE['SUI/SDI Effective Date']

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_UNE['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxState_UNE['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxState_UNE['withholdingEffectiveStartDate'])
        df_taxState_UNE['withholdingEffectiveStartDate'] = df_taxState_UNE['withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxState_UNE = df_taxState_UNE.reindex(columns=column_names)

        # creating a new master dataframe
        df_taxState_NEW = df_taxState_State.copy()

        # adding the new dataframes to the master dataframe
        df_taxState_NEW = df_taxState_NEW.append(df_taxState_UNE, ignore_index=True)

        # filling in the NaN values with blank cells
        df_taxState_NEW.fillna('', inplace=True)

        # saving Local Taxes file
        stateTaxFilePath = os.path.dirname(stateFile)
        df_taxState_NEW.to_excel(stateTaxFilePath+'\\State Taxes OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'State Taxes File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))


def localTax():
    root = tk.Tk()
    root.withdraw()

    localFile = tk.filedialog.askopenfilename()

    try:
        # keeping leading zeros for file number when imported
        df_taxLocal = pd.read_excel(localFile, dtype=str)

        # removing dashes from SSN numbers in Tax ID column
        df_taxLocal['Tax ID (SSN)'] = df_taxLocal['Tax ID (SSN)'].str.replace('-', '')

        # getting rid of the unneeded column headers
        df_taxLocal_NEW = df_taxLocal[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'School Tax Code',
             'School Tax Description', 'Local Tax Code', 'Local Tax Description', 'Local Effective Date']]

        # renaming the columns
        df_taxLocal_NEW.rename(
            columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                     'Last Name': 'lastName',
                     'First Name': 'firstName', 'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxLocal_NEW['action'] = ''
        df_taxLocal_NEW['clientId'] = ''
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = ''
        df_taxLocal_NEW['taxCode'] = ''
        df_taxLocal_NEW['filingStatus'] = ''
        df_taxLocal_NEW['exemptions'] = ''
        df_taxLocal_NEW['adjustWithHolding'] = ''
        df_taxLocal_NEW['percentage'] = ''
        df_taxLocal_NEW['amount'] = ''
        df_taxLocal_NEW['reciprocity'] = ''
        df_taxLocal_NEW['blockDate'] = ''
        df_taxLocal_NEW['calculate'] = 'True'
        df_taxLocal_NEW['spouseWork'] = ''
        df_taxLocal_NEW['additionalStateExemptions'] = ''
        df_taxLocal_NEW['nonResidentAlienAdditionalFithwh'] = ''
        df_taxLocal_NEW['ncciCode'] = ''
        df_taxLocal_NEW['psdCode'] = ''
        df_taxLocal_NEW['psdRate'] = ''
        df_taxLocal_NEW['blockIndicator'] = ''

        # adding data to GT columns
        df_taxLocal_NEW['taxCode'] = df_taxLocal_NEW['Local Tax Code'] + ' ' + df_taxLocal_NEW['Local Tax Description']
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = df_taxLocal_NEW['Local Effective Date']

        # dropping columns from the dataframe
        df_taxLocal_NEW = df_taxLocal_NEW.drop(
            columns=['School Tax Code', 'School Tax Description', 'Local Tax Code', 'Local Tax Description',
                     'Local Effective Date'])

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_NEW['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_NEW['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxLocal_NEW['withholdingEffectiveStartDate'])
        df_taxLocal_NEW['withholdingEffectiveStartDate'] = df_taxLocal_NEW['withholdingEffectiveStartDate'].dt.strftime(
            '%m/%d/%Y')

        df_taxLocal_NEW = df_taxLocal_NEW.reindex(columns=column_names)

        # creating a dataframe for school tax if present
        # getting rid of the unneeded column headers
        df_taxLocal_School = df_taxLocal[
            ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date',
             'School Tax Code',
             'School Tax Description', 'School District Effective Date']]

        # renaming the columns
        df_taxLocal_School.rename(
            columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber',
                     'Last Name': 'lastName',
                     'First Name': 'firstName', 'Tax ID (SSN)': 'ssn', 'Birth Date': 'birthDate'}, inplace=True)

        # adding new columns to match the Generic Template
        df_taxLocal_School['action'] = ''
        df_taxLocal_School['clientId'] = ''
        df_taxLocal_School['withholdingEffectiveStartDate'] = ''
        df_taxLocal_School['taxCode'] = ''
        df_taxLocal_School['filingStatus'] = ''
        df_taxLocal_School['exemptions'] = ''
        df_taxLocal_School['adjustWithHolding'] = ''
        df_taxLocal_School['percentage'] = ''
        df_taxLocal_School['amount'] = ''
        df_taxLocal_School['reciprocity'] = ''
        df_taxLocal_School['blockDate'] = ''
        df_taxLocal_School['calculate'] = 'True'
        df_taxLocal_School['spouseWork'] = ''
        df_taxLocal_School['additionalStateExemptions'] = ''
        df_taxLocal_School['nonResidentAlienAdditionalFithwh'] = ''
        df_taxLocal_School['ncciCode'] = ''
        df_taxLocal_School['psdCode'] = ''
        df_taxLocal_School['psdRate'] = ''
        df_taxLocal_School['blockIndicator'] = ''

        # adding data to GT columns
        df_taxLocal_School['taxCode'] = df_taxLocal_School['School Tax Code'] + ' ' + df_taxLocal_School[
            'School Tax Description']
        df_taxLocal_School['withholdingEffectiveStartDate'] = df_taxLocal_School['School District Effective Date']

        # dropping columns from the dataframe
        df_taxLocal_School = df_taxLocal_School.drop(
            columns=['School Tax Code', 'School Tax Description', 'School District Effective Date'])

        # reordering the columns to match the GT
        column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                        'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions',
                        'adjustWithHolding', 'percentage', 'amount',
                        'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                        'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

        # changing the date to string and forcing anything else in the column to do so as well with coerce
        df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_School['withholdingEffectiveStartDate'].astype(str), errors='coerce')

        # changing the date to a date format
        df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_School['withholdingEffectiveStartDate']).dt.date

        # changes the date to the correct format
        df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(
            df_taxLocal_School['withholdingEffectiveStartDate'])
        df_taxLocal_School['withholdingEffectiveStartDate'] = df_taxLocal_School[
            'withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')

        df_taxLocal_School = df_taxLocal_School.reindex(columns=column_names)

        # adding the new dataframes to the master dataframe
        df_taxLocal_NEW = df_taxLocal_NEW.append(df_taxLocal_School, ignore_index=True)

        # deleting rows with NaN values in taxCode column
        df_taxLocal_NEW = df_taxLocal_NEW.dropna(subset=['taxCode'])

        # filling in the NaN values with blank cells
        df_taxLocal_NEW.fillna('', inplace=True)

        # saving Local Taxes file
        localTaxFilePath = os.path.dirname(localFile)
        df_taxLocal_NEW.to_excel(localTaxFilePath+'\\Local Taxes OUTPUT.xlsx', index=False)

        messagebox.showinfo('File', 'Local Taxes File Created')

    except Exception as x:
        messagebox.showinfo("Error", "Oops! It's not you, it's me." + str(x))

main()
# buildExcel()
