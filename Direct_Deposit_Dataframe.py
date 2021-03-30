import pandas as pd

from datetime import datetime
import os

###############  DIRECT DEPOSIT  ###############
depositFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\WFN Test Files\\Set 3\\Direct Deposit_08_18_2020_02_12_17_PM.xlsx'
# keeping leading zeros for file number when imported
df_directDep = pd.read_excel(depositFile, dtype=str)

# removing dashes from SSN numbers in Tax ID column
df_directDep['Tax ID (SSN)'] = df_directDep['Tax ID (SSN)'].str.replace('-', '')

df_directDep_NEW = df_directDep[['Payroll Company Code','File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date', 'Full Net', 'Routing Number', 'Account Number', 'Deduction Code [Direct Deposit]', 'Deduction Description', 'Deduction Percent', 'Amount', 'Status', 'Effective Date']]

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
                                 'Account Number': 'accountNumber', 'Deduction Code [Direct Deposit]': 'deductionCode',
                                 'Deduction Description': 'accountDescription', 'Deduction Percent': 'rate',
                                 'Amount': 'amount'}, inplace=True)

#adding new columns to match the Generic Template
df_directDep_NEW['action'] = ''
df_directDep_NEW['clientId'] = ''
df_directDep_NEW['calculate'] = ''
df_directDep_NEW['directDepositType'] = ''
df_directDep_NEW['frequency'] = ''
df_directDep_NEW['employmentStatus'] = ''
df_directDep_NEW['terminationDate'] = ''
df_directDep_NEW['accountType'] = ''

#Reordering the columns in the dataframe to match the GT
column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                'firstName', 'deductionCode','routingNumber', 'accountNumber', 'accountType', 'accountDescription',
                'rate', 'amount', 'calculate', 'directDepositType', 'frequency', 'Full Net', 'Status', 'Effective Date']

df_directDep_NEW = df_directDep_NEW.reindex(columns = column_names)

#copying the contents of Full Net column to directDeposityType column
df_directDep_NEW['directDepositType'] = df_directDep_NEW['Full Net'].copy()

#Changing the values to GT standard values
df_directDep_NEW['directDepositType'] = df_directDep_NEW['directDepositType'].replace('Yes','Net')
df_directDep_NEW['directDepositType'] = df_directDep_NEW['directDepositType'].replace('No','Partial')

#adding 'Every pay period' to frequency column
df_directDep_NEW['frequency'] = 'Every pay period'

#Taking partials and calling them out in accountType col - per Kim's advice.
# Since not all files will have "Checking" or "Savings" in the deduction description. Humans will have to do logic
# to fill in partials and place them on the Deductions tab for now.
df_directDep_NEW['accountType'] = df_directDep_NEW['directDepositType'].copy()

#deleting Inactive direct deposits
df_directDep_NEW = df_directDep_NEW[~df_directDep_NEW['Status'].isin(['Inactive'])]

#deleting extra columns to match GT
df_directDep_NEW.drop(columns=['directDepositType', 'Full Net', 'Status', 'accountDescription'])

# filling in the NaN values with blank cells
df_directDep_NEW.fillna('', inplace=True)

##########  DIRECT DEPOSIT PARTIALS TO DEDUCTIONS  ##########
# making a dataframe for partial Direct Deposits to add to Deductions
df_directDep_Partial = df_directDep_NEW.copy()

partials = df_directDep_Partial[df_directDep_Partial['accountType']=='Net'].index
df_directDep_Partial.drop(partials, inplace=True)

#Creating a new dataframe with only the columns needed
df_directDep_Partial_NEW = df_directDep_Partial[['priorCompanyCode', 'priorEmployeeNumber', 'lastName', 'firstName',
                                                 'ssn', 'birthDate', 'Effective Date', 'deductionCode', 'rate', 'amount']]

df_directDep_Partial_NEW.rename(columns={'Effective Date': 'effectiveDates effectiveDate1', 'deductionCode': 'code',
                                         'rate': 'effectiveDates rate', 'amount': 'effectiveDates amount'}, inplace=True)

#adding new columns to match the Generic Template
df_directDep_Partial_NEW['calculate'] = 'True'
df_directDep_Partial_NEW['limits maxAmount1'] = ''
df_directDep_Partial_NEW['frequency'] = ''
df_directDep_Partial_NEW['deductionType'] = ''


#Reordering the columns in the dataframe to match the GT
column_names = ['priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName', 'firstName',
                'effectiveDates effectiveDate1', 'code', 'effectiveDates rate', 'effectiveDates amount', 'calculate',
                'limits maxAmount1', 'frequency', 'deductionType']

df_directDep_Partial_NEW = df_directDep_Partial_NEW.reindex(columns = column_names)


df_directDep_NEW.to_excel('C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\WFN Test Files\\Set 3\\Direct Deposit OUTPUT.xlsx', index=False)