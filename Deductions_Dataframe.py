import pandas as pd

from datetime import datetime
import os

###############  DEDUCTIONS  ###############
deductFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\WFN Test Files\\Set 3\\Scheduled Deductions_08_18_2020_02_06_15_PM.xlsx'

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
column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName', 'firstName',
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

# adding partials from DD dataframe to deductions

df_deduct_NEW.to_excel('C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\WFN Test Files\\Set 3\\Scheduled Deductions OUTPUT.xlsx', index=False)

