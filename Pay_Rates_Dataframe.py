import pandas as pd

from datetime import datetime
import os

payFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\Convert\\Pay Rates.xlsx'

###############  PAY RATES  ###############
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

#COMMENTING OUT PAYRATES 2-9
# # create new dataframe with Rate 2 data
# df_payRates_rate2 = df_payRates_NEW.copy()

# # adding new data to the payRates col
# df_payRates_rate2['payRate'] = df_payRates_rate2['Rate 2'].copy()
# df_payRates_rate2['sequence'] = '2'
# df_payRates_rate2['description'] = 'Rate 2'
#
# # create new dataframe with Rate 3-9 data
# df_payRates_rate3 = df_payRates_NEW.copy()
# df_payRates_rate4 = df_payRates_NEW.copy()
# df_payRates_rate5 = df_payRates_NEW.copy()
# df_payRates_rate6 = df_payRates_NEW.copy()
# df_payRates_rate7 = df_payRates_NEW.copy()
# df_payRates_rate8 = df_payRates_NEW.copy()
# df_payRates_rate9 = df_payRates_NEW.copy()

# # adding new data to the payRates col for each dataframe
# df_payRates_rate3['payRate'] = df_payRates_rate3['Rate 3'].copy()
# df_payRates_rate3['sequence'] = '3'
# df_payRates_rate3['description'] = 'Rate 3'
#
# df_payRates_rate4['payRate'] = df_payRates_rate4['Rate 4'].copy()
# df_payRates_rate4['sequence'] = '4'
# df_payRates_rate4['description'] = 'Rate 4'
#
# df_payRates_rate5['payRate'] = df_payRates_rate5['Rate 5'].copy()
# df_payRates_rate5['sequence'] = '5'
# df_payRates_rate5['description'] = 'Rate 5'
#
# df_payRates_rate6['payRate'] = df_payRates_rate6['Rate 6'].copy()
# df_payRates_rate6['sequence'] = '6'
# df_payRates_rate6['description'] = 'Rate 6'
#
# df_payRates_rate7['payRate'] = df_payRates_rate7['Rate 7'].copy()
# df_payRates_rate7['sequence'] = '7'
# df_payRates_rate7['description'] = 'Rate 7'
#
# df_payRates_rate8['payRate'] = df_payRates_rate8['Rate 8'].copy()
# df_payRates_rate8['sequence'] = '8'
# df_payRates_rate8['description'] = 'Rate 8'
#
# df_payRates_rate9['payRate'] = df_payRates_rate9['Rate 9'].copy()
# df_payRates_rate9['sequence'] = '9'
# df_payRates_rate9['description'] = 'Rate 9'

#Commenting out Payrates 2-9
# # adding the new dataframes to the master dataframe
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate2, ignore_index=True)
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate3, ignore_index=True)
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate4, ignore_index=True)
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate5, ignore_index=True)
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate6, ignore_index=True)
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate7, ignore_index=True)
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate8, ignore_index=True)
# df_payRates_NEW = df_payRates_NEW.append(df_payRates_rate9, ignore_index=True)

# deleting the unused columns
df_payRates_NEW.drop(['employmentStatus', 'terminationDate', 'Pay Frequency', 'Regular Pay Rate Description',
                      'Regular Pay Rate Amount', 'Additional Rates Effective Date'], axis=1,
                     inplace=True)

# commenting out Rates 2-9
# deleting the unused columns
# df_payRates_NEW.drop(['employmentStatus', 'terminationDate', 'Pay Frequency', 'Regular Pay Rate Description',
#                       'Regular Pay Rate Amount', 'Additional Rates Effective Date',
#                       'Rate 2', 'Rate 3', 'Rate 4', 'Rate 5', 'Rate 6', 'Rate 7', 'Rate 8', 'Rate 9'], axis=1,
#                      inplace=True)

# drop rows with NaN values in payRate
df_payRates_NEW.dropna(subset=['payRate'], inplace=True)

# filling in the NaN values with blank cells
df_payRates_NEW.fillna('', inplace=True)

df_payRates_NEW.to_excel('C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\Convert\\Pay Rates OUTPUT.xlsx', index=False)