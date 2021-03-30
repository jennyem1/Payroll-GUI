import pandas as pd

from datetime import datetime
import os

###############  EARNINGS  ###############
earnFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127160\\Scheduled Earnings.xlsx'
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
                'firstName', 'effectiveDates effectiveDate1', 'code', 'effectiveDates amount1', 'effectiveDates rate1',
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

df_earn_NEW.to_excel('C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127160\\Scheduled Earnings OUTPUT.xlsx', index=False)