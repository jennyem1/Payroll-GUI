import pandas as pd

from datetime import datetime
import os

fedFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127160\\Federal Tax Info.xlsx'
stateFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127160\\State Tax Info.xlsx'
localFile = 'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127160\\Local Tax Info.xlsx'

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

#changing the date to string and forcing anything else in the column to do so as well with coerce
df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_FIT['withholdingEffectiveStartDate'].astype(str), errors='coerce')

#changing the date to a date format
df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_FIT['withholdingEffectiveStartDate']).dt.date

#changes the date to the correct format
df_taxFed_FIT['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxFed_FIT['withholdingEffectiveStartDate'])
df_taxFed_FIT['withholdingEffectiveStartDate'] = df_taxFed_FIT['withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')


df_taxFed_FIT = df_taxFed_FIT.reindex(columns=column_names)
# adding the new dataframes to the master dataframe
df_taxFed_NEW = df_taxFed_NEW.append(df_taxFed_FIT, ignore_index=True)

# filling in the NaN values with blank cells
df_taxFed_NEW.fillna('', inplace=True)

# return df_taxFed_NEW

###############  STATE TAXES  ###############
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

# 201223 Adding the `filingStatus` and `exemptions` fields to state taxes.
df_taxState_State['filingStatus'] = df_taxState_State['State Marital Status Description']
df_taxState_State['exemptions'] = df_taxState_State['State Exemptions/Allowances']

# removing NaN values from the amount and percentage columns
df_taxState_State['State Additional Tax Amount'].fillna(0, inplace=True)
df_taxState_State['State Additional Tax Amount Percentage'].fillna(0, inplace=True)

# changing data type from str to in for Amount and Percentage columns
df_taxState_State['State Additional Tax Amount'] = df_taxState_State['State Additional Tax Amount'].astype(float)
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
             'State Exemptions/Allowances', 'State Marital Status Description', 'State Tax Code', 'State Tax Description'])

# reordering the columns to match the GT
column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions', 'adjustWithHolding', 'percentage', 'amount',
                'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

# changing the date to string and forcing anything else in the column to do so as well with coerce
df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxState_State['withholdingEffectiveStartDate'].astype(str), errors='coerce')

# changing the date to a date format
df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxState_State['withholdingEffectiveStartDate']).dt.date

# changes the date to the correct format
df_taxState_State['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxState_State['withholdingEffectiveStartDate'])
df_taxState_State['withholdingEffectiveStartDate'] = df_taxState_State['withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')

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
                'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions', 'adjustWithHolding', 'percentage', 'amount',
                'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

# changing the date to string and forcing anything else in the column to do so as well with coerce
df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxState_UNE['withholdingEffectiveStartDate'].astype(str), errors='coerce')

# changing the date to a date format
df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxState_UNE['withholdingEffectiveStartDate']).dt.date

# changes the date to the correct format
df_taxState_UNE['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxState_UNE['withholdingEffectiveStartDate'])
df_taxState_UNE['withholdingEffectiveStartDate'] = df_taxState_UNE['withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')

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
df_taxLocal_NEW = df_taxLocal[['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date', 'School Tax Code',
     'School Tax Description', 'Local Tax Code', 'Local Tax Description', 'Local Effective Date']]

# renaming the columns
df_taxLocal_NEW.rename(
    columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber', 'Last Name': 'lastName',
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
    columns=['School Tax Code', 'School Tax Description', 'Local Tax Code', 'Local Tax Description', 'Local Effective Date'])

# reordering the columns to match the GT
column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions', 'adjustWithHolding', 'percentage', 'amount',
                'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']

# changing the date to string and forcing anything else in the column to do so as well with coerce
df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxLocal_NEW['withholdingEffectiveStartDate'].astype(str), errors='coerce')

# changing the date to a date format
df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxLocal_NEW['withholdingEffectiveStartDate']).dt.date

# changes the date to the correct format
df_taxLocal_NEW['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxLocal_NEW['withholdingEffectiveStartDate'])
df_taxLocal_NEW['withholdingEffectiveStartDate'] = df_taxLocal_NEW['withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')

df_taxLocal_NEW = df_taxLocal_NEW.reindex(columns=column_names)

# creating a dataframe for school tax if present
# getting rid of the unneeded column headers
df_taxLocal_School = df_taxLocal[
    ['Payroll Company Code', 'File Number', 'Last Name', 'First Name', 'Tax ID (SSN)', 'Birth Date', 'School Tax Code',
     'School Tax Description', 'School District Effective Date']]

# renaming the columns
df_taxLocal_School.rename(
    columns={'Payroll Company Code': 'priorCompanyCode', 'File Number': 'priorEmployeeNumber', 'Last Name': 'lastName',
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
df_taxLocal_School = df_taxLocal_School.drop(columns=['School Tax Code', 'School Tax Description', 'School District Effective Date'])

# reordering the columns to match the GT
column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                'firstName', 'withholdingEffectiveStartDate', 'taxCode', 'filingStatus', 'exemptions', 'adjustWithHolding', 'percentage', 'amount',
                'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
                'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator']


# changing the date to string and forcing anything else in the column to do so as well with coerce
df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxLocal_School['withholdingEffectiveStartDate'].astype(str), errors='coerce')

# changing the date to a date format
df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxLocal_School['withholdingEffectiveStartDate']).dt.date

# changes the date to the correct format
df_taxLocal_School['withholdingEffectiveStartDate'] = pd.to_datetime(df_taxLocal_School['withholdingEffectiveStartDate'])
df_taxLocal_School['withholdingEffectiveStartDate'] = df_taxLocal_School['withholdingEffectiveStartDate'].dt.strftime('%m/%d/%Y')

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

df_taxes.loc[df_taxes['filingStatus'] == 'Married, but withhold at higher single rate', 'filingStatus'] = 'Single'
df_taxes.loc[df_taxes['filingStatus'] == 'Married filing jointly (or Qualifying widow(er))', 'filingStatus'] = 'Married'
df_taxes.loc[df_taxes['filingStatus'] == 'Single or Married Filing Separately', 'filingStatus'] = 'Single'

df_taxes.to_excel('C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Documents\\_IMPORT_FILES\\ADP_WFN\\127160\\Taxes OUTPUT.xlsx', index=False)
