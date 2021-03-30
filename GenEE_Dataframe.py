import pandas as pd

#keeping leading zeros for file number when imported
df_ee = pd.read_excel(r'C:\\Users\\jen080519\\OneDrive - Paycor, Inc\\Desktop\\In Progress ADP WFN Transform\\General EE Info.xlsx', converters={'File Number': lambda x: str(x)})

#removing dashes from SSN numbers in Tax ID column
df_ee['Tax ID (SSN)'] = df_ee['Tax ID (SSN)'].str.replace('-', '')

# #removing decimals from the Home Department Code - USE THIS CODE AFTER REMOVING ALL NAN VALUES FROM THE DATAFRAME
# df_ee['Home Department Code'] = df_ee['Home Department Code'].astype(int)
df_ee.fillna('', inplace = True)

pd.set_option('display.max_columns', None)
# pd.set_option('display.max_rows', None)

#Creating a new dataframe with only the columns needed
df_ee_NEW = df_ee[['Payroll Company Code', 'File Number', 'Tax ID (SSN)', 'Birth Date', 'Last Name', 'First Name', 'Middle Name', 'Generation Suffix Code', 'Position Status', 'Termination Date', 'Termination Reason Description', 'Home Department Code', 'Home Department Description', 'Associate ID', 'Hire Date', 'Worker Category Description', 'FLSA Description', 'Job Title Description', 'Reports To Associate ID', 'Reports To Name', 'Rehire Date', 'Legal / Preferred Address: Address Line 1', 'Legal / Preferred Address: Address Line 2', 'Legal / Preferred Address: Address Line 3', 'Legal / Preferred Address: City', 'Legal / Preferred Address: State / Territory Code', 'Legal / Preferred Address: Zip / Postal Code', 'Personal Contact: Home Phone', 'Personal Contact: Personal Email', 'Personal Contact: Personal Mobile', 'Work Contact: Work Email', 'Work Contact: Work Phone', 'Gender', 'Race Description']]

# #splitting firstName column in case some middle names shifted
# newName = df_ee_NEW['First Name'].str.split(' ', n=1, expand = True)
# #making separate first/middle name column from new data frame
# df_ee_NEW['first name'] = newName[0]
# df_ee_NEW['middleInitial'] = newName[1]
#
# df_ee_NEW.fillna('', inplace = True)
# df_ee_NEW['middleNames'] = df_ee_NEW['Middle Name'] + df_ee_NEW['middleInitial']
#
# df_ee_NEW.rename(columns={'middleNames':'middleName', 'first name':'firstName'}, inplace=True)

#changing the date to string and forcing anything else in the column to do so as well with coerce
df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date'].astype(str), errors='coerce')
df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'].astype(str), errors='coerce')
df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Rehire Date'].astype(str), errors='coerce')

#changing the date to a date format
df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date']).dt.date
df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date']).dt.date
df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Rehire Date']).dt.date

#changes the date to the correct format
df_ee_NEW['Termination Date'] = pd.to_datetime(df_ee_NEW['Termination Date'])
df_ee_NEW['Termination Date'] = df_ee_NEW['Termination Date'].dt.strftime('%m/%d/%Y')

df_ee_NEW['Hire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'])
df_ee_NEW['Hire Date'] = df_ee_NEW['Hire Date'].dt.strftime('%m/%d/%Y')

df_ee_NEW['Rehire Date'] = pd.to_datetime(df_ee_NEW['Hire Date'])
df_ee_NEW['Rehire Date'] = df_ee_NEW['Rehire Date'].dt.strftime('%m/%d/%Y')

#formatting the phone number columns
df_ee_NEW['Personal Contact: Home Phone'] = df_ee_NEW['Personal Contact: Home Phone'].str.replace('[(,),-, ]', '')
df_ee_NEW['Personal Contact: Personal Mobile'] = df_ee_NEW['Personal Contact: Personal Mobile'].str.replace('[(,),-, ]', '')
df_ee_NEW['Work Contact: Work Phone'] = df_ee_NEW['Work Contact: Work Phone'].str.replace('[(,),-, ]', '')

df_ee_NEW['Personal Contact: Personal Mobile'] = df_ee_NEW['Personal Contact: Personal Mobile'].str.replace('[-, ]', '')
df_ee_NEW['Personal Contact: Home Phone'] = df_ee_NEW['Personal Contact: Home Phone'].str.replace('[-, ]', '')
df_ee_NEW['Work Contact: Work Phone'] = df_ee_NEW['Work Contact: Work Phone'].str.replace('[-, ]', '')

# #getting rid of NaN values in the dataframe
# df_ee_NEW.fillna('', inplace = True)

#concatenating the Address Line 2 & 3
df_ee_NEW['addressLine2'] = df_ee_NEW['Legal / Preferred Address: Address Line 2'] + df_ee_NEW['Legal / Preferred Address: Address Line 3']

#changing the column names to fit the Generic Template names
df_ee_NEW.rename(columns={'Payroll Company Code':'priorCompanyCode', 'File Number':'priorEmployeeNumber',
                          'Tax ID (SSN)':'ssn', 'Birth Date':'birthDate', 'Last Name':'lastName', 'Generation Suffix Code':'suffix',
                          'Position Status':'employmentStatus', 'Termination Date':'terminationDate', 'Middle Name': 'middleName',
                          'Termination Reason Description':'terminationReason', 'Home Department Code':'departmentCode',
                          'Home Department Description':'departmentDescription', 'Associate ID':'employeeNumber',
                          'Hire Date':'hireDate', 'Worker Category Description':'statusType', 'FLSA Description':'flsa',
                          'Job Title Description':'jobTitle', 'Reports To Associate ID':'managerPriorEmployeeNumber',
                          'Reports To Name':'reportsToName', 'Rehire Date':'rehireDate', 'Legal / Preferred Address: Address Line 1':'addressLine1',
                          'Legal / Preferred Address: City':'city', 'Legal / Preferred Address: State / Territory Code':'state',
                          'Legal / Preferred Address: Zip / Postal Code':'zip', 'Personal Contact: Home Phone':'homePhone',
                          'Personal Contact: Personal Email':'homeEmail', 'Personal Contact: Personal Mobile':'mobilePhone',
                          'Work Contact: Work Email':'workEmail', 'Work Contact: Work Phone':'workPhone', 'Gender':'gender',
                          'Race Description':'ethnicity'}, inplace=True)

#splitting reports to manager name column
df_ee_NEW[['managerLastName', 'managerFirstName']] = df_ee_NEW.reportsToName.str.split(',', expand = True)

#adding new columns to match the Generic Template
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

#Reordering the columns in the dataframe to match the GT
column_names = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                'firstName', 'middleName', 'prefix', 'suffix', 'accredited', 'employeeNumber', 'addressLine1',
                'addressLine2', 'city', 'state', 'zip', 'departmentCode', 'departmentDescription', 'payrollCode',
                'paygroupDescription', 'employmentStatus', 'terminationDate', 'terminationReason', 'reHireDate',
                'hireDate', 'flsa', 'statusType', 'employeeType', 'maritalStatus', 'gender', 'ethnicity', 'jobTitle',
                'workPhone', 'workPhoneNumberExtension', 'workEmail', 'mobilePhone', 'homePhone', 'homeEmail',
                'annualHours', 'ownerOfficer', 'baseShift', 'managerPriorEmployeeNumber', 'managerFirstName',
                'managerLastName', 'managerClientId']

df_ee_NEW = df_ee_NEW.reindex(columns = column_names)


